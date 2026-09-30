"""Native Excel operations on harness working copies, never on original user files."""

import re
import ctypes
import queue
import threading
import time
from contextlib import contextmanager
from ctypes import wintypes
from functools import wraps

from .excel import ExcelError, bounds, dependency

AGGREGATIONS = {"sum": -4157, "count": -4112, "average": -4106, "max": -4136, "min": -4139}
NEW_FORMATS = {".xlsx": 51, ".xlsm": 52, ".xlsb": 50, ".xls": 56}
VBA_NAME = re.compile(r"[A-Za-z][A-Za-z0-9_]{0,63}\Z")
# Defense in depth, not a VBA sandbox. Execution also needs explicit session opt-in.
UNSAFE_VBA = re.compile(
    r"\b(?:Shell|CreateObject|GetObject|Kill|FileCopy|MkDir|RmDir|ChDir|ChDrive|SendKeys|"
    r"MsgBox|InputBox|OnTime|OnKey|ExecuteExcel4Macro|DDEInitiate|DDEExecute|SaveAs|SaveCopyAs|"
    r"FollowHyperlink|SetAttr|GetOpenFilename|GetSaveAsFilename|Workbooks|EnableEvents|AutomationSecurity)\b"
    r"|\bDeclare\b|\bOpen\s+[^\r\n]*\b(?:For|As)\b|\b(?:Line\s+Input|Input|Write|Close|Print|Put|Get)\s+#"
    r"|\bApplication\s*\.\s*Run\b|https?://",
    flags=re.IGNORECASE)


def excel_thread(operation):
    """COM initialization/uninitialization never touches a caller's existing apartment."""
    @wraps(operation)
    def call(*args, **kwargs):
        outcome = queue.Queue(maxsize=1)

        def run():
            try:
                outcome.put((True, operation(*args, **kwargs)))
            except Exception as exc:
                outcome.put((False, exc))

        worker = threading.Thread(target=run, name="harness-excel-sta", daemon=True)
        worker.start()
        worker.join()
        succeeded, value = outcome.get()
        if not succeeded:
            raise value
        return value
    return call


def running_excel_pids():
    """Include hidden/windowless Excel processes before claiming ownership of a PID."""
    class ProcessEntry(ctypes.Structure):
        _fields_ = [("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD),
                    ("th32ProcessID", wintypes.DWORD), ("th32DefaultHeapID", ctypes.c_size_t),
                    ("th32ModuleID", wintypes.DWORD), ("cntThreads", wintypes.DWORD),
                    ("th32ParentProcessID", wintypes.DWORD), ("pcPriClassBase", wintypes.LONG),
                    ("dwFlags", wintypes.DWORD), ("szExeFile", wintypes.WCHAR * 260)]
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.CreateToolhelp32Snapshot.argtypes = (wintypes.DWORD, wintypes.DWORD)
    kernel.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    kernel.Process32FirstW.argtypes = (wintypes.HANDLE, ctypes.POINTER(ProcessEntry))
    kernel.Process32NextW.argtypes = (wintypes.HANDLE, ctypes.POINTER(ProcessEntry))
    kernel.CloseHandle.argtypes = (wintypes.HANDLE,)
    snapshot = kernel.CreateToolhelp32Snapshot(2, 0)
    if snapshot == ctypes.c_void_p(-1).value:
        raise ExcelError("Could not establish isolated Excel process ownership. No workbook was opened.")
    entry = ProcessEntry()
    entry.dwSize = ctypes.sizeof(entry)
    result = set()
    try:
        more = kernel.Process32FirstW(snapshot, ctypes.byref(entry))
        while more:
            if entry.szExeFile.upper() == "EXCEL.EXE":
                result.add(entry.th32ProcessID)
            more = kernel.Process32NextW(snapshot, ctypes.byref(entry))
    finally:
        kernel.CloseHandle(snapshot)
    return result


@contextmanager
def native_book(path, cancelled=None, enable_macros=False, timeout=180, create=False):
    pythoncom = dependency("pythoncom")
    client = dependency("win32com.client")
    win32process = dependency("win32process")
    win32api = dependency("win32api")
    win32event = dependency("win32event")
    previous_pids = running_excel_pids()
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app = book = bootstrap = handle = None
    separators = None
    finished = threading.Event()
    interrupted = []
    owned = False
    try:
        try:
            app = client.DispatchEx("Excel.Application")
        except Exception as exc:
            raise ExcelError("Native Excel is required for editing, PivotTables and VBA. Install Microsoft Excel and run Setup.cmd.") from exc
        pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
        if pid in previous_pids:
            raise ExcelError("Excel did not create an isolated instance. No workbook was opened or changed.")
        owned = True
        handle = win32api.OpenProcess(0x100001, False, pid)  # SYNCHRONIZE + TERMINATE; newly owned process only.

        def watchdog():
            deadline = time.monotonic() + timeout
            while not finished.wait(0.2):
                if (cancelled and cancelled.is_set()) or time.monotonic() > deadline:
                    interrupted.append("stopped" if cancelled and cancelled.is_set() else "timed out")
                    try:
                        win32api.TerminateProcess(handle, 1)
                    except Exception:
                        pass
                    return

        monitor = threading.Thread(target=watchdog, daemon=True)
        monitor.start()
        # Late-bound Excel can silently drop optional arguments (e.g. Add/Move After).
        # Wrap this exact owned object, never EnsureDispatch("Excel.Application").
        app = client.gencache.EnsureDispatch(app)
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 1 if enable_macros else 3
        separators = (app.UseSystemSeparators, app.DecimalSeparator, app.ThousandsSeparator)
        # Model/tool arguments use invariant English formulas and number formats, not PC locale.
        app.UseSystemSeparators = False
        app.DecimalSeparator = "."
        app.ThousandsSeparator = ","
        # Set manual calculation before opening potentially untrusted workbooks.
        bootstrap = app.Workbooks.Add(-4167)
        app.Calculation = -4135
        if create:
            book = app.Workbooks.Add(-4167)
        else:
            book = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False, Password="",
                                      WriteResPassword="", IgnoreReadOnlyRecommended=True,
                                      Notify=False, AddToMru=False, Editable=True)
        if book.ReadOnly:
            raise ExcelError("Excel opened the work copy read-only. Check protection or close it in Excel.")
        if cancelled and cancelled.is_set():
            raise ExcelError("Excel operation stopped before editing.")
        yield book, app
        if cancelled and cancelled.is_set():
            raise ExcelError("Excel operation stopped; no pending changes were saved.")
        if create:
            book.SaveAs(str(path), FileFormat=NEW_FORMATS[path.suffix.lower()])
        else:
            book.Save()
    except ExcelError:
        raise
    except Exception as exc:
        if interrupted:
            raise ExcelError(f"Owned Excel operation {interrupted[0]}; originals remain untouched.") from exc
        raise ExcelError(f"Excel operation failed ({type(exc).__name__}). Check sheet/field names, workbook protection and Excel availability.") from exc
    finally:
        for workbook in (book, bootstrap):
            if workbook is not None and owned:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
        if app is not None and owned:
            if separators:
                try:
                    app.UseSystemSeparators = False
                    app.DecimalSeparator = separators[1]
                    app.ThousandsSeparator = separators[2]
                    app.UseSystemSeparators = separators[0]
                except Exception:
                    pass
            try:
                app.Quit()
            except Exception:
                pass
        finished.set()
        if handle is not None:
            try:
                monitor.join(timeout=2)
                if interrupted:
                    # TerminateProcess is asynchronous: wait for file locks to be released.
                    win32event.WaitForSingleObject(handle, 10000)
                win32api.CloseHandle(handle)
            except Exception:
                pass
        book = bootstrap = app = None
        pythoncom.CoUninitialize()


def worksheet(book, name):
    try:
        return book.Worksheets(name) if name is not None else book.Worksheets(1)
    except Exception as exc:
        raise ExcelError("Sheet does not exist. Inspect the workbook and use its actual sheet name.") from exc


def sheet_name(name):
    if not isinstance(name, str) or not name or len(name) > 31 or re.search(r"[\[\]:*?/\\]", name) or name.startswith("'") or name.endswith("'"):
        raise ExcelError("Use a valid Excel sheet name (1–31 characters, no []:*?/\\).")


@excel_thread
def create_workbook(path, sheet="Data", cancelled=None):
    if path.suffix.lower() not in NEW_FORMATS:
        raise ExcelError("Create a new XLSX, XLSM, XLSB or XLS workbook.")
    sheet_name(sheet)
    with native_book(path, cancelled, create=True) as (book, app):
        book.Worksheets(1).Name = sheet
        return {"sheet": sheet, "workbook_created": True}


@excel_thread
def add_sheet(path, sheet, cancelled=None):
    sheet_name(sheet)
    with native_book(path, cancelled) as (book, app):
        if book.ProtectStructure:
            raise ExcelError("Workbook structure is protected.")
        for existing in book.Sheets:
            if existing.Name.casefold() == sheet.casefold():
                raise ExcelError("Sheet already exists; choose a new name.")
        ws = book.Worksheets.Add(After=book.Sheets(book.Sheets.Count))
        ws.Name = sheet
        return {"sheet": sheet, "sheet_added": True}


def check_formula(value):
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        if ("|" in value or re.search(r"\b(?:WEBSERVICE|RTD|CALL|REGISTER|EXEC|RUN|DDE|HYPERLINK|IMAGE)\s*\(", value, re.I)
                or re.search(r"https?://|file://|\[[^\]]+\.(?:xls\w*|csv)\]", value, re.I)):
            raise ExcelError("External/network/DDE/executable formulas are not allowed.")


@excel_thread
def set_cells(path, sheet, cell_range, values, cancelled=None):
    c1, r1, c2, r2 = bounds(cell_range, max_cells=10000)
    if not isinstance(values, list) or len(values) != r2 - r1 + 1:
        raise ExcelError("values must be a two-dimensional array matching the exact target range.")
    for row in values:
        if not isinstance(row, list) or len(row) != c2 - c1 + 1:
            raise ExcelError("Every values row must match the target range width.")
        for value in row:
            if value is not None and not isinstance(value, (str, int, float, bool)):
                raise ExcelError("Cell values must be strings, numbers, booleans or null.")
            check_formula(value)
    with native_book(path, cancelled) as (book, app):
        ws = worksheet(book, sheet)
        if ws.ProtectContents:
            raise ExcelError("Target sheet is protected. Use an unprotected copy; protection is not bypassed.")
        target = ws.Range(cell_range)
        # Formula assignment otherwise coerces text IDs/dates/+strings into numbers.
        matrix = tuple(tuple("'" + value if isinstance(value, str) and not value.startswith("=") else value
                             for value in row) for row in values)
        try:
            target.Formula2 = matrix
        except Exception:
            target.Formula = matrix
        ws.Calculate()
        return {"sheet": ws.Name, "range": cell_range, "cells_changed": len(values) * len(values[0])}


@excel_thread
def fill_formula(path, sheet, cell_range, formula, cancelled=None):
    c1, r1, c2, r2 = bounds(cell_range, max_cells=2000000)
    if not isinstance(formula, str) or not formula.startswith("="):
        raise ExcelError("Provide an Excel formula beginning with =.")
    check_formula(formula)
    with native_book(path, cancelled) as (book, app):
        ws = worksheet(book, sheet)
        if ws.ProtectContents:
            raise ExcelError("Target sheet is protected.")
        target = ws.Range(cell_range)
        target.Cells(1, 1).Formula = formula
        if r2 > r1:
            target.FillDown()
        if c2 > c1:
            target.FillRight()
        ws.Calculate()
        return {"sheet": ws.Name, "range": cell_range, "formula": formula,
                "cells_changed": (r2 - r1 + 1) * (c2 - c1 + 1)}


@excel_thread
def format_cells(path, sheet, cell_range, number_format=None, bold=None, fill_color=None, cancelled=None):
    bounds(cell_range, max_cells=2000000)
    if number_format is not None and (not isinstance(number_format, str) or len(number_format) > 120):
        raise ExcelError("Use a short Excel number format.")
    if bold is not None and type(bold) is not bool:
        raise ExcelError("bold must be true or false.")
    if fill_color is not None and not re.fullmatch(r"#[0-9A-Fa-f]{6}", fill_color):
        raise ExcelError("fill_color must be a hex color such as #D1FAE5.")
    with native_book(path, cancelled) as (book, app):
        ws = worksheet(book, sheet)
        target = ws.Range(cell_range)
        if number_format is not None:
            target.NumberFormat = number_format
        if bold is not None:
            target.Font.Bold = bold
        if fill_color:
            red, green, blue = (int(fill_color[i:i + 2], 16) for i in (1, 3, 5))
            target.Interior.Color = red + green * 256 + blue * 65536
        return {"sheet": ws.Name, "range": cell_range, "formatted": True}


@excel_thread
def create_pivot(path, source_sheet, source_range, row_fields, data_fields, column_fields=None,
                 target_sheet="Pivot", target_cell="A3", name="HarnessPivot", cancelled=None):
    c1, r1, c2, r2 = bounds(source_range, max_cells=2000000)
    bounds(target_cell, max_cells=1)
    if r2 <= r1:
        raise ExcelError("Pivot source must contain a header row and at least one data row.")
    if not VBA_NAME.fullmatch(name) or not isinstance(target_sheet, str) or not target_sheet or len(target_sheet) > 31:
        raise ExcelError("Use a simple pivot name and a valid target sheet name (31 characters max).")
    sheet_name(target_sheet)
    if not isinstance(row_fields, list) or not isinstance(column_fields or [], list) or not isinstance(data_fields, list) or not data_fields:
        raise ExcelError("Provide row_fields/column_fields arrays and at least one data_fields entry.")
    with native_book(path, cancelled) as (book, app):
        source = worksheet(book, source_sheet).Range(source_range)
        raw = source.Rows(1).Value2
        headers = [str(value) if value is not None else "" for value in (raw[0] if isinstance(raw, tuple) else [raw])]
        if any(not value for value in headers) or len(set(headers)) != len(headers):
            raise ExcelError("Pivot source headers must be nonempty and unique.")
        for field in row_fields + (column_fields or []):
            if field not in headers:
                raise ExcelError("Pivot field not found in the actual source headers.")
        for spec in data_fields:
            if not isinstance(spec, dict) or spec.get("field") not in headers or spec.get("aggregation", "sum") not in AGGREGATIONS:
                raise ExcelError("Use a real source field and sum/count/average/min/max aggregation.")
        try:
            destination = book.Worksheets(target_sheet)
        except Exception:
            destination = book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
            destination.Name = target_sheet
        if destination.UsedRange.Count > 1 or destination.Range(target_cell).Value2 is not None:
            raise ExcelError("Pivot destination is not empty. Choose a new target sheet to preserve existing data.")
        escaped_sheet = source_sheet.replace("'", "''")
        address = f"'{escaped_sheet}'!R{r1}C{c1}:R{r2}C{c2}"
        cache = book.PivotCaches().Create(SourceType=1, SourceData=address)
        pivot = cache.CreatePivotTable(TableDestination=destination.Range(target_cell), TableName=name)
        for orientation, fields in ((1, row_fields), (2, column_fields or [])):
            for position, field in enumerate(fields, 1):
                pivot.PivotFields(field).Orientation = orientation
                pivot.PivotFields(field).Position = position
        for spec in data_fields:
            aggregation = spec.get("aggregation", "sum")
            caption = spec.get("name") or f"{aggregation.title()} of {spec['field']}"
            pivot.AddDataField(pivot.PivotFields(spec["field"]), caption, AGGREGATIONS[aggregation])
        pivot.RefreshTable()
        return {"pivot": name, "sheet": destination.Name, "range": pivot.TableRange2.Address, "source_range": source_range,
                "row_fields": row_fields, "column_fields": column_fields or [], "data_fields": data_fields}


@excel_thread
def convert_to_macro(source, destination, cancelled=None):
    with native_book(source, cancelled) as (book, app):
        book.SaveAs(str(destination), FileFormat=52)


@excel_thread
def create_vba_module(path, module_name, code, cancelled=None):
    if not isinstance(module_name, str) or not VBA_NAME.fullmatch(module_name):
        raise ExcelError("Use a simple VBA module name.")
    if not isinstance(code, str) or not code.strip() or len(code) > 64000:
        raise ExcelError("VBA code must be nonempty and at most 64,000 characters.")
    code = re.sub(r"(?im)^Attribute\s+[^\r\n]*(?:\r?\n|$)", "", code)
    with native_book(path, cancelled) as (book, app):
        try:
            components = book.VBProject.VBComponents
            components.Count
        except Exception as exc:
            raise ExcelError("Excel blocks VBA project access. Enable 'Trust access to the VBA project object model' once in Excel Trust Center. See VBA-SETUP.md.") from exc
        if book.FileFormat not in (52, 50, 56, 53, 17, 18, 55):
            raise ExcelError("VBA requires a macro-capable work copy (.xlsm/.xlsb/.xls/.xltm/.xlam).")
        try:
            components(module_name)
        except Exception:
            pass
        else:
            raise ExcelError("This VBA module already exists; use a new module name to preserve existing code.")
        module = components.Add(1)
        module.Name = module_name
        module.CodeModule.AddFromString(code)
        source = module.CodeModule.Lines(1, module.CodeModule.CountOfLines)
        return {"module": module_name, "vba_created": True, "macros_executed": False, "_vba_source": source}


@excel_thread
def run_vba_macro(path, module_name, macro_name, expected_code, cancelled=None):
    if not VBA_NAME.fullmatch(module_name) or not VBA_NAME.fullmatch(macro_name):
        raise ExcelError("Use simple module and macro names.")
    if UNSAFE_VBA.search(expected_code):
        raise ExcelError("VBA execution blocked: a common shell/file/network/dialog/scheduling primitive was detected. Use noninteractive Excel-only code. Screening is not a sandbox.")
    if not re.search(r"(?im)^\s*(?:Public\s+)?Sub\s+" + re.escape(macro_name) + r"\s*\(\s*\)", expected_code):
        raise ExcelError("Only a public, zero-argument Sub created by this harness session can be run.")
    with native_book(path, cancelled, enable_macros=True, timeout=60) as (book, app):
        module = book.VBProject.VBComponents(module_name).CodeModule
        actual = module.Lines(1, module.CountOfLines)
        if actual.replace("\r\n", "\n").strip() != expected_code.replace("\r\n", "\n").strip():
            raise ExcelError("VBA source changed after creation. Recreate/review the module before execution.")
        book.Activate()
        app.Run("'" + book.Name.replace("'", "''") + "'!" + module_name + "." + macro_name)
        return {"macro": module_name + "." + macro_name, "macros_executed": True}
