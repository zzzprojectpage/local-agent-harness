"""Bounded, read-only Excel adapters. Nothing here saves workbooks or executes macros."""

import importlib
import json
import math
import re
import zipfile
from contextlib import ExitStack
from datetime import date, datetime, time
from pathlib import Path

WORKBOOK_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam", ".xls", ".xlt", ".xlsb"}
MAX_WORKBOOK_BYTES = 100 * 1024 * 1024
MAX_CELLS = 400


class ExcelError(ValueError):
    pass


def is_workbook(path):
    return Path(path).suffix.lower() in WORKBOOK_SUFFIXES


def dependency(name):
    try:
        return importlib.import_module(name)
    except ImportError as exc:
        raise ExcelError(f"Excel reader '{name}' is missing. Close the app, run Setup.cmd, then open Launch.cmd.") from exc


def column_name(number):
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def bounds(cell_range, max_cells=MAX_CELLS):
    if not isinstance(cell_range, str):
        raise ExcelError("Use a cell range such as A1:H12.")
    match = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?", cell_range.strip())
    if not match:
        raise ExcelError("Use an A1 range such as A1:H12; pass the sheet name separately.")
    def column(text):
        result = 0
        for char in text.upper():
            result = result * 26 + ord(char) - 64
        return result
    c1, r1 = column(match[1]), int(match[2])
    c2, r2 = column(match[3] or match[1]), int(match[4] or match[2])
    if not (1 <= c1 <= c2 <= 16384 and 1 <= r1 <= r2 <= 1048576):
        raise ExcelError("Range is outside Excel's limits or reversed.")
    if (c2 - c1 + 1) * (r2 - r1 + 1) > max_cells:
        raise ExcelError(f"Use at most {max_cells:,} cells per call; ask for smaller ranges.")
    return c1, r1, c2, r2


def value_text(value):
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if not isinstance(value, str):
        value = getattr(value, "text", "[unsupported cell value type]")
    return value if len(value) <= 600 else value[:600] + " [cell text truncated]"


def validate(path):
    path = Path(path)
    if not is_workbook(path):
        raise ExcelError("Supported: XLSX, XLSM, XLSB, XLS, XLTX, XLTM, XLT and XLAM.")
    if not path.is_file():
        raise ExcelError("Workbook file does not exist.")
    if path.stat().st_size > MAX_WORKBOOK_BYTES:
        raise ExcelError("Workbook exceeds the 100 MB safety limit. Use a smaller copy.")
    if path.suffix.lower() not in {".xls", ".xlt"}:
        if not zipfile.is_zipfile(path):
            raise ExcelError("Workbook is encrypted, damaged, or not the stated format. If encrypted, decrypt a copy in Excel first.")
        with zipfile.ZipFile(path) as archive:
            items = archive.infolist()
            if len(items) > 20000 or sum(i.file_size for i in items) > 256 * 1024 * 1024:
                raise ExcelError("Expanded workbook exceeds the archive safety limit.")
            if any(i.flag_bits & 1 for i in items):
                raise ExcelError("Encrypted workbook: decrypt a copy in Excel first.")
    return path


def select_sheet(names, requested):
    if not names:
        raise ExcelError("This workbook/add-in contains no readable worksheets. Macros are not executed.")
    if requested is None:
        return names[0]
    if not isinstance(requested, str) or requested not in names:
        raise ExcelError("Sheet name does not exist. Use inspect_workbook to see the actual sheet names.")
    return requested


def open_xml(path, sheet=None, cell_range=None):
    module = dependency("openpyxl")
    with ExitStack() as stack:
        stream = stack.enter_context(path.open("rb"))
        book = module.load_workbook(stream, read_only=True, data_only=False, keep_links=False)
        stack.callback(book.close)
        names = [ws.title for ws in book.worksheets]
        if cell_range is None:
            with zipfile.ZipFile(path) as archive:
                has_vba = any(n.lower().endswith("vbaproject.bin") for n in archive.namelist())
            return {"format": path.suffix.lower()[1:], "read_only": True, "macros_executed": False,
                    "has_vba": has_vba, "sheet_count": len(names),
                    "sheets": [{"name": ws.title, "rows": ws.max_row, "columns": ws.max_column,
                                "state": ws.sheet_state} for ws in book.worksheets[:100]],
                    "note": "Formula values are cached, not calculated. Charts, VBA and external data are not executed."}
        name = select_sheet(names, sheet)
        c1, r1, c2, r2 = bounds(cell_range)
        cache_stream = stack.enter_context(path.open("rb"))
        cached = module.load_workbook(cache_stream, read_only=True, data_only=True, keep_links=False)
        stack.callback(cached.close)
        kwargs = dict(min_row=r1, max_row=r2, min_col=c1, max_col=c2)
        values, formulas = [], {}
        for row_number, (formula_row, value_row) in enumerate(zip(book[name].iter_rows(**kwargs), cached[name].iter_rows(**kwargs)), r1):
            values.append([value_text(cell.value) for cell in value_row])
            for col_number, cell in enumerate(formula_row, c1):
                if getattr(cell, "data_type", None) == "f":
                    formulas[f"{column_name(col_number)}{row_number}"] = value_text(cell.value)
        while len(values) < r2 - r1 + 1:
            values.append([None] * (c2 - c1 + 1))
        return {"sheet": name, "range": cell_range.upper(), "values": values, "formulas": formulas,
                "note": "Values use Excel's saved cache; null for an uncached formula. Formulas are never executed."}


def binary_xls(path, sheet=None, cell_range=None):
    module = dependency("xlrd")
    book = module.open_workbook(str(path), on_demand=True)
    try:
        names = book.sheet_names()
        if cell_range is None:
            sheets = []
            for index, name in enumerate(names[:100]):
                ws = book.sheet_by_index(index)
                sheets.append({"name": name, "rows": ws.nrows, "columns": ws.ncols,
                               "state": "hidden" if ws.visibility else "visible"})
                book.unload_sheet(index)
            return {"format": path.suffix.lower()[1:], "read_only": True, "macros_executed": False,
                    "sheet_count": len(names), "sheets": sheets,
                    "note": "XLS reader returns saved values; formula source and VBA are not read/executed."}
        name = select_sheet(names, sheet)
        ws = book.sheet_by_name(name)
        c1, r1, c2, r2 = bounds(cell_range)
        values = []
        for r in range(r1 - 1, r2):
            row = []
            for c in range(c1 - 1, c2):
                cell = ws.cell(r, c) if r < ws.nrows and c < ws.ncols else None
                value = cell.value if cell else None
                if cell:
                    if cell.ctype in (module.XL_CELL_EMPTY, module.XL_CELL_BLANK):
                        value = None
                    elif cell.ctype == module.XL_CELL_DATE:
                        value = module.xldate_as_datetime(value, book.datemode)
                    elif cell.ctype == module.XL_CELL_BOOLEAN:
                        value = bool(value)
                    elif cell.ctype == module.XL_CELL_ERROR:
                        value = module.error_text_from_code.get(value, "#ERROR!")
                row.append(value_text(value))
            values.append(row)
        return {"sheet": name, "range": cell_range.upper(), "values": values, "formulas": {},
                "note": "XLS saved values; formula source is unavailable with this reader."}
    finally:
        book.release_resources()


def binary_xlsb(path, sheet=None, cell_range=None):
    module = dependency("pyxlsb")
    with module.open_workbook(str(path)) as book:
        names = book.sheets
        if cell_range is None:
            sheets = []
            for name in names[:100]:
                with book.get_sheet(name) as ws:
                    dim = ws.dimension
                    sheets.append({"name": name, "rows": dim.r + dim.h if dim else 0,
                                   "columns": dim.c + dim.w if dim else 0})
            return {"format": "xlsb", "read_only": True, "macros_executed": False,
                    "sheet_count": len(names), "sheets": sheets,
                    "note": "XLSB saved values; dates may be Excel serial numbers. VBA is not executed."}
        name = select_sheet(names, sheet)
        c1, r1, c2, r2 = bounds(cell_range)
        values = [[None] * (c2 - c1 + 1) for _ in range(r2 - r1 + 1)]
        with book.get_sheet(name) as ws:
            for row in ws.rows(sparse=True):
                if row and row[0].r >= r2:
                    break
                for cell in row:
                    if r1 - 1 <= cell.r < r2 and c1 - 1 <= cell.c < c2:
                        values[cell.r - r1 + 1][cell.c - c1 + 1] = value_text(cell.v)
        return {"sheet": name, "range": cell_range.upper(), "values": values, "formulas": {},
                "note": "XLSB saved values; formula source unavailable; dates may be Excel serial numbers."}


def reader(path):
    if path.suffix.lower() in {".xls", ".xlt"}:
        return binary_xls
    return binary_xlsb if path.suffix.lower() == ".xlsb" else open_xml


def inspect_workbook(path):
    path = validate(path)
    try:
        return reader(path)(path)
    except ExcelError:
        raise
    except Exception as exc:
        raise ExcelError(f"Cannot inspect workbook ({type(exc).__name__}). Check its format or use a repaired copy.") from exc


def read_excel(path, sheet=None, cell_range="A1:H12"):
    path = validate(path)
    bounds(cell_range)
    try:
        return reader(path)(path, sheet, cell_range)
    except ExcelError:
        raise
    except Exception as exc:
        raise ExcelError(f"Cannot read workbook ({type(exc).__name__}). Check its format or use a repaired copy.") from exc


def workbook_preview(path):
    metadata = inspect_workbook(path)
    preview = {"workbook": metadata}
    if metadata["sheets"]:
        preview["preview"] = read_excel(path, metadata["sheets"][0]["name"], "A1:F8")
    preview["instruction"] = "This is a read-only preview. Use read_excel with the actual sheet name and an A1 range for more cells."
    return json.dumps(preview, ensure_ascii=False)
