"""Folder-scoped file tools; autonomous edits are published from working copies."""

import fnmatch
import difflib
import os
import tempfile
import uuid
import hashlib
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .excel import ExcelError, bounds, column_name, inspect_workbook, is_workbook, read_excel, validate, workbook_preview

MAX_FILE_BYTES = 256 * 1024
MAX_RESULT_CHARS = 2800
TEXT_SUFFIXES = {
    ".txt", ".md", ".rst", ".csv", ".tsv", ".json", ".jsonc", ".jsonl",
    ".xml", ".html", ".htm", ".css", ".scss", ".js", ".jsx", ".ts", ".tsx",
    ".py", ".pyw", ".ps1", ".bat", ".cmd", ".sh", ".sql", ".yaml", ".yml",
    ".toml", ".ini", ".cfg", ".conf", ".log", ".c", ".h", ".cpp", ".hpp",
    ".cs", ".java", ".go", ".rs", ".rb", ".php", ".vue", ".svelte", ".bas",
    ".vbs", ".r", ".m", ".modelfile", ".gitignore", ".editorconfig",
}
SKIP_DIRS = {
    ".git", ".ssh", ".aws", ".azure", "node_modules", "__pycache__", ".venv",
    "venv", ".local-harness-backups", ".pytest_cache",
}


def is_sensitive(path):
    for part in Path(path).parts:
        name = part.lower()
        if (name in SKIP_DIRS or name.startswith(".env") or
                name in {"credentials", "credentials.json", "secrets.json", "id_rsa", "id_ed25519"} or
                name.endswith((".pem", ".key", ".pfx", ".p12"))):
            return True
    return False


def read_text_file(path):
    """Read explicitly attached or folder-scoped text; never decode a workbook as text."""
    path = Path(path)
    if is_sensitive(path.name):
        raise WorkspaceError("Sensitive filenames cannot be read or attached.")
    if not is_text_path(path):
        raise WorkspaceError("Attach text, code, Markdown or CSV files. Binary documents are not supported.")
    try:
        with path.open("rb") as handle:
            data = handle.read(MAX_FILE_BYTES + 1)
    except OSError as exc:
        raise WorkspaceError(str(exc)) from exc
    if len(data) > MAX_FILE_BYTES:
        raise WorkspaceError("Text files must be at most 256 KB; attach a smaller excerpt.")
    try:
        encoding = "utf-16" if data.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8-sig"
        text = data.decode(encoding)
    except UnicodeError as exc:
        raise WorkspaceError("This is not UTF-8/UTF-16 text. Save a text/CSV copy first.") from exc
    if "\x00" in text:
        raise WorkspaceError("Binary content cannot be read as text.")
    return text


def is_text_path(path):
    path = Path(path)
    return path.suffix.lower() in TEXT_SUFFIXES or path.name.lower() in {
        "readme", "license", "makefile", "dockerfile", "modelfile", ".gitignore", ".editorconfig"
    }


@dataclass(frozen=True)
class WriteRequest:
    path: Path
    diff: str
    is_new: bool
    byte_count: int


class WorkspaceError(ValueError):
    pass


class Workspace:
    def __init__(self, root, autonomous=False, allow_vba=False):
        self.root = Path(root).resolve(strict=True)
        if not self.root.is_dir():
            raise WorkspaceError("Choose a folder, not a file.")
        self.autonomous = autonomous
        self.allow_vba = allow_vba
        self._outputs = {}
        self._owned_hashes = {}
        self._vba_modules = {}

    def excel_target(self, path):
        target = self.resolve(path)
        output = self._outputs.get(target, target)
        if self.resolve(str(output)) != output:
            raise WorkspaceError("The working-copy path changed. Choose the folder again.")
        return output

    def resolve(self, path):
        if not isinstance(path, str) or not path or "\x00" in path:
            raise WorkspaceError("Provide a nonempty file path.")
        target = (self.root / path).resolve()
        if not target.is_relative_to(self.root):
            raise WorkspaceError("File access must stay inside the selected working folder.")
        relative = target.relative_to(self.root)
        if is_sensitive(relative) or ":" in str(relative):
            raise WorkspaceError("Sensitive paths and alternate data streams are blocked.")
        return target

    def read_file(self, path, start_line=1, max_lines=120):
        target = self.excel_target(path)
        if is_workbook(target):
            try:
                return workbook_preview(target)
            except ExcelError as exc:
                raise WorkspaceError(str(exc)) from exc
        if type(start_line) is not int or type(max_lines) is not int or start_line < 1 or max_lines < 1:
            raise WorkspaceError("Line numbers must be positive integers.")
        max_lines = min(max_lines, 200)
        lines = read_text_file(target).splitlines()
        result = "\n".join(
            f"{index + 1}: {line}"
            for index, line in enumerate(lines)
            if start_line - 1 <= index < start_line - 1 + max_lines
        )
        if len(result) > MAX_RESULT_CHARS:
            result = result[:MAX_RESULT_CHARS] + "\n[Excerpt truncated; request fewer lines.]"
        elif len(lines) > start_line - 1 + max_lines:
            result += f"\n[More lines available; next start_line: {start_line + max_lines}]"
        return result or "[No lines at this position.]"

    def inspect_workbook(self, path):
        try:
            result = inspect_workbook(self.excel_target(path))
            result["actual_path"] = self.excel_target(path).relative_to(self.root).as_posix()
            return result
        except ExcelError as exc:
            raise WorkspaceError(str(exc)) from exc

    def read_excel(self, path, sheet=None, cell_range="A1:H12"):
        try:
            return read_excel(self.excel_target(path), sheet, cell_range)
        except ExcelError as exc:
            raise WorkspaceError(str(exc)) from exc

    def _excel_edit(self, path, operation, cancelled=None, macro_format=False):
        if not self.autonomous:
            raise WorkspaceError("Enable autonomous work copies in Settings to edit Excel without per-action prompts.")
        requested = self.resolve(path)
        source = self.excel_target(path)
        validate(source)
        if cancelled and cancelled.is_set():
            raise ExcelError("Stopped before editing; no file was written.")
        original = source.read_bytes()
        digest = hashlib.sha256(original).hexdigest()
        backup = self._backup(source, original)
        suffix = ".xlsm" if macro_format and source.suffix.lower() in {".xlsx", ".xltx"} else source.suffix
        owned = digest == self._owned_hashes.get(source)
        target = source if owned and suffix.lower() == source.suffix.lower() else self._copy_name(requested, suffix)
        staging = converted = None
        try:
            with tempfile.NamedTemporaryFile(dir=source.parent, prefix=".harness-stage-", suffix=source.suffix, delete=False) as handle:
                staging = Path(handle.name)
                handle.write(original)
            if macro_format and suffix.lower() != source.suffix.lower():
                from .excel_edit import convert_to_macro
                converted = staging.with_suffix(suffix)
                convert_to_macro(staging, converted, cancelled)
            result = operation(converted or staging)
            if cancelled and cancelled.is_set():
                raise ExcelError("Operation stopped; no pending edit was published.")
            if result.get("sheet") and result.get("range"):
                c1, r1, c2, r2 = bounds(result["range"], max_cells=2000000)
                sample = f"{column_name(c1)}{r1}:{column_name(min(c2, c1 + 7))}{min(r2, r1 + 11)}"
                result["saved_preview"] = read_excel(converted or staging, result["sheet"], sample)
                result["saved_read_verified"] = True
            if self.excel_target(path) != source or source.read_bytes() != original:
                raise WorkspaceError("Workbook changed during the operation. No edit was published; retry on the current file.")
            if self.resolve(str(target)) != target or (target != source and target.exists()):
                raise WorkspaceError("Output path changed or already exists. No edit was published.")
            if cancelled and cancelled.is_set():
                raise ExcelError("Stopped after verification; no pending edit was published.")
            os.replace(converted or staging, target)
        finally:
            for temporary in (converted, staging):
                if temporary and temporary.exists():
                    temporary.unlink()
        if owned:
            for alias, output in list(self._outputs.items()):
                if output == source:
                    self._outputs[alias] = target
            for (output, module), code in list(self._vba_modules.items()):
                if output == source:
                    self._vba_modules[(target, module)] = code
        self._outputs[requested] = target
        self._owned_hashes[target] = hashlib.sha256(target.read_bytes()).hexdigest()
        return {"status": "written", "output_path": target.relative_to(self.root).as_posix(),
                "backup": backup.relative_to(self.root).as_posix(), "original_preserved": True, **result}

    def _copy_name(self, requested, suffix=None):
        return requested.with_name(requested.stem + ".harness-" + uuid.uuid4().hex[:8] + (suffix or requested.suffix))

    def _backup(self, source, content):
        backup_root = self.root / ".local-harness-backups"
        if not backup_root.resolve().is_relative_to(self.root) or backup_root.is_symlink():
            raise WorkspaceError("Unsafe backup folder. Nothing was changed.")
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:8]
        backup = backup_root / stamp / source.relative_to(self.root)
        backup.parent.mkdir(parents=True, exist_ok=True)
        backup.write_bytes(content)
        return backup

    def set_cells(self, path, sheet=None, cell_range="A1", values=None, cancelled=None):
        from .excel_edit import set_cells
        return self._excel_edit(path, lambda target: set_cells(target, sheet, cell_range, values, cancelled), cancelled)

    def create_workbook(self, path, sheet="Data", cancelled=None):
        from .excel_edit import create_workbook
        if not self.autonomous:
            raise WorkspaceError("Enable autonomous work copies in Settings to create workbooks.")
        target = self.resolve(path)
        if target.exists():
            raise WorkspaceError("Workbook already exists. Choose a new path; existing workbooks are not replaced.")
        target.parent.mkdir(parents=True, exist_ok=True)
        staging = None
        try:
            with tempfile.NamedTemporaryFile(dir=target.parent, prefix=".harness-stage-", suffix=target.suffix, delete=False) as handle:
                staging = Path(handle.name)
            result = create_workbook(staging, sheet, cancelled)
            if cancelled and cancelled.is_set():
                raise ExcelError("Stopped; no workbook was published.")
            if self.resolve(path) != target or target.exists():
                raise WorkspaceError("Output path changed or already exists; no workbook was published.")
            # On Windows, rename refuses an existing destination instead of replacing it.
            staging.rename(target)
        finally:
            if staging and staging.exists():
                staging.unlink()
        self._owned_hashes[target] = hashlib.sha256(target.read_bytes()).hexdigest()
        return {"status": "written", "output_path": target.relative_to(self.root).as_posix(), "backup": None, **result}

    def add_sheet(self, path, sheet, cancelled=None):
        from .excel_edit import add_sheet
        return self._excel_edit(path, lambda target: add_sheet(target, sheet, cancelled), cancelled)

    def fill_formula(self, path, sheet=None, cell_range="A1", formula=None, cancelled=None):
        from .excel_edit import fill_formula
        return self._excel_edit(path, lambda target: fill_formula(target, sheet, cell_range, formula, cancelled), cancelled)

    def format_cells(self, path, sheet=None, cell_range="A1", number_format=None, bold=None, fill_color=None, cancelled=None):
        from .excel_edit import format_cells
        return self._excel_edit(path, lambda target: format_cells(target, sheet, cell_range, number_format, bold, fill_color, cancelled), cancelled)

    def create_pivot(self, path, source_sheet, source_range, row_fields, data_fields, column_fields=None,
                     target_sheet="Pivot", target_cell="A3", name="HarnessPivot", cancelled=None):
        from .excel_edit import create_pivot
        return self._excel_edit(path, lambda target: create_pivot(target, source_sheet, source_range, row_fields, data_fields,
                               column_fields, target_sheet, target_cell, name, cancelled), cancelled)

    def create_vba_module(self, path, module_name, code, cancelled=None):
        from .excel_edit import create_vba_module
        result = self._excel_edit(path, lambda target: create_vba_module(target, module_name, code, cancelled), cancelled, macro_format=True)
        self._vba_modules[(self.excel_target(path), module_name)] = result.pop("_vba_source")
        return result

    def import_vba_module(self, path, module_path, module_name=None, cancelled=None):
        import re
        source = self.resolve(module_path)
        if source.suffix.lower() != ".bas":
            raise WorkspaceError("Import a .bas standard module inside the selected folder.")
        code = read_text_file(source)
        name = re.search(r'(?im)^Attribute VB_Name\s*=\s*"([A-Za-z][A-Za-z0-9_]*)"', code)
        return self.create_vba_module(path, module_name or (name[1] if name else source.stem), code, cancelled)

    def run_vba_macro(self, path, module_name, macro_name, cancelled=None):
        from .excel_edit import run_vba_macro
        if not self.allow_vba:
            raise WorkspaceError("VBA execution is disabled. Enable 'Allow VBA execution' once in Settings for this trusted folder.")
        code = self._vba_modules.get((self.excel_target(path), module_name))
        if code is None:
            raise WorkspaceError("Only VBA created/imported by this harness session may be executed. Existing workbook autorun macros are never invoked.")
        return self._excel_edit(path, lambda target: run_vba_macro(target, module_name, macro_name, code, cancelled), cancelled)

    def list_files(self, path=".", pattern="*"):
        directory = self.resolve(path)
        if not directory.is_dir():
            raise WorkspaceError("list_files requires a folder.")
        if not isinstance(pattern, str) or len(pattern) > 200:
            raise WorkspaceError("Use a short glob pattern, such as *.py.")
        result = []
        visited = 0
        for base, dirs, files in os.walk(directory, followlinks=False):
            safe_dirs = []
            for name in dirs:
                candidate = Path(base) / name
                try:
                    self.resolve(str(candidate))
                    if not candidate.is_symlink() and not is_sensitive(name):
                        safe_dirs.append(name)
                except (OSError, WorkspaceError):
                    pass
            dirs[:] = sorted(safe_dirs)
            for name in sorted(files):
                visited += 1
                if visited > 4000:
                    return sorted(result)
                candidate = Path(base) / name
                relative = candidate.relative_to(self.root).as_posix()
                if is_sensitive(relative) or name.startswith(".harness-stage-") or name.startswith("~$"):
                    continue
                try:
                    self.resolve(relative)
                except (OSError, WorkspaceError):
                    continue
                if fnmatch.fnmatch(relative, pattern) or fnmatch.fnmatch(name, pattern):
                    result.append(relative)
                    if len(result) >= 100:
                        return sorted(result)
        return sorted(result)

    def search_text(self, query, path=".", pattern="*"):
        if not isinstance(query, str) or not query or len(query) > 200:
            raise WorkspaceError("Use a nonempty search string of at most 200 characters.")
        results = []
        for name in self.list_files(path, pattern):
            try:
                text = read_text_file(self.resolve(name))
            except WorkspaceError:
                continue
            for number, line in enumerate(text.splitlines(), 1):
                if query.casefold() in line.casefold():
                    results.append({"path": name, "line": number, "text": line[:240]})
                    if len(results) >= 20:
                        return results
        return results

    def write_file(self, path, content, approve=None, cancelled=None):
        target = self.excel_target(path) if self.autonomous else self.resolve(path)
        if not is_text_path(target):
            raise WorkspaceError("Only text/code/CSV files can be written; workbooks and binary files are protected.")
        if not isinstance(content, str) or "\x00" in content:
            raise WorkspaceError("File content must be text without NUL bytes.")
        data = content.encode("utf-8")
        if len(data) > MAX_FILE_BYTES:
            raise WorkspaceError("Writes are limited to 256 KB.")
        old_text = read_text_file(target) if target.exists() else ""
        original = target.read_bytes() if target.exists() else None
        if original == data:
            return {"status": "unchanged", "path": target.relative_to(self.root).as_posix()}
        diff = "\n".join(difflib.unified_diff(
            old_text.splitlines(), content.splitlines(),
            fromfile=str(target) + " (before)", tofile=str(target) + " (after)", lineterm=""
        ))
        if not diff:
            diff = "[Only line endings or encoding differ.]"
        request = WriteRequest(target, diff, original is None, len(data))
        if (not self.autonomous and (approve is None or not approve(request))) or (cancelled and cancelled.is_set()):
            return {"status": "denied", "path": target.relative_to(self.root).as_posix()}
        if (self.excel_target(path) if self.autonomous else self.resolve(path)) != target:
            raise WorkspaceError("The target path changed during approval. No file was written.")
        current = target.read_bytes() if target.exists() else None
        if current != original:
            raise WorkspaceError("The file changed during approval. Review it again; no file was written.")
        backup = None
        if original is not None:
            backup_path = self._backup(target, original)
            backup = backup_path.relative_to(self.root).as_posix()
        source = target
        requested = self.resolve(path)
        if self.autonomous and original is not None and hashlib.sha256(original).hexdigest() != self._owned_hashes.get(target):
            target = self._copy_name(requested)
        target.parent.mkdir(parents=True, exist_ok=True)
        if self.resolve(str(target)) != target or (target != source and target.exists()):
            raise WorkspaceError("The target path changed. No file was written.")
        staging = None
        try:
            with tempfile.NamedTemporaryFile(dir=target.parent, prefix=".harness-", delete=False) as handle:
                staging = Path(handle.name)
                handle.write(data)
            if cancelled and cancelled.is_set():
                return {"status": "denied", "path": target.relative_to(self.root).as_posix()}
            if (source.read_bytes() if source.exists() else None) != original:
                raise WorkspaceError("File changed during writing. No edit was published.")
            os.replace(staging, target)
        finally:
            if staging and staging.exists():
                staging.unlink()
        if self.autonomous:
            self._outputs[requested] = target
            self._owned_hashes[target] = hashlib.sha256(data).hexdigest()
        return {"status": "written", "path": target.relative_to(self.root).as_posix(),
                "output_path": target.relative_to(self.root).as_posix(), "backup": backup}
