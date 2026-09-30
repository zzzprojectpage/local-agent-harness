"""Folder-scoped, bounded text-file tools. No shell and no delete operation."""

import fnmatch
import difflib
import os
import tempfile
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

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
    def __init__(self, root):
        self.root = Path(root).resolve(strict=True)
        if not self.root.is_dir():
            raise WorkspaceError("Choose a folder, not a file.")

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
        target = self.resolve(path)
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
                if is_sensitive(relative):
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
        target = self.resolve(path)
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
        if approve is None or not approve(request) or (cancelled and cancelled.is_set()):
            return {"status": "denied", "path": target.relative_to(self.root).as_posix()}
        if self.resolve(path) != target:
            raise WorkspaceError("The target path changed during approval. No file was written.")
        current = target.read_bytes() if target.exists() else None
        if current != original:
            raise WorkspaceError("The file changed during approval. Review it again; no file was written.")
        backup = None
        if original is not None:
            backup_root = self.root / ".local-harness-backups"
            if backup_root.is_symlink() or not backup_root.resolve().is_relative_to(self.root):
                raise WorkspaceError("Unsafe backup directory; no file was written.")
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:8]
            backup_path = backup_root / stamp / target.relative_to(self.root)
            backup_path.parent.mkdir(parents=True, exist_ok=True)
            backup_path.write_bytes(original)
            backup = backup_path.relative_to(self.root).as_posix()
        target.parent.mkdir(parents=True, exist_ok=True)
        if self.resolve(path) != target:
            raise WorkspaceError("The target path changed. No file was written.")
        staging = None
        try:
            with tempfile.NamedTemporaryFile(dir=target.parent, prefix=".harness-", delete=False) as handle:
                staging = Path(handle.name)
                handle.write(data)
            if cancelled and cancelled.is_set():
                return {"status": "denied", "path": target.relative_to(self.root).as_posix()}
            os.replace(staging, target)
        finally:
            if staging and staging.exists():
                staging.unlink()
        return {"status": "written", "path": target.relative_to(self.root).as_posix(), "backup": backup}
