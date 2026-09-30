import tempfile
import os
import subprocess
import unittest
from pathlib import Path

from local_harness.workspace import Workspace, WorkspaceError


class WorkspaceTests(unittest.TestCase):
    def test_read_is_numbered_and_cannot_escape_the_selected_folder(self):
        with tempfile.TemporaryDirectory() as temp:
            parent = Path(temp)
            root = parent / "project"
            root.mkdir()
            (root / "notes.txt").write_text("apples\npears\n", encoding="utf-8")
            (parent / "private.txt").write_text("outside", encoding="utf-8")
            workspace = Workspace(root)
            self.assertEqual(workspace.read_file("notes.txt"), "1: apples\n2: pears")
            with self.assertRaises(WorkspaceError):
                workspace.read_file("../private.txt")
            with self.assertRaises(WorkspaceError):
                workspace.read_file(str(parent / "private.txt"))

    def test_search_skips_secret_and_binary_files_and_reads_unicode(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "report.csv").write_text("name,total\nDiana,42\n", encoding="utf-8")
            (root / ".env").write_text("name=hidden", encoding="utf-8")
            (root / "binary.txt").write_bytes(b"name\x00hidden")
            (root / "notes.txt").write_text("Bună ziua\n", encoding="utf-16")
            workspace = Workspace(root)
            self.assertEqual(workspace.list_files(), ["binary.txt", "notes.txt", "report.csv"])
            self.assertEqual(workspace.search_text("Diana"), [
                {"path": "report.csv", "line": 2, "text": "Diana,42"}
            ])
            self.assertEqual(workspace.read_file("notes.txt"), "1: Bună ziua")
            for path in (".env", "binary.txt"):
                with self.assertRaises(WorkspaceError):
                    workspace.read_file(path)

    def test_writes_require_approval_and_preserve_an_existing_file_in_backup(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            target = root / "notes.txt"
            target.write_text("old\n", encoding="utf-8")
            workspace = Workspace(root)
            denied = workspace.write_file("notes.txt", "new\n", approve=lambda request: False)
            self.assertEqual(denied["status"], "denied")
            self.assertEqual(target.read_text(encoding="utf-8"), "old\n")
            requests = []

            def approve(request):
                requests.append(request)
                return True

            result = workspace.write_file("notes.txt", "new\n", approve=approve)
            self.assertEqual(result["status"], "written")
            self.assertEqual(target.read_text(encoding="utf-8"), "new\n")
            self.assertEqual((root / result["backup"]).read_text(encoding="utf-8"), "old\n")
            self.assertIn("-old", requests[0].diff)
            self.assertNotIn(result["backup"], workspace.list_files())
            with self.assertRaises(WorkspaceError):
                workspace.write_file("book.xlsx", "corrupt", approve=approve)

    def test_file_changed_during_approval_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            target = root / "notes.txt"
            target.write_text("original", encoding="utf-8")

            def approve(request):
                target.write_text("human change", encoding="utf-8")
                return True

            with self.assertRaises(WorkspaceError):
                Workspace(root).write_file("notes.txt", "model change", approve=approve)
            self.assertEqual(target.read_text(encoding="utf-8"), "human change")

    def test_external_directory_links_are_blocked(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / "project"
            outside = Path(temp) / "outside"
            root.mkdir()
            outside.mkdir()
            (outside / "notes.txt").write_text("external", encoding="utf-8")
            try:
                (root / "link").symlink_to(outside, target_is_directory=True)
            except OSError:
                if os.name != "nt":
                    self.skipTest("Symlinks unavailable on this machine.")
                result = subprocess.run(["cmd", "/c", "mklink", "/J", str(root / "link"), str(outside)], capture_output=True)
                if result.returncode:
                    self.skipTest("Neither symlinks nor Windows directory junctions are available.")
            with self.assertRaises(WorkspaceError):
                Workspace(root).read_file("link/notes.txt")
            self.assertEqual(Workspace(root).list_files(), [])
            with self.assertRaises(WorkspaceError):
                Workspace(root).write_file("link/new.txt", "outside", approve=lambda request: True)


if __name__ == "__main__":
    unittest.main()
