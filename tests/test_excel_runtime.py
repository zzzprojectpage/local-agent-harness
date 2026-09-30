"""Owned-process and cancellation acceptance checks using only synthetic workbooks."""

import ctypes
import os
import tempfile
import threading
import time
import unittest
from ctypes import wintypes
from pathlib import Path

from local_harness.workspace import Workspace


def excel_pids():
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
        raise ctypes.WinError(ctypes.get_last_error())
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


@unittest.skipUnless(os.name == "nt", "Native Excel requires Windows")
class ExcelRuntimeTests(unittest.TestCase):
    def test_stop_kills_only_owned_excel_and_does_not_publish_a_partial_macro(self):
        import pythoncom
        import win32com.client
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        user = win32com.client.DispatchEx("Excel.Application")
        user.DisplayAlerts = False
        user.EnableEvents = False
        user.AutomationSecurity = 3
        visible_book = user.Workbooks.Add(-4167)
        visible_book.Worksheets(1).Range("A1").Value = 77
        before = excel_pids()
        try:
            with tempfile.TemporaryDirectory() as temp:
                root = Path(temp)
                workspace = Workspace(root, autonomous=True, allow_vba=True)
                workspace.create_workbook("source.xlsm")
                workspace.create_vba_module("source.xlsm", "Waiting", "Public Sub WaitForever()\nDo\nLoop\nEnd Sub")
                original = (root / "source.xlsm").read_bytes()
                cancelled = threading.Event()
                timer = threading.Timer(3, cancelled.set)
                started = time.monotonic()
                timer.start()
                try:
                    with self.assertRaisesRegex(ValueError, "[Ss]topped"):
                        workspace.run_vba_macro("source.xlsm", "Waiting", "WaitForever", cancelled)
                finally:
                    timer.cancel()
                    timer.join()
                self.assertLess(time.monotonic() - started, 15)
                self.assertEqual((root / "source.xlsm").read_bytes(), original)
                self.assertEqual(list(root.glob(".harness-stage-*")), [])
            self.assertEqual(user.Workbooks.Count, 1)
            self.assertEqual(visible_book.Worksheets(1).Range("A1").Value, 77)
            deadline = time.monotonic() + 8
            after = excel_pids()
            while after - before and time.monotonic() < deadline:
                time.sleep(0.1)
                after = excel_pids()
            self.assertFalse(after - before, "An automation-owned EXCEL.EXE remained after Stop.")
        finally:
            try:
                visible_book.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                user.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
