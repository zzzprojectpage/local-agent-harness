"""Live end-to-end test through the native window, with optional window-only screenshot."""

import argparse
import tempfile
import time
import tkinter as tk
from pathlib import Path
from unittest.mock import patch

from local_harness.app import HarnessApp


def pump(root, condition, timeout=300):
    deadline = time.monotonic() + timeout
    while not condition():
        if time.monotonic() > deadline:
            raise AssertionError("Desktop action timed out.")
        root.update()
        time.sleep(0.02)
    root.update()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--model", default="qwen3:0.6b")
    parser.add_argument("--screenshot", type=Path)
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix="harness-desktop-") as temp:
        directory = Path(temp).resolve()
        workspace = directory / "workspace"
        workspace.mkdir()
        (workspace / "notes.txt").write_text("Verification code: HARNESS-4821\n", encoding="utf-8")
        root = tk.Tk()
        app = HarnessApp(root, data_dir=directory / "settings")
        try:
            root.geometry("1060x760+40+40")
            root.lift()
            pump(root, lambda: app.connection_text.get().startswith("Connected"), timeout=20)
            app.model_box.set(args.model)
            app.model_box.event_generate("<<ComboboxSelected>>")
            with patch("local_harness.app.filedialog.askdirectory", return_value=str(workspace)):
                app.folder_button.invoke()
            app.prompt.insert("1.0", "Use read_file to read notes.txt, then tell me its verification code.")
            app.send_button.invoke()
            pump(root, lambda: not app.busy)
            transcript = app.transcript.get("1.0", "end")
            assert "Tool: read_file" in transcript, "The desktop must execute the real file tool."
            assert "HARNESS-4821" in transcript, "The visible answer must use the actual file contents."
            assert app.status_text.get().startswith("Finished"), transcript
            print("PASS: live desktop Qwen selection + folder picker + file call + visible answer", flush=True)
            if args.screenshot:
                # Pillow is optional and verification-only, not an application dependency.
                from PIL import ImageGrab
                root.lift()
                root.update()
                args.screenshot.parent.mkdir(parents=True, exist_ok=True)
                # DWM frame bounds are physical pixels; Tk coordinates can be DPI-virtualized.
                import ctypes
                from ctypes import wintypes
                ctypes.windll.user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
                ctypes.windll.user32.GetAncestor.restype = wintypes.HWND
                hwnd = ctypes.windll.user32.GetAncestor(root.winfo_id(), 2)
                bounds = wintypes.RECT()
                result = ctypes.windll.dwmapi.DwmGetWindowAttribute(hwnd, 9, ctypes.byref(bounds), ctypes.sizeof(bounds))
                assert result == 0, "Could not obtain the window's physical screenshot bounds."
                bbox = (bounds.left, bounds.top, bounds.right, bounds.bottom)
                ImageGrab.grab(bbox=bbox).save(args.screenshot)
                print("Screenshot:", args.screenshot, flush=True)
            if "mimo-2.6-low-fast" in app.model_box["values"] and "mimo-2.6-low-fast:latest" not in app.installed_models:
                with patch("local_harness.app.messagebox.askyesno", return_value=True):
                    app.model_box.set("mimo-2.6-low-fast")
                    app.model_box.event_generate("<<ComboboxSelected>>")
                app.prompt.insert("1.0", "Say hello.")
                app.send_button.invoke()
                pump(root, lambda: not app.busy, timeout=20)
                assert "Model not installed" in app.transcript.get("1.0", "end")
                print("PASS: live desktop MiMo preset selection reports missing model without fallback", flush=True)
            root.geometry("800x620+40+40")
            root.update()
            assert app.send_button.winfo_y() >= 0 and app.send_button.winfo_viewable()
            assert app.side_canvas.yview()[1] < 1, "Sidebar should scroll at the compact laptop size."
            print("PASS: compact 800x620 layout keeps composer visible and sidebar scrollable", flush=True)
        finally:
            app.close()


if __name__ == "__main__":
    main()
