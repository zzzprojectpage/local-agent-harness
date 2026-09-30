import tempfile
import gc
import time
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import patch

from local_harness.app import HarnessApp
from local_harness.ollama import OllamaError


class DesktopModelAdapter:
    requests = []

    def __init__(self, url):
        self.url = url

    def list_models(self):
        return [{"name": "qwen3:0.6b", "size": 522653767}]

    def show_model(self, model, cancelled=None):
        if model.startswith("mimo"):
            raise OllamaError("Model not installed: " + model)
        return {"capabilities": ["tools"]}

    def chat(self, model, messages, **kwargs):
        self.requests.append((model, messages))
        content = "Selected model: " + model
        kwargs["on_chunk"]({"content": content})
        return {"role": "assistant", "content": content, "tool_calls": [], "stats": {}}

    def cancel(self):
        pass


class DesktopWriteAdapter(DesktopModelAdapter):
    def __init__(self, url):
        super().__init__(url)
        self.turn = 0

    def chat(self, model, messages, **kwargs):
        self.turn += 1
        if self.turn == 1:
            return {"role": "assistant", "content": "", "stats": {}, "tool_calls": [
                {"function": {"name": "write_file", "arguments": {"path": "notes.txt", "content": "proposed\n"}}}
            ]}
        content = "Tool result received: " + messages[-1]["content"]
        kwargs["on_chunk"]({"content": content})
        return {"role": "assistant", "content": content, "tool_calls": [], "stats": {}}


class DesktopTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = tk.Tk()
        self.root.withdraw()
        DesktopModelAdapter.requests = []
        self.app = HarnessApp(self.root, data_dir=Path(self.temp.name) / "data", client_factory=DesktopModelAdapter)
        self.pump(lambda: self.app.connection_text.get().startswith("Connected"))

    def tearDown(self):
        self.app.close()
        self.app = None
        self.root = None
        gc.collect()  # Finalize retired Tk objects on the GUI thread, not a future worker's GC.
        self.temp.cleanup()

    def pump(self, condition, timeout=3):
        deadline = time.monotonic() + timeout
        while not condition() and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertTrue(condition(), "Desktop action did not complete.")

    def test_select_qwen_and_send_from_the_desktop(self):
        self.app.model_box.set("qwen3:0.6b")
        self.app.model_box.event_generate("<<ComboboxSelected>>")
        self.app.prompt.insert("1.0", "Hello")
        self.app.send_button.invoke()
        self.pump(lambda: not self.app.busy)
        self.assertEqual(DesktopModelAdapter.requests[-1][0], "qwen3:0.6b")
        self.assertIn("Selected model: qwen3:0.6b", self.app.transcript.get("1.0", "end"))

    def test_missing_mimo_can_be_selected_but_cannot_silently_use_qwen(self):
        self.assertIn("mimo-2.6-low-fast", self.app.model_box["values"])
        self.app.model_box.set("mimo-2.6-low-fast")
        self.app.model_box.event_generate("<<ComboboxSelected>>")
        self.app.prompt.insert("1.0", "Hello")
        self.app.send_button.invoke()
        self.pump(lambda: not self.app.busy)
        self.assertEqual(DesktopModelAdapter.requests, [])
        self.assertIn("Model not installed: mimo-2.6-low-fast", self.app.transcript.get("1.0", "end"))

    def test_folder_picker_and_attach_button_supply_text_to_the_model(self):
        root = Path(self.temp.name)
        file = root / "notes.txt"
        file.write_text("A fact from the attachment", encoding="utf-8")
        with patch("local_harness.app.filedialog.askdirectory", return_value=str(root)):
            self.app.folder_button.invoke()
        self.assertIn(str(root.resolve()), self.app.folder_text.get())
        with patch("local_harness.app.filedialog.askopenfilenames", return_value=(str(file),)):
            self.app.attach_button.invoke()
        self.assertIn("notes.txt", self.app.attachments_list.get(0))
        self.app.prompt.insert("1.0", "Read my attachment")
        self.app.send_button.invoke()
        self.pump(lambda: not self.app.busy)
        messages = DesktopModelAdapter.requests[-1][1]
        self.assertIn("A fact from the attachment", messages[-1]["content"])

    def write_from_ui(self, allow):
        root = Path(self.temp.name)
        target = root / "notes.txt"
        target.write_text("original\n", encoding="utf-8")
        self.app.client_factory = DesktopWriteAdapter
        with patch("local_harness.app.filedialog.askdirectory", return_value=str(root)):
            self.app.folder_button.invoke()
        choice = "Allow this write" if allow else "Deny"
        clicked = []
        previews = []

        def choose():
            def widgets(parent):
                for child in parent.winfo_children():
                    yield child
                    yield from widgets(child)

            for widget in widgets(self.root):
                if isinstance(widget, tk.Text):
                    text = widget.get("1.0", "end")
                    if "-original" in text:
                        previews.append(text)
                if widget.winfo_class() == "TButton" and widget.cget("text") == choice:
                    clicked.append(choice)
                    widget.invoke()
                    return
            self.root.after(10, choose)

        self.root.after(10, choose)
        self.app.prompt.insert("1.0", "Replace notes.txt")
        self.app.send_button.invoke()
        self.pump(lambda: not self.app.busy)
        self.assertEqual(clicked, [choice])
        self.assertTrue(previews, "The user must see the actual diff before approving.")
        return root, target

    def test_deny_button_leaves_original_file_unchanged(self):
        root, target = self.write_from_ui(False)
        self.assertEqual(target.read_text(encoding="utf-8"), "original\n")
        self.assertIn('"status": "denied"', self.app.transcript.get("1.0", "end"))

    def test_allow_button_writes_and_backs_up_original(self):
        root, target = self.write_from_ui(True)
        self.assertEqual(target.read_text(encoding="utf-8"), "proposed\n")
        backup = list((root / ".local-harness-backups").glob("*/notes.txt"))
        self.assertEqual(len(backup), 1)
        self.assertEqual(backup[0].read_text(encoding="utf-8"), "original\n")

    def test_settings_only_offers_response_limits_that_fit_the_context(self):
        self.app.settings_button.invoke()

        def widgets(parent):
            for child in parent.winfo_children():
                yield child
                yield from widgets(child)

        combos = [widget for widget in widgets(self.root) if widget.winfo_class() == "TCombobox"]
        context = next(widget for widget in combos if "32768" in tuple(str(v) for v in widget["values"]))
        response = next(widget for widget in combos if "128" in tuple(str(v) for v in widget["values"]))
        context.set("8192")
        context.event_generate("<<ComboboxSelected>>")
        response.set("2048")
        context.set("2048")
        context.event_generate("<<ComboboxSelected>>")
        self.assertEqual(tuple(str(v) for v in response["values"]), ("128", "256", "512", "768"))
        self.assertEqual(response.get(), "768")


if __name__ == "__main__":
    unittest.main()
