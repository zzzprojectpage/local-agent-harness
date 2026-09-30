"""Native Tk desktop UI. Worker threads never touch Tk widgets."""

import json
import queue
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from .agent import Agent, Attachment
from .ollama import Cancelled, OllamaClient, OllamaError
from .workspace import Workspace, WorkspaceError, is_sensitive, read_text_file

APP_DIR = Path(__file__).resolve().parent.parent
MIMO_MODELS = [f"mimo-2.6-{level}{suffix}" for level in ("low", "medium", "high") for suffix in ("-fast", "")]
DEFAULTS = {"url": "http://127.0.0.1:11434", "model": "qwen3:0.6b", "num_ctx": 4096,
            "num_predict": 768, "tool_mode": "auto", "think": False}
COLORS = {"bg": "#111519", "panel": "#1A2027", "input": "#0D1116", "text": "#F1F5F9",
          "muted": "#A3AFBF", "border": "#566273", "accent": "#A7F3D0", "error": "#FCA5A5"}


class WorkerChannel:
    """Thread-safe messages/approvals without owning any Tk objects or the UI controller."""
    def __init__(self):
        self.events = queue.Queue(maxsize=512)
        self.closed = threading.Event()

    def emit(self, kind, value):
        while not self.closed.is_set():
            try:
                self.events.put((kind, value), timeout=0.1)
                return True
            except queue.Full:
                continue
        return False

    def approve(self, request, cancelled):
        answer = {"allowed": False}
        ready = threading.Event()
        if not self.emit("approval", (request, answer, ready)):
            return False
        while not ready.wait(0.1):
            if self.closed.is_set() or cancelled.is_set():
                return False
        return answer["allowed"] and not self.closed.is_set() and not cancelled.is_set()


def load_settings(directory):
    defaults = dict(DEFAULTS)
    try:
        path = Path(directory) / "settings.json"
        if path.stat().st_size > 65536:
            return defaults
        settings = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(settings, dict):
            return defaults
        OllamaClient(settings.get("url", defaults["url"]))
        for name in ("num_ctx", "num_predict"):
            if type(settings.get(name, defaults[name])) is not int:
                return defaults
        if (settings.get("num_ctx", 4096) not in (2048, 4096, 8192, 16384, 32768) or
                settings.get("num_predict", 768) not in (128, 256, 512, 768, 1024, 2048) or
                settings.get("tool_mode", "auto") not in ("auto", "native", "json") or
                type(settings.get("think", False)) is not bool or
                not isinstance(settings.get("model", ""), str) or
                settings.get("num_predict", 768) >= settings.get("num_ctx", 4096) // 2):
            return defaults
        defaults.update({k: settings[k] for k in defaults if k in settings})
    except (OSError, ValueError, OllamaError, TypeError):
        pass
    return defaults


class HarnessApp:
    def __init__(self, root, data_dir=None, client_factory=OllamaClient):
        self.root = root
        self.data_dir = Path(data_dir) if data_dir else APP_DIR / "data"
        self.settings = load_settings(self.data_dir)
        self.client_factory = client_factory
        self.channel = WorkerChannel()
        self.events = self.channel.events
        self.installed_models = set()
        self.workspace = None
        self.attachments = []
        self.agent = None
        self.busy = False
        self.closed = False
        self.refreshing = False
        self.pending_approval = None
        self.after_ids = set()
        self.selected_model = self.settings["model"]
        self.connection_text = tk.StringVar(value="Connecting to local Ollama…")
        self.model_info = tk.StringVar(value="Checking installed models")
        self.folder_text = tk.StringVar(value="No folder selected\nAttachments-only chat")
        self.status_text = tk.StringVar(value="Ready")
        self.activity = "Ready"
        self.started = 0
        self.assistant_chars = 0
        self._build()
        self._welcome()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.bind("<Control-Return>", self.send)
        self.root.bind("<Control-o>", lambda e: self.attach_files())
        self.root.bind("<Control-Shift-O>", lambda e: self.choose_folder())
        self.root.bind("<Control-l>", lambda e: self.new_chat())
        self._later(30, self._drain)
        self._later(100, self.refresh_models)
        self._later(1000, self._tick)

    def _later(self, delay, callback):
        holder = [None]

        def run():
            self.after_ids.discard(holder[0])
            if not self.closed:
                callback()

        holder[0] = self.root.after(delay, run)
        self.after_ids.add(holder[0])

    def _build(self):
        c = COLORS
        self.root.title("Local Agent Harness")
        width = min(1060, self.root.winfo_screenwidth() - 80)
        height = min(760, self.root.winfo_screenheight() - 110)
        self.root.geometry(f"{width}x{height}")
        self.root.minsize(800, 620)
        self.root.configure(bg=c["bg"])
        self.root.option_add("*Font", "{Segoe UI} 10")
        self.root.option_add("*TCombobox*Listbox.background", c["input"])
        self.root.option_add("*TCombobox*Listbox.foreground", c["text"])
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background=c["bg"])
        style.configure("Side.TFrame", background=c["panel"])
        style.configure("TLabel", background=c["bg"], foreground=c["text"])
        style.configure("Side.TLabel", background=c["panel"], foreground=c["text"])
        style.configure("Muted.TLabel", background=c["panel"], foreground=c["muted"], font=("Segoe UI", 9))
        style.configure("Section.TLabel", background=c["panel"], foreground=c["muted"], font=("Segoe UI", 9, "bold"))
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"))
        style.configure("TButton", background="#293340", foreground=c["text"], bordercolor=c["border"], padding=(12, 8))
        style.map("TButton", background=[("active", "#38485C"), ("disabled", "#20262E")],
                  foreground=[("disabled", "#7F8B9A")])
        style.configure("Accent.TButton", background=c["accent"], foreground="#0D261C", bordercolor=c["accent"])
        style.map("Accent.TButton", background=[("active", "#D1FAE5"), ("disabled", "#334A43")],
                  foreground=[("disabled", "#A3AFBF")])
        style.configure("TCombobox", fieldbackground=c["input"], background=c["panel"], foreground=c["text"],
                        arrowcolor=c["text"], bordercolor=c["border"], padding=6)
        style.map("TCombobox", fieldbackground=[("readonly", c["input"])], foreground=[("readonly", c["text"])])
        style.configure("TEntry", fieldbackground=c["input"], foreground=c["text"], padding=6)
        style.configure("TCheckbutton", background=c["bg"], foreground=c["text"])
        self.root.rowconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)
        top = ttk.Frame(self.root, padding=(24, 18))
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)
        ttk.Label(top, text="Local Agent Harness", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(top, text="Your models. Your files. A small, native window.", foreground=c["muted"]).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.new_button = ttk.Button(top, text="New chat", command=self.new_chat)
        self.new_button.grid(row=0, column=1, rowspan=2, padx=(12, 8))
        self.settings_button = ttk.Button(top, text="Settings", command=self.show_settings)
        self.settings_button.grid(row=0, column=2, rowspan=2)
        body = ttk.Frame(self.root)
        body.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))
        body.rowconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        side_container = ttk.Frame(body, style="Side.TFrame")
        side_container.grid(row=0, column=0, sticky="ns", padx=(0, 14))
        side_container.rowconfigure(0, weight=1)
        self.side_canvas = tk.Canvas(side_container, width=268, height=1, background=c["panel"], highlightthickness=0)
        self.side_canvas.grid(row=0, column=0, sticky="nsew")
        side_scrollbar = ttk.Scrollbar(side_container, orient="vertical", command=self.side_canvas.yview)
        side_scrollbar.grid(row=0, column=1, sticky="ns")
        self.side_canvas.configure(yscrollcommand=side_scrollbar.set)
        side = ttk.Frame(self.side_canvas, style="Side.TFrame", padding=16)
        side_window = self.side_canvas.create_window((0, 0), window=side, anchor="nw", width=268)
        side.bind("<Configure>", lambda e: self.side_canvas.configure(scrollregion=self.side_canvas.bbox("all")))
        self.side_canvas.bind("<Configure>", lambda e: self.side_canvas.itemconfigure(side_window, width=e.width))
        self.root.bind("<MouseWheel>", self._side_wheel, add="+")
        side.columnconfigure(0, weight=1)
        side.rowconfigure(15, weight=1)
        ttk.Label(side, text="MODEL", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        self.model_box = ttk.Combobox(side, values=MIMO_MODELS, width=23)
        self.model_box.set(self.selected_model)
        self.model_box.grid(row=1, column=0, sticky="ew", pady=(8, 6))
        self.model_box.bind("<<ComboboxSelected>>", self.model_selected)
        self.model_box.bind("<FocusOut>", self.model_selected)
        ttk.Label(side, textvariable=self.model_info, style="Muted.TLabel", wraplength=226).grid(row=2, column=0, sticky="w")
        self.refresh_button = ttk.Button(side, text="Refresh models", command=self.refresh_models)
        self.refresh_button.grid(row=3, column=0, sticky="ew", pady=(12, 6))
        ttk.Label(side, textvariable=self.connection_text, style="Muted.TLabel", wraplength=226).grid(row=4, column=0, sticky="w")
        ttk.Label(side, text="WORKING FOLDER", style="Section.TLabel").grid(row=5, column=0, sticky="w", pady=(28, 8))
        ttk.Label(side, textvariable=self.folder_text, style="Side.TLabel", wraplength=226).grid(row=6, column=0, sticky="w")
        folder_actions = ttk.Frame(side, style="Side.TFrame")
        folder_actions.grid(row=7, column=0, sticky="ew", pady=(12, 0))
        folder_actions.columnconfigure(0, weight=1)
        self.folder_button = ttk.Button(folder_actions, text="Choose folder…", command=self.choose_folder)
        self.folder_button.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.clear_folder_button = ttk.Button(folder_actions, text="Clear", command=self.clear_folder)
        self.clear_folder_button.grid(row=0, column=1)
        ttk.Label(side, text="ATTACHMENTS", style="Section.TLabel").grid(row=8, column=0, sticky="w", pady=(28, 8))
        self.attach_button = ttk.Button(side, text="Attach files…", command=self.attach_files)
        self.attach_button.grid(row=9, column=0, sticky="ew")
        self.attachments_list = tk.Listbox(side, height=4, bg=c["input"], fg=c["text"], selectbackground="#385349",
                                          selectforeground=c["text"], highlightbackground=c["border"], borderwidth=0,
                                          activestyle="none", exportselection=False)
        self.attachments_list.grid(row=10, column=0, sticky="ew", pady=8)
        self.remove_button = ttk.Button(side, text="Remove selected", command=self.remove_attachment)
        self.remove_button.grid(row=11, column=0, sticky="ew")
        ttk.Label(side, text="Text, code, Markdown and CSV.\nExcerpts are sent with your next message.", style="Muted.TLabel", wraplength=226).grid(row=12, column=0, sticky="w", pady=(8, 0))
        ttk.Label(side, text="LOCAL BY DESIGN\nFolder-scoped reads. Approved writes.\nNo shell. No cloud fallback.", style="Muted.TLabel", wraplength=226).grid(row=16, column=0, sticky="sw", pady=(24, 0))
        main = ttk.Frame(body)
        main.grid(row=0, column=1, sticky="nsew")
        main.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)
        self.transcript = ScrolledText(main, wrap="word", state="disabled", bg=c["bg"], fg=c["text"],
                                      borderwidth=0, highlightthickness=1, highlightbackground=c["border"],
                                      padx=20, pady=20, font=("Consolas", 10), spacing1=2, spacing3=3)
        self.transcript.grid(row=0, column=0, sticky="nsew")
        self.transcript.tag_configure("role", foreground=c["accent"], font=("Segoe UI", 11, "bold"), spacing1=12)
        self.transcript.tag_configure("muted", foreground=c["muted"])
        self.transcript.tag_configure("error", foreground=c["error"])
        self.transcript.tag_configure("tool", foreground="#A5C9F7", font=("Consolas", 9))
        ttk.Label(main, text="MESSAGE", foreground=c["muted"], font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", pady=(16, 8))
        self.prompt = tk.Text(main, height=4, wrap="word", bg=c["input"], fg=c["text"], insertbackground=c["accent"],
                              highlightthickness=1, highlightbackground=c["border"], highlightcolor=c["accent"],
                              borderwidth=0, padx=12, pady=10, font=("Segoe UI", 11), undo=True)
        self.prompt.grid(row=2, column=0, sticky="ew")
        actions = ttk.Frame(main)
        actions.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        actions.columnconfigure(0, weight=1)
        ttk.Label(actions, text="Ctrl+Enter to send · Enter for a new line", foreground=c["muted"], font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w")
        self.stop_button = ttk.Button(actions, text="Stop", command=self.stop, state="disabled")
        self.stop_button.grid(row=0, column=1, padx=(8, 8))
        self.send_button = ttk.Button(actions, text="Send", style="Accent.TButton", command=self.send)
        self.send_button.grid(row=0, column=2)
        ttk.Label(main, textvariable=self.status_text, foreground=c["muted"], font=("Segoe UI", 9), wraplength=600).grid(row=4, column=0, sticky="w", pady=(10, 0))
        self.prompt.focus_set()

    def _side_wheel(self, event):
        canvas = self.side_canvas
        if (canvas.winfo_rootx() <= event.x_root < canvas.winfo_rootx() + canvas.winfo_width() and
                canvas.winfo_rooty() <= event.y_root < canvas.winfo_rooty() + canvas.winfo_height()):
            canvas.yview_scroll(-int(event.delta / 120), "units")
            return "break"

    def _append(self, text, tag=None):
        self.transcript.configure(state="normal")
        self.transcript.insert("end", text, tag or ())
        self.transcript.configure(state="disabled")
        self.transcript.see("end")

    def _welcome(self):
        self._append("A quiet place for your local agents.\n", "role")
        self._append("Choose an installed Ollama model. Attach a text file, or choose a folder so the agent can read, search and propose edits.\n\n"
                     "Every edit needs your approval. Switching models or folders starts a fresh chat. MiMo presets marked not installed will never switch to another model.\n", "muted")

    def _save(self):
        try:
            self.data_dir.mkdir(parents=True, exist_ok=True)
            temp = self.data_dir / "settings.json.tmp"
            temp.write_text(json.dumps(self.settings, indent=2), encoding="utf-8")
            temp.replace(self.data_dir / "settings.json")
        except OSError:
            self.status_text.set("Settings could not be saved; this session still works.")

    def _confirm_reset(self, description):
        return not (self.agent and self.agent.history) or messagebox.askyesno("Start a new chat?", description + " starts a new chat. Continue?", parent=self.root)

    def _reset(self):
        self.agent = None
        self.transcript.configure(state="normal")
        self.transcript.delete("1.0", "end")
        self.transcript.configure(state="disabled")
        self._welcome()
        self.status_text.set("Ready")

    def new_chat(self):
        if not self.busy and self._confirm_reset("New chat"):
            self._reset()
            self.prompt.focus_set()

    def model_selected(self, event=None):
        if self.busy:
            return
        model = self.model_box.get().strip()
        if not model or model == self.selected_model:
            return
        if not self._confirm_reset("Switching models"):
            self.model_box.set(self.selected_model)
            return
        self.selected_model = model
        self.settings["model"] = model
        self._save()
        self._reset()
        self._model_info()

    def _model_info(self):
        installed = self.selected_model in self.installed_models or self.selected_model + ":latest" in self.installed_models
        self.model_info.set("Installed locally · ready to use" if installed else "Not installed here · select after installation on your target laptop")

    def refresh_models(self):
        if self.busy or self.refreshing or self.closed:
            return
        self.refreshing = True
        self.refresh_button.configure(state="disabled")
        self.connection_text.set("Connecting to local Ollama…")
        channel = self.channel
        factory = self.client_factory
        url = self.settings["url"]

        def worker():
            try:
                models = factory(url).list_models()
                channel.emit("models", models)
            except Exception as exc:
                channel.emit("models_error", str(exc))

        threading.Thread(target=worker, daemon=True).start()

    def choose_folder(self):
        if self.busy:
            return
        path = filedialog.askdirectory(title="Choose the folder this agent may access", parent=self.root, mustexist=True)
        if not path or not self._confirm_reset("Changing the working folder"):
            return
        try:
            self.workspace = Workspace(path)
        except (WorkspaceError, OSError) as exc:
            messagebox.showerror("Cannot use folder", str(exc), parent=self.root)
            return
        self.folder_text.set(str(self.workspace.root))
        self._reset()

    def clear_folder(self):
        if not self.busy and self.workspace and self._confirm_reset("Clearing the working folder"):
            self.workspace = None
            self.folder_text.set("No folder selected\nAttachments-only chat")
            self._reset()

    def attach_files(self):
        if self.busy:
            return
        paths = filedialog.askopenfilenames(title="Attach text, code, Markdown or CSV", parent=self.root,
                                            filetypes=[("Text, code and CSV", "*.txt *.md *.csv *.tsv *.py *.json *.js *.ts *.ps1"), ("All files", "*.*")])
        for name in paths:
            try:
                if len(self.attachments) >= 6:
                    raise WorkspaceError("At most 6 attachments per message. Choose a working folder for more files.")
                path = Path(name).resolve(strict=True)
                if is_sensitive(path):
                    raise WorkspaceError("Sensitive files and paths cannot be attached.")
                text = read_text_file(path)
                remaining = 6000 - sum(len(item.text) for item in self.attachments)
                if remaining < 300:
                    raise WorkspaceError("The 6,000-character attachment budget is full. Remove files or choose a folder.")
                limit = min(4000, remaining - 100)
                excerpt = text[:limit]
                truncated = len(text) > limit
                if truncated:
                    excerpt += "\n[Attachment excerpt truncated. Choose its folder to read more lines.]"
                self.attachments.append(Attachment(path.name, excerpt))
                self.attachments_list.insert("end", path.name + (" · excerpt" if truncated else ""))
            except (OSError, WorkspaceError) as exc:
                messagebox.showerror("Cannot attach file", f"{Path(name).name}:\n{exc}", parent=self.root)

    def remove_attachment(self):
        if not self.busy:
            for index in reversed(self.attachments_list.curselection()):
                self.attachments.pop(index)
                self.attachments_list.delete(index)

    def _set_busy(self, busy):
        self.busy = busy
        for widget in (self.send_button, self.new_button, self.settings_button, self.folder_button,
                       self.clear_folder_button, self.attach_button, self.remove_button):
            widget.configure(state="disabled" if busy else "normal")
        self.model_box.configure(state="disabled" if busy else "normal")
        self.refresh_button.configure(state="disabled" if busy or self.refreshing else "normal")
        self.stop_button.configure(state="normal" if busy else "disabled")
        self.prompt.configure(state="disabled" if busy else "normal")

    def send(self, event=None):
        if self.busy:
            return "break"
        self.model_selected()
        prompt = self.prompt.get("1.0", "end-1c").strip()
        if not prompt or not self.selected_model:
            return "break"
        if self.agent is None or self.agent.cancelled.is_set():
            channel = self.channel
            self.agent = Agent(self.client_factory(self.settings["url"]), self.selected_model, self.workspace,
                               options={"num_ctx": self.settings["num_ctx"], "num_predict": self.settings["num_predict"], "temperature": 0.2},
                               tool_mode=self.settings["tool_mode"], think=self.settings["think"],
                               on_event=channel.emit)
            cancelled = self.agent.cancelled
            self.agent.approve = lambda request: channel.approve(request, cancelled)
        attachments = tuple(self.attachments)
        self._append("\nYou\n", "role")
        self._append(prompt + "\n")
        if attachments:
            self._append("Attached: " + ", ".join(item.name for item in attachments) + "\n", "muted")
        self._append("\n" + self.selected_model + "\n", "role")
        self.prompt.delete("1.0", "end")
        self.attachments.clear()
        self.attachments_list.delete(0, "end")
        self.assistant_chars = 0
        self.started = time.monotonic()
        self._set_busy(True)
        agent = self.agent
        channel = self.channel

        def worker():
            try:
                answer = agent.run(prompt, attachments)
                channel.emit("done", (answer, agent.last_stats))
            except Exception as exc:
                channel.emit("failed", str(exc))

        threading.Thread(target=worker, daemon=True).start()
        return "break"

    def stop(self):
        if self.agent and self.busy:
            self.agent.stop()
            if self.pending_approval:
                self.pending_approval[0]["allowed"] = False
                self.pending_approval[1].set()
            self.activity = "Stopping…"
            self.status_text.set(self.activity)
            self.stop_button.configure(state="disabled")

    def _approval_dialog(self, request):
        window = tk.Toplevel(self.root)
        window.title("Approve this file write")
        window.geometry("820x560")
        window.minsize(600, 420)
        window.configure(bg=COLORS["bg"])
        window.transient(self.root)
        window.grab_set()
        window.columnconfigure(0, weight=1)
        window.rowconfigure(2, weight=1)
        ttk.Label(window, text="Create file?" if request.is_new else "Replace this file?", style="Title.TLabel").grid(row=0, column=0, sticky="w", padx=18, pady=(18, 8))
        ttk.Label(window, text=f"{request.path}\n{request.byte_count:,} UTF-8 bytes · Existing contents will be backed up.", wraplength=760).grid(row=1, column=0, sticky="w", padx=18, pady=(0, 12))
        diff = ScrolledText(window, wrap="none", font=("Consolas", 10), bg=COLORS["input"], fg=COLORS["text"])
        diff.grid(row=2, column=0, sticky="nsew", padx=18)
        diff.insert("1.0", request.diff)
        diff.configure(state="disabled")
        allowed = [False]

        def finish(value):
            allowed[0] = value
            window.destroy()

        buttons = ttk.Frame(window, padding=18)
        buttons.grid(row=3, column=0, sticky="e")
        deny = ttk.Button(buttons, text="Deny", command=lambda: finish(False))
        deny.pack(side="left", padx=(0, 10))
        ttk.Button(buttons, text="Allow this write", style="Accent.TButton", command=lambda: finish(True)).pack(side="left")
        window.protocol("WM_DELETE_WINDOW", lambda: finish(False))
        window.bind("<Escape>", lambda e: finish(False))
        deny.focus_set()
        self.root.wait_window(window)
        return allowed[0]

    def _drain(self):
        if self.closed:
            return
        for _ in range(200):
            if self.closed:
                return
            try:
                kind, value = self.events.get_nowait()
            except queue.Empty:
                break
            if kind in ("models", "models_error"):
                self.refreshing = False
                self.refresh_button.configure(state="disabled" if self.busy else "normal")
                if kind == "models":
                    self.installed_models = {m["name"] for m in value if isinstance(m, dict) and isinstance(m.get("name"), str)}
                    presets = [m for m in MIMO_MODELS if m not in self.installed_models and m + ":latest" not in self.installed_models]
                    self.model_box.configure(values=sorted(self.installed_models) + presets)
                    self.connection_text.set(f"Connected · {len(self.installed_models)} installed models")
                    self._model_info()
                else:
                    self.connection_text.set("Ollama offline · start Ollama, then Refresh")
                    self.model_info.set("Installed-model list unavailable")
                    self.status_text.set(value)
            elif kind == "text" and value:
                self._append(value)
                self.assistant_chars += len(value)
            elif kind == "status":
                self.activity = value
                self.status_text.set(value)
            elif kind == "tool":
                self._append(f"\n  Tool: {value['name']} {value.get('path', '')}\n", "tool")
            elif kind == "tool_result":
                result = value["result"]
                if isinstance(result, dict) and result.get("error"):
                    self._append("  Tool refused: " + result["error"] + "\n", "error")
                elif isinstance(result, dict) and result.get("status"):
                    self._append("  Write: " + result["status"] + (" · backup: " + result["backup"] if result.get("backup") else "") + "\n", "tool")
            elif kind == "approval":
                request, answer, ready = value
                self.pending_approval = (answer, ready)
                if not self.closed and self.agent and not self.agent.cancelled.is_set():
                    answer["allowed"] = self._approval_dialog(request)
                ready.set()
                self.pending_approval = None
            elif kind in ("done", "failed"):
                if kind == "done":
                    answer, stats = value
                    if not self.assistant_chars:
                        self._append(answer)
                    tokens = stats.get("eval_count", 0)
                    self.status_text.set(f"Finished · {time.monotonic() - self.started:.1f}s · {tokens} response tokens")
                else:
                    self._append("\n" + value, "error")
                    self.status_text.set("Stopped / failed · see message above. No model fallback was used.")
                self._append("\n")
                self._set_busy(False)
                self.prompt.focus_set()
        if not self.closed:
            self._later(30, self._drain)

    def _tick(self):
        if self.closed:
            return
        if self.busy:
            self.status_text.set(f"{self.activity} · {int(time.monotonic() - self.started)}s · CPU models can take several minutes")
        self._later(1000, self._tick)

    def show_settings(self):
        if self.busy:
            return
        window = tk.Toplevel(self.root)
        window.title("Local harness settings")
        window.configure(bg=COLORS["bg"])
        window.resizable(False, False)
        window.transient(self.root)
        window.grab_set()
        frame = ttk.Frame(window, padding=24)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Keep it local. Keep it small.", style="Title.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 18))
        variables = {key: tk.StringVar(value=str(self.settings[key])) for key in ("url", "num_ctx", "num_predict", "tool_mode")}
        fields = [("Ollama address", "url", None), ("Context tokens", "num_ctx", (2048, 4096, 8192, 16384, 32768)),
                  ("Response token limit", "num_predict", (128, 256, 512, 768, 1024, 2048)), ("Tool mode", "tool_mode", ("auto", "native", "json"))]
        for row, (label, key, choices) in enumerate(fields, 1):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", padx=(0, 20), pady=8)
            widget = ttk.Combobox(frame, textvariable=variables[key], values=choices, state="readonly", width=28) if choices else ttk.Entry(frame, textvariable=variables[key], width=32)
            widget.grid(row=row, column=1, sticky="ew", pady=8)
            if key == "num_ctx":
                context_box = widget
            elif key == "num_predict":
                response_box = widget

        def update_response_choices(event=None):
            allowed = [value for value in (128, 256, 512, 768, 1024, 2048) if value < int(variables["num_ctx"].get()) // 2]
            response_box.configure(values=allowed)
            if int(variables["num_predict"].get()) not in allowed:
                variables["num_predict"].set(str(allowed[-1]))

        context_box.bind("<<ComboboxSelected>>", update_response_choices)
        update_response_choices()
        thinking = tk.BooleanVar(value=self.settings["think"])
        ttk.Checkbutton(frame, text="Enable model thinking (slower; model must support it)", variable=thinking).grid(row=5, column=0, columnspan=2, sticky="w", pady=12)
        ttk.Label(frame, text="Auto uses native tools when supported. JSON is the compatibility mode for models/templates with broken native calls.\n\n4096 context is the laptop-friendly default. No cloud endpoints, API keys, command execution or automatic model downloads.", wraplength=490, foreground=COLORS["muted"]).grid(row=6, column=0, columnspan=2, sticky="w", pady=(8, 18))

        def save():
            try:
                values = {key: var.get().strip() for key, var in variables.items()}
                OllamaClient(values["url"])
                values["num_ctx"] = int(values["num_ctx"])
                values["num_predict"] = int(values["num_predict"])
                if values["num_predict"] >= values["num_ctx"] // 2:
                    raise ValueError("Response limit must be less than half the context.")
                values["think"] = thinking.get()
            except (ValueError, OllamaError) as exc:
                messagebox.showerror("Invalid settings", str(exc), parent=window)
                return
            if not self._confirm_reset("Changing settings"):
                return
            self.settings.update(values)
            self._save()
            self._reset()
            window.destroy()
            self.refresh_models()

        ttk.Button(frame, text="Save settings", style="Accent.TButton", command=save).grid(row=7, column=1, sticky="e")
        ttk.Button(frame, text="Cancel", command=window.destroy).grid(row=7, column=0, sticky="w")

    def close(self):
        self.channel.closed.set()
        self.stop()
        self.closed = True
        if self.pending_approval:
            self.pending_approval[0]["allowed"] = False
            self.pending_approval[1].set()
        for callback_id in self.after_ids:
            self.root.after_cancel(callback_id)
        self.after_ids.clear()
        self.root.destroy()
