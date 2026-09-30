# Local Agent Harness

A small, native Windows desktop window for your **local Ollama models**. No Electron,
browser, Node server, API key, subscription or third-party Python package is required.

## Open it

1. Start Ollama on this laptop.
2. Double-click **`Launch.cmd`** in this folder.
3. Pick a model. **Refresh models** reads the models actually installed in Ollama.
4. Use **Attach files…** for text/code/Markdown/CSV, or **Choose folder…** to let the
   agent use real file tools in that folder.
5. Type a message and click **Send** (or press **Ctrl+Enter**).

`Debug.cmd` launches with a console if the window does not open. Direct launch:

```powershell
py -3 run.py
```

Requires **Python 3.10+ with tkinter** and Ollama. Python's standard Windows installer
includes Tk; keep its Tcl/Tk option selected. The app was tested here with Python 3.12.7,
Tk 8.6 and Ollama 0.34.4. There is no `pip install` step.

## Copy to the stronger 16 GB laptop

Copy this entire **Local Agent Harness** folder to its Desktop. Install Python/Ollama
there if needed, start Ollama, and launch the same `Launch.cmd`. Install/create your
existing models on **that laptop**, then refresh the model list. The app does not bundle,
download, rebuild, remove or modify models.

The six presets from your existing MiMo recipes are included in the selector:

- `mimo-2.6-low-fast`, `mimo-2.6-medium-fast`, `mimo-2.6-high-fast`
- `mimo-2.6-low`, `mimo-2.6-medium`, `mimo-2.6-high`

A preset is **not proof the model is installed**. Uninstalled presets are marked as such;
requests fail visibly rather than silently switching to Qwen or any other model. You can
also type another installed Ollama model name into the selector. Models from `/api/tags`
appear automatically, including `excel:latest` and Qwen. The model's Ollama `SYSTEM`
instructions are preserved. Your MiMo recipes also embed an Excel-only restriction in
their template: this harness does not remove it.

## Files and permissions

The harness, not the model alone, executes these tools:

| Tool | Behavior |
| --- | --- |
| `list_files` | Up to 100 files under the selected folder; simple glob patterns. |
| `read_file` | Numbered excerpts from UTF-8/UTF-16 text/code/CSV; use `start_line` for more. |
| `search_text` | Literal, case-insensitive text search, up to 20 matches among the first 100 matching files. |
| `write_file` | Create/replace text only after a **diff approval dialog**. No approval, no write. |

- Paths escaping the selected folder, including external junctions/symlinks, are blocked.
- `.env*`, common credential/key files, `.git`, `.ssh`, virtual environments and
  `node_modules` are skipped. This is a filename/path guard, **not a secret scanner**.
- Existing files are backed up under the chosen folder's
  **`.local-harness-backups/<timestamp-id>/…`** before an approved replacement. The model
  cannot read that backup directory. To restore, manually copy the backup over the file.
- If a file changes while its diff is awaiting approval, the write is refused.
- There is **no shell, deletion tool, plugin/MCP loader or arbitrary Python execution**.
- Attachments may come from outside the chosen folder, but only their explicit text
  excerpts are sent. Attaching a file does **not** authorize access to its neighbors.
- Text files are limited to 256 KB. Up to 6 attachments are allowed per message,
  4,000 characters per excerpt and 6,000 characters total. Truncation is clearly marked.
- Attachment excerpts are consumed by the next message. Choose a working folder to
  read additional lines or handle more files.
- **XLSX/XLSM/XLS/PDF/DOCX/images are not supported** in this basic, dependency-free build.
  Export the needed sheet/range as CSV or attach a text excerpt. The harness will never
  replace an Excel workbook with generated text. It does not automate Excel/COM or run VBA.

Folder confinement is application-level safety, not an OS sandbox against another
malicious process changing links concurrently. Only choose folders you trust. Inspect
generated code and CSV edits before approving them.

## Model and laptop settings

- **Auto** tool mode uses native Ollama tools when the selected model advertises them.
- **JSON** mode is the compatibility option for broken native tool-calling templates.
  It asks for a whole-response JSON envelope and executes only the known file tools.
  The legacy MiMo `<tool_call><function=…>{…}</function></tool_call>` envelope is supported.
- Plain-language answers stream in every mode. JSON/XML/code envelopes wait for the full
  response so a partial tool call is never executed or displayed as a completed action.
- **Native** forces native tools; if a model/template rejects them, use JSON instead.
- Default context: **4096 tokens**; response limit: **768**; model thinking: **off**.
  Start with the `-fast` MiMo variant on a CPU-only laptop. Increase context only if needed.
- Idle models use a 60-second keep-alive. Ollama owns model memory and scheduling.
  Selecting another model does not forcibly unload a model used by other apps.
- **Stop** interrupts the active request and prevents further file tools. Approved writes
  already completed are not undone. Cold CPU models may take minutes to evaluate prompts.
- Switching models, changing the folder or changing settings starts a fresh chat, with
  confirmation if there is conversation history.
- Recent completed turns are bounded (at most 4) and oldest turns are dropped as needed.
  The context-size guard is a conservative character estimate, not a model tokenizer.
  Oversized current turns fail with instructions instead of silently clipping file data.
- The sidebar scrolls on smaller laptop screens; the chat and Send/Stop controls stay visible.

## Local data

Only preferences (model name, local address, context, response limit and tool mode) are
saved in **`data/settings.json`**. Chats and attachment contents stay in memory and are
not auto-saved. Closing the window loses the conversation. No folder is auto-trusted on
restart. The HTTP client accepts loopback Ollama addresses only: no cloud or LAN endpoints.

## Verify

From this folder in PowerShell:

```powershell
py -W error::ResourceWarning -m unittest discover -s tests -v
py -m compileall -q local_harness run.py verify_live.py
py -u verify_live.py --model "qwen3:0.6b"
py -u verify_live.py --model "qwen3:0.6b" --mode json
py -3.10 -u verify_desktop.py
```

Unit/integration tests exercise the actual Tk controls, native/JSON/legacy tool protocols,
loopback HTTP streaming, missing-model rejection, stalled-request cancellation and folder
permissions. The live verifier creates only a temporary text fixture and checks that the
model **actually calls `read_file`**, uses its result, and receives an attachment. Small
models can still give wrong answers; a harness supplies tools, not reasoning capability.

`verify_desktop.py` tests the visible desktop against real Ollama. Its optional
`--screenshot verification/desktop.png` flag needs Pillow only for capturing an image;
Pillow is **not** an application dependency. Verification artifacts are not runtime files.

`demo-workspace` contains harmless sample files to try with **Choose folder…**:
“Read notes.txt using the read_file tool and tell me its verification code.”

API reference used: [Ollama native API](https://docs.ollama.com/api/chat) and
[tool calling](https://docs.ollama.com/capabilities/tool-calling).
