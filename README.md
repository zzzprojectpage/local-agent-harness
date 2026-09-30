# Local Agent Harness

A small, native Windows desktop **Excel agent for local Ollama models**. It supplies
real workbook tools—not just advice or generated code. No Electron, browser server,
cloud inference, API key or subscription.

## Open it

1. Install **Python 3.10+ with Tcl/Tk**, **Ollama** and **Microsoft Excel desktop** on Windows.
   Workbook reading does not require Excel; editing, native pivots and VBA do.
2. Run **`Setup.cmd`** once. It installs the pinned Excel dependencies into this folder's
   `.venv`; no model downloads or global Excel settings are changed.
3. Start Ollama, then double-click **`Launch.cmd`**. Pick an installed model.
4. **Choose folder…** containing your workbook. Autonomous working-copy edits are on
   by default: no approval dialog for each edit. **List files (no model)** or `/files`
   lists the folder even if the model cannot generate an answer.
5. Ask for an operation and click **Send** (or **Ctrl+Enter**). The tool log shows the
   actual **Output** and **backup** paths. The original input is not overwritten.
6. For VBA creation/import and explicit execution, follow **[VBA-SETUP.md](VBA-SETUP.md)**.
   Execution requires one trusted-folder/session opt-in in Settings—not a prompt per macro.

Examples (use your actual file, sheet and header names):

- “Inspect sales.xlsx and read the first 12 rows of Sales.”
- “Change B2 to 5 and put =B2*C2 in D2 in sales.xlsx, sheet Sales. Verify the saved result.”
- “Create a native PivotTable in sales.xlsx from Data!A1:D100, rows Region, sum of Amount,
  on a new sheet Summary. Verify its totals.”
- “Create a standard VBA module in sales.xlsx with a public zero-argument Sub that sets
  Data!B2 to 42, run it, and verify the cell.”

`Debug.cmd` launches with a console if the window does not open. Direct launch:

```powershell
.venv\Scripts\python.exe run.py
```

`Launch.cmd` prefers the setup-created environment. Without it, text chat can still use
an installed Python; workbook operations report missing dependencies with setup guidance.

## Copy to the stronger 16 GB laptop

Download the latest repository ZIP or pull `main` in your existing clone. Put the files
on the laptop's Desktop and run **Setup.cmd on that laptop**, then Launch.cmd. Do not
transfer `.venv` between laptops; recreate it with Setup. Keep your workbooks outside
the application folder. Install/create existing models on **that laptop**, then refresh.
The app does not bundle, download, rebuild, remove or modify models.

The six presets from your existing MiMo recipes are included in the selector:

- `mimo-2.6-low-fast`, `mimo-2.6-medium-fast`, `mimo-2.6-high-fast`
- `mimo-2.6-low`, `mimo-2.6-medium`, `mimo-2.6-high`

A preset is **not proof the model is installed**. Uninstalled presets are marked as such;
requests fail visibly rather than silently switching to Qwen or any other model. You can
also type another installed Ollama model name into the selector. Models from `/api/tags`
appear automatically, including `excel:latest` and Qwen. The model's Ollama `SYSTEM`
instructions are preserved. Your MiMo recipes also embed an Excel-only restriction in
their template: this harness does not remove it.

## Workbook formats and capabilities

| Formats | Reads | Native Excel edits |
| --- | --- | --- |
| XLSX, XLSM | Values, sheet metadata, formula source and saved formula caches | Yes, preserving native workbook features through Excel |
| XLTX, XLTM, XLAM | Worksheet data/metadata; add-ins without worksheets have no data preview | Yes, subject to Excel's format/protection restrictions |
| XLS, XLT | Saved values and sheet metadata; no formula source from the binary reader | Yes |
| XLSB | Saved values/metadata; dates may be Excel serial numbers; no formula source | Yes, remains XLSB |
| CSV, TSV | Text excerpts | Text working-copy edits |

These are common Excel formats, **not a promise to open every damaged, encrypted or
vendor-specific file**. Decrypt password-protected workbooks in Excel to a trusted copy
first. Sheet/workbook protection is not bypassed. PDF/Word/images, ODS, Data Model/Power
Query/slicer authoring and arbitrary COM calls are not tools in this release.

Workbook reads are read-only and do not execute macros or refresh external data.
Read at most **400 cells per call**; use explicit small A1 ranges. Formula caches can
be absent/stale until calculated in Excel. Workbook input limit: **100 MB**, expanded
OOXML archive limit: **256 MB**. Work only with trusted files and review final outputs.

## Tools and permissions

The harness, not the model alone, executes these tools:

| Tool | Behavior |
| --- | --- |
| `list_files` | Up to 100 files under the selected folder; simple glob patterns. |
| `read_file` | Numbered text/code/CSV excerpts, or a native workbook preview. |
| `search_text` | Literal, case-insensitive text search, up to 20 matches among the first 100 matching files. |
| `write_file` | Autonomous text/code/CSV copies; safe mode retains the diff approval dialog. Cannot overwrite a workbook with text. |
| `inspect_workbook`, `read_excel` | Actual sheet metadata and bounded cells/formulas. |
| `create_workbook`, `add_sheet` | New native workbook or worksheet; no replacement of an existing name. |
| `set_cells`, `fill_formula`, `format_cells` | Change values/formulas, fill relative formulas, number formats, bold and cell fill. |
| `create_pivot` | A real interactive Excel PivotTable, not a static pandas summary. Row/column fields; sum/count/average/min/max. Uses an empty/new destination sheet. |
| `create_vba_module`, `import_vba_module` | New standard VBA module; XLSX/XLTX becomes a separate XLSM copy. |
| `run_vba_macro` | Explicitly run a named zero-argument Sub from a session-created/imported module when trusted execution is enabled. |

- Paths escaping the selected folder, including external junctions/symlinks, are blocked.
- `.env*`, common credential/key files, `.git`, `.ssh`, virtual environments and
  `node_modules` are skipped. This is a filename/path guard, **not a secret scanner**.
- Existing inputs become **`name.harness-<id>.xlsx`** (or their original extension), not
  in-place edits. Reuse the original path in the same folder/session: reads and edits
  follow its working copy. Output aliases are in memory; on restart choose the output file
  to continue. Text safe mode is the exception: an explicitly approved edit replaces its target.
- Previous versions are backed up under the chosen folder's
  **`.local-harness-backups/<timestamp-id>/…`** before each edit. The model
  cannot read that backup directory. To restore, manually copy the backup over the file.
- Failed/cancelled Excel actions do not publish partial changes. Concurrent changes are
  detected before publication. Completed earlier operations are not undone by Stop.
- Cell/formula/format/pivot writes are reopened for a small saved-value/formula preview
  before publication. This verifies the saved data is readable, not that every formula or
  model interpretation is correct. Text identifiers such as `00123` remain text.
- There is **no terminal, deletion, plugin/MCP or arbitrary Python execution tool**.
  **Opt-in VBA is executable code and is not sandboxed**; folder guards cannot confine it.
  Common unsafe VBA/formula APIs are screened, but that is not a security proof.
- Attachments can be outside the chosen folder; only explicit text/Excel previews are
  sent. Attaching does **not** authorize file editing or access to neighbors.
- Text files are limited to 256 KB. Up to 6 attachments are allowed per message,
  4,000 characters per excerpt and 6,000 characters total. Truncation is clearly marked.
- Attachment excerpts are consumed by the next message. Choose a working folder to
  read additional lines or handle more files.

Folder confinement is application-level safety, not an OS sandbox against another
malicious process changing links concurrently. VBA may access the whole PC under your
Windows account. Only choose trusted folders/workbooks and enable trusted VBA deliberately.

## Model and laptop settings

- **Auto** tool mode uses native Ollama tools when the selected model advertises them.
- **JSON** mode is the compatibility option for broken native tool-calling templates.
  It asks for a whole-response JSON envelope and executes only the known file tools.
  The legacy MiMo `<tool_call><function=…>{…}</function></tool_call>` envelope is supported.
- Plain-language read/chat answers stream. Edit-turn answers wait for completion checks;
  failed tools trigger one correction attempt, not an accepted false-success answer.
  JSON/XML/code envelopes wait for validation; partial calls are never executed.
- **Native** forces native tools; if a model/template rejects them, use JSON instead.
- Default context: **4096 tokens**; response limit: **768**; model thinking: **off**.
  Start with the `-fast` MiMo variant on a CPU-only laptop. Increase context only if needed.
- **8192 context** is useful for multi-step Excel/VBA tasks. There is a bounded 16-step
  tool loop. Context estimation is conservative, not a model tokenizer.
- Empty final replies get one recovery attempt with thinking off (and JSON tools in
  Auto mode). The transcript reports **counts/reason only**, not private reasoning/data.
  This does not identify the other laptop's original model/template fault by itself.
- Idle models use a 60-second keep-alive. Ollama owns model memory and scheduling.
  Selecting another model does not forcibly unload a model used by other apps.
- **Stop** interrupts inference and active native automation, affecting only the Excel
  process the harness owns. Native operations time out at 180 seconds; macro runs at 60.
  Cold CPU models can still take minutes to evaluate prompts.
- Switching models, changing the folder or changing settings starts a fresh chat, with
  confirmation if there is conversation history.
- Recent completed turns are bounded (at most 4) and oldest turns are dropped as needed.
  The context-size guard is a conservative character estimate, not a model tokenizer.
  Oversized current turns fail with instructions instead of silently clipping file data.
- The sidebar scrolls on smaller laptop screens; the chat and Send/Stop controls stay visible.

## Local data

Only preferences (model name, local address, context, response limit, tool mode and edit mode) are
saved in **`data/settings.json`**. Chats and attachment contents stay in memory and are
not auto-saved. Closing loses the conversation and working-copy aliases. No folder or
VBA execution grant is auto-trusted on restart. The model HTTP client accepts loopback
Ollama addresses only; VBA's access is separate and not constrained by that client.

## Verify

From this folder in PowerShell:

```powershell
.venv\Scripts\python.exe -W error::ResourceWarning -m unittest discover -s tests -v
.venv\Scripts\python.exe -m compileall -q local_harness run.py verify_live.py verify_excel.py
.venv\Scripts\python.exe -u verify_live.py --model "qwen3:0.6b"
.venv\Scripts\python.exe -u verify_excel.py --model "qwen3:0.6b" --mode json
.venv\Scripts\python.exe -u verify_desktop.py
```

Unit/integration tests exercise the actual Tk controls, native/JSON/legacy tool protocols,
loopback HTTP streaming, missing-model rejection, stalled-request cancellation and folder
permissions, native formula caches, binary Excel readers/edits, native pivot XML/totals,
working-copy rollback and actual VBA cell changes. Native tests require installed Excel;
the VBA test requires project access already enabled. Tests never enable it globally.
The live text verifier creates a temporary fixture and checks that the
model **actually calls `read_file`** and uses its result. `verify_excel.py` requires an
actual model-executed native pivot with independently checked totals, not its prose claim.
Small models can fail even with working tools. A harness supplies capabilities and checks,
not the reasoning quality of a larger model. MiMo inference needs MiMo installed locally.
In this update, **Qwen 0.6B failed pivot creation** in native and JSON modes; the existing
**4B `excel:latest` succeeded**. Use a capable tool-calling model for autonomous Excel,
not the smallest chat model simply because it responds quickly.

`verify_desktop.py` tests the visible desktop against real Ollama. Its optional
`--screenshot verification/desktop.png` flag needs Pillow only for capturing an image;
Pillow is **not** an application dependency. Verification artifacts are not runtime files.

`demo-workspace` contains harmless sample files to try with **Choose folder…**:
“Read notes.txt using the read_file tool and tell me its verification code.”

API reference used: [Ollama native API](https://docs.ollama.com/api/chat) and
[tool calling](https://docs.ollama.com/capabilities/tool-calling),
[Excel automation security](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity),
[PivotCaches.Create](https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcaches.create),
[openpyxl](https://openpyxl.readthedocs.io/en/stable/),
[xlrd](https://xlrd.readthedocs.io/en/latest/), [pyxlsb](https://github.com/willtrnr/pyxlsb).
