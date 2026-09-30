# Observed verification — 2026-09-30

## Excel-agent update (supersedes the original text-only limitations below)

### Changed files / entry points

- `local_harness/excel.py:12,102,135,176,204`: bounded OOXML/XLS/XLSB readers,
  template/add-in worksheet inspection, saved caches/formula source where supported.
- `local_harness/excel_edit.py:27,49,81,306,388`: isolated STA worker per native action,
  pre-existing PID protection, native values/formulas/formatting/pivots and explicit VBA.
- `local_harness/workspace.py:149,220,263,278,340`: transactional work copies,
  per-action backups, saved previews, VBA provenance and autonomous text copies.
- `local_harness/agent.py:150,157,184`: actual Excel tool schemas/dispatch, capability
  instructions, bounded same-model recovery, independent listing and completion checks.
- `local_harness/app.py:19,359,373,570`: autonomous default, Excel attachments,
  independent file listing, full output paths, safe-mode retention and session-only VBA consent.
- `Setup.cmd:1`, `requirements.txt:1`, `Launch.cmd:4`, `VBA-SETUP.md:1`, `README.md:7`:
  fresh-environment installation, portable launch/update and one-time VBA prerequisites.
- `tests/`, synthetic `tests/fixtures/sample.xls` / `sample.xlsb`, `verify_excel.py:1`:
  real workbook/model verification. Fixture author metadata sanitized; no external links.

### Observed commands

PowerShell, from `Desktop\Local Agent Harness`, with `$env:PYTHONDONTWRITEBYTECODE='1'`:

| Command | Observed result |
| --- | --- |
| `cmd /c "Setup.cmd < nul"` | Fresh Python **3.12.7** virtual environment; all five pinned direct dependencies installed successfully; no global macro setting changed. |
| `.venv\Scripts\python.exe -W error::ResourceWarning -m unittest discover -s tests -v` | **45 passed**, **108.724s**, no skips. |
| `py -3.10 -W error::ResourceWarning -m unittest discover -s tests -v` | **45 passed**, **108.954s**, no skips. |
| `.venv\Scripts\python.exe -m compileall -q local_harness run.py verify_live.py verify_excel.py verify_desktop.py` | Passed. |
| `py -3.10 -m compileall -q local_harness run.py verify_live.py verify_excel.py verify_desktop.py` | Passed. |
| `git diff --check` | Passed. |
| `py -3.10 -u verify_live.py --model "qwen3:0.6b"` | Actual `read_file`, correct HARNESS-4821 answer, attachment ATTACH-7319; six missing MiMo presets rejected without fallback. File-tool portion **153.2s**. |
| `.venv\Scripts\python.exe -u verify_excel.py --model "excel:latest"` | Actual model-executed native PivotTable; pivot XML present; independent saved-cell check: **A=30, B=5, Grand Total=35**; original hash unchanged. |
| `.venv\Scripts\python.exe -u verify_desktop.py` | Actual desktop model/folder/file call and visible answer passed; absent MiMo preset correctly rejected; **800×620** composer/sidebar layout passed. |
| `Start-Process -FilePath $env:ComSpec -ArgumentList '/c','Launch.cmd' -WorkingDirectory (Get-Location).Path -PassThru` | Owned-process smoke: Launch.cmd created a visible Local Agent Harness window using the setup environment; the newly observed window closed normally through CloseMainWindow. Existing windows were excluded. |

Native tests verify two-cell formula caches, relative formula fill, literal identifiers,
format persistence, XLS/XLSB edits, original/backup preservation, failure cleanup,
VBA creation/import/explicit runs (cells **42** and **73**) and an existing `Auto_Open`
procedure **not** invoked by opening/running another requested Sub. A deliberately infinite
synthetic macro is stopped: no partial edit published, no new EXCEL.EXE remains, and a
pre-existing simulated user workbook stays open with its cell value **77** unchanged.
The Excel project-access setting was already enabled here; tests did not enable it.

### Failures caught, not hidden

- Initial Qwen 0.6B pivot run omitted required arguments and claimed success. Native and
  JSON runs still failed the multi-step task after bounded retries. The new guard refuses
  a final success answer while relevant tool errors remain or no requested pivot succeeded.
  It does not confer reasoning capability on that model.
- Fresh Python 3.12 late-bound COM calls mishandled optional worksheet arguments;
  wrapping the exact owned object with its type library fixed the observed worksheet order.
- COM initialization/uninitialization on a caller apartment disconnected a pre-existing
  COM proxy. Native actions now run entirely in dedicated STA threads; the original proxy
  and process remain usable through successful, failed and stopped actions.
- TerminateProcess is asynchronous; waiting on the exact owned process handle fixed a
  cancelled macro's temporary-file lock/cleanup failure.
- Invariant number formats initially became localized grouping formats. Owned-instance
  separators are set temporarily and restored; saved **0.00** is verified independently.

### Review

**Standards:** strengthened common VBA/DDE screening and weakened the error wording to
avoid claiming a sandbox; narrowed conversational edit detection; required actual pivot
tool evidence for pivot creation requests. Remaining bounds/heuristics are documented.
Do not interpret filename guards, regex screening or final-message checks as proof that
arbitrary VBA or every model answer is safe/correct.

**Spec:** added automatic bounded saved-cell read-back and an actual owned-process Stop
test. Worksheet-less add-ins have metadata but no cell data; VBA code enumeration was
not requested. Review suggestions that 53/17/18 were nonmacro formats were rejected against
Microsoft's XlFileFormat reference (XLTM/XLT/XLA). Programmatic opening does not request
Auto_Open; both Microsoft Workbooks.Open's explicit RunAutoMacros example and the live
synthetic Auto_Open regression support that behavior.

### Limits / not verified

- The other laptop and its installed MiMo models were **not accessible**. Its original
  all-model empty-response root cause was **not reproduced/confirmed**. Listing now bypasses
  inference; other empty replies report privacy-safe counts and use one recovery attempt.
- Model-driven pivot success was observed with the existing **4B Qwen-based excel:latest**,
  not MiMo. Native/JSON protocol regression tests use deterministic model adapters.
- Common formats are supported subject to 100 MB/expanded archive limits, protection and
  installed Excel. Damaged/encrypted/vendor-specific workbooks and every legacy file variant
  are not universally supported. Synthetic XLSM/templates/add-in extensions were exercised;
  this is not exhaustive fidelity testing of complex real-world add-ins/Power Query/slicers.
- VBA execution is explicit trusted code, **not folder-confined or OS-sandboxed**. Events
  are disabled and global security is unchanged, but malicious VBA/XLM/add-ins/connections
  cannot be made safe by text screening. Only enable it for trusted workbooks/code.
- Read-only parser work may take time on large or late ranges; Stop prevents subsequent
  calls, but parser imports/loading are not a hard real-time cancellation boundary.
- Attachments/listing are bounded but currently load/browse on the GUI thread. Large
  attachment parsing or a slow file system can momentarily pause the UI.
- Work-copy aliases/VBA provenance live in memory. Reopen the returned output file after
  restart; previously imported modules are not silently reauthorized for execution.

## Original release verification (historical)

## Delivered files

- `Launch.cmd:1`, `Debug.cmd:1`, `run.py:1`: native Windows launchers.
- `local_harness/app.py:1`: model selector, six MiMo presets, folder picker, attachments,
  streaming transcript, bounded Tk-free worker channel, settings and diff approval dialog.
- `local_harness/agent.py:1`: exact-model agent loop; native/JSON/legacy MiMo tool calls.
- `local_harness/workspace.py:1`: constrained text tools, approval, backups and race checks.
- `local_harness/ollama.py:1`: loopback REST client, streaming and interruptible Windows reads.
- `tests/`, `verify_live.py:1`, `verify_desktop.py:1`: automated verification.
- `README.md:1`, `SPEC.md:1`, `demo-workspace/`: transfer instructions, scope and sample files.

All task files are new under `Desktop\Local Agent Harness`. No existing Ollama models,
MiMo recipes, editor configuration, other Desktop files or Default Project files were changed.
This is a standalone delivery folder, not a commit to the unrelated Default Project repo.

## Exact commands and results

Run from the delivery folder in PowerShell. The final test run used
`$env:PYTHONDONTWRITEBYTECODE='1'` to avoid generating delivery caches.

| Command | Observed result |
| --- | --- |
| `py -W error::ResourceWarning -m unittest discover -s tests -v` | **23 passed**, Python 3.12.7, 4.978s; no skips. |
| `py -3.10 -W error::ResourceWarning -m unittest discover -s tests -v` | **23 passed**, Python 3.10.11, 4.749s; no skips. |
| `py -u verify_live.py --model 'qwen3:0.6b'` | Real `read_file` call, correct `HARNESS-4821` answer, attachment `ATTACH-7319` received. Native/auto mode passed. |
| `py -u verify_live.py --model 'qwen3:0.6b' --mode json` | Real file calls + correct file/attachment answers; all six absent MiMo models rejected without fallback. |
| `py -3.10 -u verify_desktop.py --screenshot 'verification\desktop.png'` | Actual desktop selection, folder picker, live file call and visible answer passed; missing MiMo preset failed correctly; 800×620 layout passed. Screenshot inspected. |
| `py -m compileall -q local_harness run.py verify_live.py verify_desktop.py` | Final source passed on Python 3.12.7. |
| `py -3.10 -m compileall -q local_harness run.py verify_live.py verify_desktop.py` | Final source passed on Python 3.10.11. Generated bytecode removed afterward. |
| `cmd /c Launch.cmd` | Started a Python window without a terminal. Win32 window enumeration confirmed a visible `Local Agent Harness` root window. |

The launched Python/Tk process was observed at **32.4 MB working set / 18.9 MB private memory**
after startup. This excludes Ollama and model memory; it is an observation on this laptop,
not a RAM/performance guarantee on the target laptop.

## Review

### Standards

The read-only standards review found no acceptance blockers. Addressed the concrete
robustness items: worker event queue now bounded; workers/approval callbacks do not own
Tk objects; cancel no longer extracts the response socket's private `_sock` attribute.
The character/token estimate remains explicitly documented. A proposed string-prefix
path shortcut was rejected because resolved-path confinement is the security boundary.
Architecture/naming suggestions were not applied merely to enlarge a basic app.

### Spec

The read-only spec review identified JSON-mode streaming, invalid settings combinations,
excerpt navigation, partial MiMo checks and generated artifacts. Fixed prose streaming
with full-envelope validation for tool calls, filtered response limits by context, reduced
source excerpts to fit tool results, and checked all six absent MiMo models. The screenshot
is a verification-only artifact, not an application dependency; generated bytecode caches
are removed from the delivery.

## Assumptions / not verified

- The other Windows laptop has Python 3.10+ with Tk and Ollama. The folder-copy procedure
  is documented, but the actual other laptop was not accessible for testing.
- MiMo model recipes are present locally, but **no MiMo model is installed in this Ollama**.
  Selection, missing-model rejection and its legacy protocol adapter were verified;
  actual MiMo inference/file-use behavior was **not** verified.
- Text/code/Markdown/CSV only. No native Excel, PDF, image or terminal-command automation.
- Folder guards are not an OS sandbox against another malicious process racing link changes.
  Context estimation is not model-tokenizer exact. Model reasoning accuracy is not guaranteed.
