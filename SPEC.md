# Delivery contract

User-approved scope: a minimal native desktop harness for existing local Ollama agents,
portable to another Windows 16 GB RAM laptop. Model switching, explicit text/code file
attachments, working-folder selection, real folder read/search tools, approval before
text edits, no terminal command execution. Deliver all files in a dedicated Desktop folder.

## Acceptance seams

1. Visible UI: model selection (including MiMo presets), folder picker, attachments and chat.
2. Model adapter: exact selected Ollama model; streaming response; Stop; missing model never falls back.
3. Folder permissions: actual reads/search; no traversal or external-link escapes; approved,
   backed-up text writes; denied writes leave the original untouched.
4. Live evidence: test available Qwen here; report honestly if MiMo is not installed.

No existing model recipes, Copilot/Jcode/OpenCode configuration or unrelated project files
may be changed. No model downloads. No browser/Electron dependency. Standard Python/Tk is
the runtime requirement. Native Excel/binary-document automation is outside this basic scope.

The Desktop output is a standalone folder, not a change to the Default Project repository.

## Excel-agent scope amendment (2026-09-30)

The user superseded the text-only, approval-per-edit scope: support native Excel files,
autonomous cell/value/formula edits, real PivotTables and VBA creation/execution.
Continue testing at the existing workspace, model-adapter and visible-UI seams.

- Read XLSX/XLSM/XLSB/XLS, XLTX/XLTM/XLT and XLAM without executing macros.
- Autonomous Excel and text edits use session-owned working copies with backups;
  originals remain unchanged. Safe, approval-based text editing remains available.
- Native editing/PivotTables require installed Microsoft Excel and pywin32.
- VBA modules can be created/imported; XLSX is converted to a separate XLSM copy.
  Execution requires one explicit session/folder opt-in and Excel's user-enabled
  'Trust access to the VBA project object model'. Never change global macro settings.
- Explicitly run only session-created/imported modules, not workbook autorun events.
  VBA is not an OS sandbox: the opt-in must disclose that trusted code can access the PC.
- Tools report output/backup paths. Failed/stopped operations do not publish partial edits.
- Folder listing works without inference; empty-model diagnostics expose counts, not data.
- Verify native PivotTable totals, formula caches, VBA cell changes and owned-instance
  cleanup on synthetic workbooks. Do not claim every local model can reason/use tools.
