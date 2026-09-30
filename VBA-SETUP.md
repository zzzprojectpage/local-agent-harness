# One-time VBA setup

Use Microsoft Excel desktop on Windows. Run `Setup.cmd` to install the Python adapter.
Workbook reads, cell edits and PivotTables do **not** need VBA project access.

## Allow the agent to create/import VBA modules

In Excel: **File → Options → Trust Center → Trust Center Settings → Macro Settings →
Trust access to the VBA project object model**. Check that box and accept the dialogs.
Restart Excel if required. The harness does not change this setting for you. A managed
company laptop may block it; ask its administrator rather than bypassing the policy.

You do **not** need to select “Enable all macros” globally or trust an entire drive.
Leave global macro protections enabled. Office policy or downloaded-file restrictions
may still prevent explicit execution; use trusted, locally created files.

## Allow explicit execution for this session

1. In the harness, choose the folder containing your trusted workbook.
2. Open **Settings**. Leave **Autonomous edits on backed-up working copies** checked.
3. Check **Allow trusted VBA execution in this folder/session** and save.
4. Ask for a standard VBA module and its named, zero-argument `Public Sub` to be run.

This is one session/folder authorization, not a prompt for each macro. Changing folders
or restarting the app revokes it. Only modules created/imported through the harness in
this session can be run. Existing modules are not replaced; choose a new module name.

XLSX/XLTX inputs are converted to a separate XLSM working copy before VBA insertion.
The tool log gives the actual output and backup paths. To import an existing script,
place a `.bas` standard-module file in the chosen folder and request `import_vba_module`.
UserForms, document/event modules and their binary resources are not imported.

**VBA is not sandboxed.** It can access the PC, other workbooks and files outside the
chosen folder. Text screening rejects common shell/network/file-system/dialog APIs,
but cannot prove code safe. Enable execution only for code/workbooks you trust.
VBA must not request terminal commands, external access or unattended dialogs. Prefer
`ThisWorkbook.Worksheets("Data").Range("B2")` to unqualified `Range`/`ActiveWorkbook`.

The harness disables Excel events before opening files and does not call `Auto_Open`
or `Workbook_Open`. It temporarily enables macros only in its isolated, owned Excel
instance when explicitly running a requested macro. This is **not** a guarantee that
arbitrary existing VBA/XLM/add-ins or data connections in a malicious workbook are safe.
Macro runs time out after 60 seconds; other native operations after 180 seconds.
Stop terminates only the owned automation instance. Completed edits remain; an interrupted
operation is not published. Do not use this feature on untrusted workbook downloads.
