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
