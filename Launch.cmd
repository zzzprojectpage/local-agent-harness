@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\pythonw.exe" (
    start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0run.py"
    exit /b
)
where pyw >nul 2>&1
if not errorlevel 1 (
    start "" pyw -3 "%~dp0run.py"
    exit /b
)
where pythonw >nul 2>&1
if not errorlevel 1 (
    start "" pythonw "%~dp0run.py"
    exit /b
)
echo Python 3.10 or newer with tkinter is required.
echo Install Python from python.org on your target laptop, then try again.
echo Run Setup.cmd once for Excel support.
pause
