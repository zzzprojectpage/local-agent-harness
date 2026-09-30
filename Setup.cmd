@echo off
setlocal
cd /d "%~dp0"
echo Installing the Excel readers and native Windows Excel adapter into this folder's .venv.
echo This does not install models or change Excel macro security settings.
if exist ".venv\Scripts\python.exe" goto install
where py >nul 2>&1
if not errorlevel 1 (
    py -3 -m venv ".venv"
    if errorlevel 1 goto failed
    goto install
)
where python >nul 2>&1
if errorlevel 1 goto missing
python -m venv ".venv"
if errorlevel 1 goto failed
:install
".venv\Scripts\python.exe" -c "import sys, tkinter; assert sys.version_info >= (3,10), 'Python 3.10+ with Tk is required'"
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r "requirements.txt"
if errorlevel 1 goto failed
echo.
echo Setup completed. Open Launch.cmd.
echo Excel editing and PivotTables require Microsoft Excel desktop.
echo For VBA project access, see VBA-SETUP.md. No global settings were changed.
pause
exit /b 0
:missing
echo Install Python 3.10 or newer from python.org with Tcl/Tk selected, then run Setup.cmd again.
pause
exit /b 1
:failed
echo Setup failed. Read the error above, fix it, and run Setup.cmd again.
pause
exit /b 1
