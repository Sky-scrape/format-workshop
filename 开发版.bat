@echo off
rem Dev-checkout entry point. Forwards to tools\app_ctl.py -- the same logic
rem as the main .bat -- but pins port 5001 and uses the .venv interpreter
rem only, so the usage copy (portable package / installed setup, port 5000)
rem can run at the same time without either one seeing the other. A dev run
rem on some system Python would silently defeat that isolation.
chcp 65001 >nul
cd /d "%~dp0"
set "PAPER_PORT=5001"
set "PYW=%~dp0.venv\Scripts\pythonw.exe"
if exist "%PYW%" goto launch

echo [dev] .venv not found. Create it once, then retry:
echo     py -3 -m venv .venv
echo     .venv\Scripts\python -m pip install -r requirements-dev.txt
pause
exit /b 1

:launch
rem Quotes matter: an unquoted %PYW% would split on paths with spaces.
"%PYW%" "tools\app_ctl.py" %*
if errorlevel 1 pause
