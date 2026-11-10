@echo off
rem Dev-checkout entry. No arguments -> launch the desktop shell (native
rem window; pages are answered by the in-process engine -- no port, no
rem background service). Any argument (stop / status / shortcut / quiet)
rem forwards to tools\app_ctl.py; stop/status only concern the optional
rem browser-service mode (python app.py).
chcp 65001 >nul
cd /d "%~dp0"
set "PAPER_PORT=5001"
set "PYW=%~dp0.venv\Scripts\pythonw.exe"
if exist "%PYW%" goto pick

echo [dev] .venv not found. Create it once, then retry:
echo     py -3 -m venv .venv
echo     .venv\Scripts\python -m pip install -r requirements-dev.txt
pause
exit /b 1

:pick
if not "%~1"=="" goto forward

rem start (not a plain call): desktop.py lives for the whole session, and a
rem blocking call would keep this console open the entire time.
start "" "%PYW%" "%~dp0desktop.py"
exit /b 0

:forward
rem app_ctl reports go to stdout; pythonw has no console (sys.stdout/stderr
rem are None) so every print would be silently dropped. PYW is always the
rem venv pythonw here and its console sibling exists, so inline-substitute.
"%PYW:pythonw.exe=python.exe%" "tools\app_ctl.py" %*
if errorlevel 1 pause
exit /b %errorlevel%
