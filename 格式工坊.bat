@echo off
rem Thin wrapper. All logic and all Chinese text live in tools\app_ctl.py --
rem cmd parses .bat byte-by-byte with the console codepage, so non-ASCII
rem content in a .bat is unreliable. Keep this file ASCII only.
chcp 65001 >nul
cd /d "%~dp0"

rem Interpreter: the portable package ships its own under python\; a source
rem checkout falls back to the project venv, then to the system pythonw.
set "PYW=%~dp0python\pythonw.exe"
if exist "%PYW%" goto pick
set "PYW=%~dp0.venv\Scripts\pythonw.exe"
if exist "%PYW%" goto pick
set "PYW=pythonw"
where pythonw >nul 2>nul || set "PYW=python"

:pick
rem No arguments -> desktop shell (native window; pages are answered by the
rem in-process engine -- no port, no background service). Named arguments
rem (stop / status / shortcut / quiet) forward to tools\app_ctl.py.
if not "%~1"=="" goto forward

rem start (not a plain call): desktop.py lives for the whole session, and a
rem blocking call would keep this console open the entire time.
start "" "%PYW%" "%~dp0desktop.py"
exit /b 0

:forward
rem app_ctl reports (stop / status / shortcut) go to stdout, and pythonw has
rem no console -- sys.stdout/stderr are None there, so every print is dropped
rem and the user only sees a silent window. Forward with the console python;
rem only the no-argument shell launch keeps pythonw (double-click must not
rem flash a console window).
set "PYC=%PYW:pythonw.exe=python.exe%"
if "%PYC%"=="pythonw" set "PYC=python"
if not exist "%PYC%" set "PYC=%PYW%"
rem Quotes matter: the portable package may sit in a path with spaces
rem ("C:\Program Files\..."), and an unquoted %PYC% would split there.
"%PYC%" "tools\app_ctl.py" %*
if errorlevel 1 pause
