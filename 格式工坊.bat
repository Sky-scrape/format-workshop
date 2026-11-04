@echo off
rem Thin wrapper. All logic and all Chinese text live in tools\app_ctl.py --
rem cmd parses .bat byte-by-byte with the console codepage, so non-ASCII
rem content in a .bat is unreliable. Keep this file ASCII only.
chcp 65001 >nul
cd /d "%~dp0"

rem The portable package ships its own interpreter under python\. In a source
rem checkout that directory does not exist; prefer the project venv (.venv)
rem next, so dev dependencies stay out of any system Python. Fall back to the
rem system pythonw only when neither is present.
set "PYW=%~dp0python\pythonw.exe"
if exist "%PYW%" goto launch
set "PYW=%~dp0.venv\Scripts\pythonw.exe"
if exist "%PYW%" goto launch
set "PYW=pythonw"
where pythonw >nul 2>nul || set "PYW=python"

:launch
rem Quotes matter: the portable package may sit in a path with spaces
rem ("C:\Program Files\..."), and an unquoted %PYW% would split there.
"%PYW%" "tools\app_ctl.py" %*
if errorlevel 1 pause
