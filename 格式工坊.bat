@echo off
rem Thin wrapper. All logic and all Chinese text live in tools\app_ctl.py --
rem cmd parses .bat byte-by-byte with the console codepage, so non-ASCII
rem content in a .bat is unreliable. Keep this file ASCII only.
chcp 65001 >nul
cd /d "%~dp0"

rem The portable package ships its own interpreter under python\. Prefer it, so
rem the app depends on nothing installed on the machine. In a source checkout
rem that directory does not exist and we fall back to the system pythonw --
rem same behaviour as before this line was added.
set "PYW=%~dp0python\pythonw.exe"
if exist "%PYW%" goto launch
set "PYW=pythonw"
where pythonw >nul 2>nul || set "PYW=python"

:launch
rem Quotes matter: the portable package may sit in a path with spaces
rem ("C:\Program Files\..."), and an unquoted %PYW% would split there.
"%PYW%" "tools\app_ctl.py" %*
if errorlevel 1 pause
