@echo off
rem Thin wrapper. All logic and all Chinese text live in tools\app_ctl.py --
rem cmd parses .bat byte-by-byte with the console codepage, so non-ASCII
rem content in a .bat is unreliable. Keep this file ASCII only.
chcp 65001 >nul
cd /d "%~dp0"

set "PYW=pythonw"
where pythonw >nul 2>nul || set "PYW=python"

%PYW% "tools\app_ctl.py" %*
if errorlevel 1 pause
