@echo off
setlocal EnableExtensions

REM Run the bot from CMD / Scheduled Task.
REM Important: avoid UnicodeEncodeError on Windows consoles by forcing UTF-8.
chcp 65001 >nul 2>nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

set "HERE=%~dp0"
if "%HERE%"=="" set "HERE=.\"
powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%Run-HPECaseBot.ps1" %*
exit /b %errorlevel%
