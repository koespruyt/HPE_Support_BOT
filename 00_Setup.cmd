@echo off
setlocal
set "HERE=%~dp0"
if "%HERE%"=="" set "HERE=.\"
powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%00_Setup.ps1"
exit /b %errorlevel%
