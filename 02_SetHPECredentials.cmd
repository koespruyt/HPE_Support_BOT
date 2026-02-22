@echo off
setlocal
set "here=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%here%02_SetHPECredentials.ps1"
endlocal
