@echo off
setlocal EnableExtensions

REM ==========================================================
REM HPE CaseBot installer (CMD wrapper)
REM - Calls PowerShell installer to avoid CMD parsing issues.
REM ==========================================================

set "PKGDIR=%~dp0"
if "%PKGDIR:~-1%"=="\" set "PKGDIR=%PKGDIR:~0,-1%"

echo.
echo [HPE CaseBot] Package location: %PKGDIR%
echo.

set "TARGET="
set /p TARGET=Install folder (default = package folder) ^> 
if not defined TARGET (
  echo Using in-place install.
  set "TARGET=%PKGDIR%"
)

for %%I in ("%TARGET%") do set "TARGET=%%~fI"

powershell -NoProfile -ExecutionPolicy Bypass -File "%PKGDIR%\Install.ps1" -Target "%TARGET%" -Package "%PKGDIR%"
exit /b %errorlevel%
