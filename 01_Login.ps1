<#
01_Login.ps1  (RUN AFTER 00_Setup.ps1)
Starts interactive login to HPE and saves session cookies into hpe_state.json.

Usage:
  powershell -ExecutionPolicy Bypass -File .\01_Login.ps1
#>
[CmdletBinding()]
param(
  [string]$OutState = "hpe_state.json"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$marker = Join-Path $root ".setup.ok"
if (-not (Test-Path $marker)) {
  throw "Setup not done. Run: powershell -ExecutionPolicy Bypass -File .\00_Setup.ps1"
}

$venvPy = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $venvPy)) {
  throw "Missing venv python. Run setup again: .\00_Setup.ps1"
}

& $venvPy .\01_login_save_state.py --out $OutState
