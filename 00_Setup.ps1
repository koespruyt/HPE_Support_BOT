# 00_Setup.ps1
# - Creates/updates venv (.venv)
# - Installs Python deps (requirements.txt)
# - Installs Playwright Chromium (needed for portal browsing)
# This script is non-interactive and safe to run multiple times.

$ErrorActionPreference = "Stop"

function Write-Info($m){ Write-Host "[Setup] $m" }

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

# Pick system Python (prefer 'py -3.12' if available)
$SysPyExe = $null
$SysPyArgs = @()

try { & py -3.12 -V *> $null; $SysPyExe = "py"; $SysPyArgs = @("-3.12") } catch {}
if (-not $SysPyExe) { try { & py -V *> $null; $SysPyExe = "py"; $SysPyArgs = @() } catch {} }
if (-not $SysPyExe) { try { & python -V *> $null; $SysPyExe = "python"; $SysPyArgs = @() } catch {} }

if (-not $SysPyExe) {
  throw "Python not found. Install Python 3.12+ (preferred) or ensure 'py' or 'python' is available."
}

Write-Info "Using system python: $SysPyExe $($SysPyArgs -join ' ')"

# Ensure venv
$VenvPy = Join-Path $Root ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $VenvPy)) {
  Write-Info "Creating venv .venv ..."
  & $SysPyExe @SysPyArgs -m venv ".venv"
}

# Upgrade pip
Write-Info "Upgrading pip..."
& $VenvPy -m pip install --upgrade pip

# Install deps
$Req = Join-Path $Root "requirements.txt"
if (Test-Path -LiteralPath $Req) {
  Write-Info "Installing requirements..."
  & $VenvPy -m pip install -r $Req
} else {
  Write-Info "No requirements.txt found; skipping pip install."
}

# Install Playwright browser
Write-Info "Installing Playwright Chromium..."
& $VenvPy -m playwright install chromium

# Marker file
$Marker = Join-Path $Root ".setup.ok"
Set-Content -LiteralPath $Marker -Encoding UTF8 -Value ("Setup OK: " + (Get-Date).ToString("s"))
Write-Info "OK (marker: .setup.ok)"
