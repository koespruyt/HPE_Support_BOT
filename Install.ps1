param(
  [Parameter(Mandatory=$true)][string]$Target,
  [Parameter(Mandatory=$true)][string]$Package
)

$ErrorActionPreference = "Stop"

function Write-Info($m){ Write-Host "[HPE CaseBot] $m" }
function Write-Warn($m){ Write-Host "[HPE CaseBot] WARNING: $m" -ForegroundColor Yellow }
function Write-Err($m){ Write-Host "[HPE CaseBot] ERROR: $m" -ForegroundColor Red }

function Resolve-FullPath([string]$p){
  try { return (Resolve-Path -LiteralPath $p).Path } catch { return (New-Object System.IO.DirectoryInfo($p)).FullName }
}

$Package = Resolve-FullPath $Package
$Target  = Resolve-FullPath $Target

Write-Info "Package location: $Package"
Write-Info "Install root:      $Target"

if (!(Test-Path -LiteralPath $Target)) {
  Write-Info "Creating install folder..."
  New-Item -ItemType Directory -Path $Target | Out-Null
}

# Copy package -> target if needed
if ($Target.TrimEnd('\') -ne $Package.TrimEnd('\')) {
  Write-Info "Copying files to target (this may take a moment)..."
  $null = robocopy $Package $Target /MIR /XD ".venv" "out_hpe" "out_hpe_archive" 2>$null
  if ($LASTEXITCODE -ge 8) { throw "Robocopy failed with code $LASTEXITCODE" }
}

Set-Location -LiteralPath $Target

# Unblock scripts (avoid security prompts)
Write-Info "Unblocking scripts..."
Get-ChildItem -Recurse -File -Include *.ps1,*.cmd,*.py,*.json,*.txt -ErrorAction SilentlyContinue |
  ForEach-Object { try { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue } catch {} }

# ---- Python detection / optional install
Write-Info "Checking Python..."
$pyCmd = $null
try { & py -3.12 -V *> $null; $pyCmd = @("py","-3.12") } catch {}
if (-not $pyCmd) { try { & py -V *> $null; $pyCmd = @("py") } catch {} }
if (-not $pyCmd) { try { & python -V *> $null; $pyCmd = @("python") } catch {} }

if (-not $pyCmd) {
  Write-Warn "Python not found (py/python not available)."
  $hasWinget = $false
  try { $null = Get-Command winget -ErrorAction Stop; $hasWinget = $true } catch {}
  if ($hasWinget) {
    $ans = Read-Host "Install Python 3.12 via winget now? (Y/N)"
    if ($ans -match '^(y|yes)$') {
      Write-Info "Running: winget install -e --id Python.Python.3.12"
      winget install -e --id Python.Python.3.12
      Start-Sleep -Seconds 3
      try { & py -3.12 -V *> $null; $pyCmd = @("py","-3.12") } catch {}
      if (-not $pyCmd) { try { & python -V *> $null; $pyCmd = @("python") } catch {} }
    }
  }
}

if (-not $pyCmd) {
  Write-Err "Python still not available. Install Python 3.12+ and rerun install_me.cmd."
  exit 2
}

# Verify version >= 3.11 (Playwright + modern TLS; adjust if you want strict 3.12)
$verOk = $false
$ver = $null
try {
  $exe    = $pyCmd[0]
  $pyArgs = @()
  if ($pyCmd.Length -gt 1) { $pyArgs += $pyCmd[1..($pyCmd.Length-1)] }

  $ver = & $exe @pyArgs -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"
  $parts = $ver.Trim().Split('.')
  if ($parts.Count -ge 2) {
    $major = [int]$parts[0]
    $minor = [int]$parts[1]
    if ($major -gt 3 -or ($major -eq 3 -and $minor -ge 11)) { $verOk = $true }
  }
} catch {}

if (-not $verOk) {
  if ($ver) {
    Write-Err "Python version too old. Need 3.11+ (prefer 3.12). Detected: $ver"
  } else {
    Write-Err "Could not detect Python version. Ensure 'py' or 'python' works and rerun."
  }
  exit 2
}
Write-Info "Python OK: $ver"

# ---- Run setup (venv + deps + playwright browser)
Write-Info "Running setup (00_Setup.ps1)..."
powershell -NoProfile -ExecutionPolicy Bypass -File ".\00_Setup.ps1"

# ---- Optional: login now
$loginNow = Read-Host "Run login now to create/update hpe_state.json? (Y/N)"
if ($loginNow -match '^(y|yes)$') {
  powershell -NoProfile -ExecutionPolicy Bypass -File ".\01_Login.ps1"
}

# ---- Optional: scheduled task
$taskNow = Read-Host "Create/update Scheduled Task to run every 10 minutes? (Y/N)"
if ($taskNow -match '^(y|yes)$') {
  $min = Read-Host "Minutes interval (default 10)"
  if ([string]::IsNullOrWhiteSpace($min)) { $min = 10 }
  if (-not ($min -as [int])) { $min = 10 }
  powershell -NoProfile -ExecutionPolicy Bypass -File ".\03_CreateTask_10min.ps1" -RootPath "$Target" -Minutes ([int]$min)
}

$runNow = Read-Host "Run bot once now (quick test, shows output)? (Y/N): "
if ([string]::IsNullOrWhiteSpace($runNow) -or $runNow -match '^(y|yes)$') {
  Write-Info "Running bot once..."
  $runPs1 = Join-Path $Target "Run-HPECaseBot.ps1"
  powershell -NoProfile -ExecutionPolicy Bypass -File "$runPs1" -Headless
}

Write-Info "Install complete."
Write-Host ""
Write-Host "Next:" -ForegroundColor Cyan
Write-Host "  1) Login (interactive):     .\01_Login.cmd   (CMD)  OR  powershell -ExecutionPolicy Bypass -File .\01_Login.ps1"
Write-Host "  2) Run bot once:            .\Run-HPECaseBot.cmd   (CMD)  OR  powershell -ExecutionPolicy Bypass -File .\Run-HPECaseBot.ps1"
Write-Host "  3) Output folder:           .\out_hpe"
