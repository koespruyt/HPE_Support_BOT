#requires -Version 5.1

<#
.SYNOPSIS
  Wrapper om HPE CaseBot veilig te runnen (console of Scheduled Task).

.DESCRIPTION
  - Logt altijd naar:   .\out_hpe\run.log
  - Nagios statusfile: .\out_hpe\nagios\hpe_casebot.status
  - Archive:           .\out_hpe_archive\YYYY-MM-DD\

  TIP: vanuit CMD gebruik je best .\Run-HPECaseBot.cmd
#>

[CmdletBinding()]
param(
  [string]$OutDir = "",

  [ValidateSet('csv','json','both')]
  [string]$Format = "both",

  [switch]$Headless,

  [int]$Max = 0,

  [int]$TimeoutSeconds = 0,

  [switch]$NoArchive
)

$ErrorActionPreference = 'Stop'

# Script root (works from PowerShell, CMD wrapper, and Scheduled Task)
$script:ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }

function NowStr { (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }

function Ensure-Dir([string]$Path) {
  if ([string]::IsNullOrWhiteSpace($Path)) { return }
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Force -Path $Path | Out-Null
  }
}

function Append-Log([string]$Line) {
  $root = $script:ScriptRoot
  if ([string]::IsNullOrWhiteSpace($OutDir)) { $script:OutDir = Join-Path $root 'out_hpe' }

  Ensure-Dir $script:OutDir
  $log = Join-Path $script:OutDir 'run.log'
  Add-Content -LiteralPath $log -Encoding UTF8 -Value ("[{0}] {1}" -f (NowStr), $Line)
}

function Write-Status([string]$State, [string]$Message) {
  $root = $script:ScriptRoot
  if ([string]::IsNullOrWhiteSpace($OutDir)) { $script:OutDir = Join-Path $root 'out_hpe' }

  $nag = Join-Path $script:OutDir 'nagios'
  Ensure-Dir $nag
  $statusFile = Join-Path $nag 'hpe_casebot.status'

  $payload = @{
    generated_at = (Get-Date).ToString('o')
    state        = $State
    message      = $Message
  } | ConvertTo-Json -Depth 4

  Set-Content -LiteralPath $statusFile -Encoding UTF8 -Value $payload
  return $statusFile
}

function Archive-OutDir {
  $root = $script:ScriptRoot
  if ([string]::IsNullOrWhiteSpace($OutDir)) { $script:OutDir = Join-Path $root 'out_hpe' }

  $archiveRoot = Join-Path $root 'out_hpe_archive'
  $day = (Get-Date).ToString('yyyy-MM-dd')
  $dest = Join-Path $archiveRoot $day

  Ensure-Dir $archiveRoot
  Ensure-Dir $dest

  if (Test-Path -LiteralPath $script:OutDir) {
    robocopy $script:OutDir $dest /E /NFL /NDL /NJH /NJS /NC /NS | Out-Null
  }

  Append-Log "Archive to: $dest"
}

try {
  $Root = $script:ScriptRoot
  if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $OutDir = Join-Path $Root 'out_hpe'
  }

  Ensure-Dir $OutDir

  Append-Log "START Root=$Root OutDir=$OutDir"

  # Preconditions
  $setupMarker = Join-Path $Root '.setup.ok'
  if (-not (Test-Path -LiteralPath $setupMarker)) {
    Write-Status 'CRITICAL' "Setup not done. Run: powershell -ExecutionPolicy Bypass -File .\\00_Setup.ps1" | Out-Null
    throw "Setup not done. Run: powershell -ExecutionPolicy Bypass -File .\\00_Setup.ps1"
  }

  $py = Join-Path $Root '.venv\Scripts\python.exe'
  if (-not (Test-Path -LiteralPath $py)) {
    Write-Status 'CRITICAL' "Python venv missing. Run: powershell -ExecutionPolicy Bypass -File .\\00_Setup.ps1" | Out-Null
    throw "Python venv missing: $py"
  }

  $state = Join-Path $Root 'hpe_state.json'
  if (-not (Test-Path -LiteralPath $state)) {
    Write-Status 'CRITICAL' "Missing state file: hpe_state.json (run .\\01_Login.ps1)" | Out-Null
    throw "Missing state file: $state"
  }

  $selectors = Join-Path $Root 'hpe_selectors.json'
  if (-not (Test-Path -LiteralPath $selectors)) {
    Write-Status 'CRITICAL' "Missing selectors file: hpe_selectors.json" | Out-Null
    throw "Missing selectors file: $selectors"
  }

  $scriptPy = Join-Path $Root 'hpe_cases_overview.py'
  if (-not (Test-Path -LiteralPath $scriptPy)) {
    Write-Status 'CRITICAL' "Missing script: hpe_cases_overview.py" | Out-Null
    throw "Missing script: $scriptPy"
  }

  # Build args
  $pyArgs = @(
    $scriptPy,
    '--state', $state,
    '--selectors', $selectors,
    '--outdir', $OutDir,
    '--format', $Format
  )

  if ($Headless) { $pyArgs += '--headless' }
  if ($Max -gt 0) { $pyArgs += @('--max', "$Max") }
  if ($TimeoutSeconds -gt 0) { $pyArgs += @('--timeout', "$TimeoutSeconds") }

  Append-Log ("RUN {0} {1}" -f $py, ($pyArgs -join ' '))

  # Run Python and capture all output safely (no streaming errors -> no swallowed traceback)
  $prevEap = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  $raw = & $py @pyArgs 2>&1
  $exitCode = $LASTEXITCODE
  $ErrorActionPreference = $prevEap

  $lines = @()
  foreach ($o in $raw) {
    $s = [string]$o
    if ([string]::IsNullOrWhiteSpace($s)) { continue }
    foreach ($line in ($s -split "`r?`n")) {
      if ([string]::IsNullOrWhiteSpace($line)) { continue }
      $lines += $line
      Write-Host $line
      Append-Log $line
    }
  }

  if ($exitCode -ne 0) {
    $short = ($lines | Select-Object -First 1)
    if ([string]::IsNullOrWhiteSpace($short)) { $short = "Python exited with code $exitCode" }

    Write-Status 'CRITICAL' $short | Out-Null
    Append-Log "FAIL exit=$exitCode"

    if (-not $NoArchive) { Archive-OutDir }
    exit 1
  }

  # SUCCESS
  Write-Status 'OK' "OK" | Out-Null
  Append-Log "OK"

  if (-not $NoArchive) { Archive-OutDir }
  exit 0
}
catch {
  $Root = $script:ScriptRoot
  if ([string]::IsNullOrWhiteSpace($OutDir)) { $OutDir = Join-Path $Root 'out_hpe' }
  Ensure-Dir $OutDir
  Ensure-Dir (Join-Path $OutDir 'debug')

  $full = ($_ | Out-String).TrimEnd()
  if ([string]::IsNullOrWhiteSpace($full)) {
    $full = $_.Exception.ToString()
  }

  # Write full wrapper exception to debug file
  $dbg = Join-Path (Join-Path $OutDir 'debug') ('wrapper_exception_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.txt')
  Set-Content -LiteralPath $dbg -Encoding UTF8 -Value $full

  # Short Nagios message
  $short = $_.Exception.Message
  if ([string]::IsNullOrWhiteSpace($short)) { $short = "Wrapper exception" }
  Write-Status 'CRITICAL' $short | Out-Null

  Append-Log ("EXCEPTION - {0}" -f $short)
  Append-Log ("DEBUG - {0}" -f $dbg)

  if (-not $NoArchive) { Archive-OutDir }
  exit 99
}
