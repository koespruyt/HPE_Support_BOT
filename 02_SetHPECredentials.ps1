#requires -Version 5.1
<#
.SYNOPSIS
  Store HPE portal credentials encrypted for the CURRENT Windows user (DPAPI).

.DESCRIPTION
  Creates/updates: .\hpe_credential.xml
  - Uses Export-Clixml which encrypts the SecureString using DPAPI (current user)
  - Locks down ACL to: current user + SYSTEM + Administrators
  - Provides GUI credential prompt when possible, otherwise console fallback.

.NOTES
  Scheduled Task MUST run as the SAME Windows user to decrypt DPAPI.
#>

[CmdletBinding()]
param(
  [string]$OutFile,
  [switch]$Force
)

$ErrorActionPreference = "Stop"

# Robust script directory detection (works even when $PSScriptRoot is empty)
$scriptPath = $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($scriptPath)) {
  throw "Cannot determine script path. Start this script with: powershell -File .\02_SetHPECredentials.ps1"
}
$here = Split-Path -Parent $scriptPath
if ([string]::IsNullOrWhiteSpace($OutFile)) {
  $OutFile = Join-Path $here "hpe_credential.xml"
}

function Set-SafeAcl {
  param([Parameter(Mandatory)][string]$Path)

  $domain = $env:USERDOMAIN
  $name   = $env:USERNAME
  if ([string]::IsNullOrWhiteSpace($domain)) { $domain = $env:COMPUTERNAME }

  $user = "{0}\{1}" -f $domain, $name

  # SIDs (language independent)
  $sidSystem = "*S-1-5-18"      # SYSTEM
  $sidAdmins = "*S-1-5-32-544"  # BUILTIN\Administrators

  $args = @(
    $Path,
    "/inheritance:r",
    "/grant:r", ("{0}:(F)" -f $user),
    "/grant:r", ("{0}:(F)" -f $sidSystem),
    "/grant:r", ("{0}:(F)" -f $sidAdmins)
  )

  & icacls @args | Out-Null
}

function Prompt-CredConsole {
  Write-Host "No credential popup available/used. Falling back to console prompts..." -ForegroundColor Yellow
  $u = Read-Host "HPE Username (e-mail or user)"
  $p = Read-Host "HPE Password" -AsSecureString
  return New-Object System.Management.Automation.PSCredential($u, $p)
}

function Get-HpeCredential {
  # Try GUI prompt first, BUT also fallback if it returns $null/empty (cancelled or UI not available)
  try {
    $c = Get-Credential -Message "Enter your HPE portal username + password (no MFA). Stored DPAPI-encrypted for current user."
    if ($null -ne $c -and -not [string]::IsNullOrWhiteSpace($c.UserName)) { return $c }
    return (Prompt-CredConsole)
  } catch {
    return (Prompt-CredConsole)
  }
}

Write-Host ("This will store HPE credentials encrypted for user: {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME) -ForegroundColor Cyan
Write-Host ("Output: {0}" -f $OutFile) -ForegroundColor Cyan

if (Test-Path -LiteralPath $OutFile) {
  if (-not $Force) {
    $ans = Read-Host "File already exists. Overwrite? (Y/N)"
    if ($ans -notin @("Y","y")) {
      Write-Host "Aborted. No changes made." -ForegroundColor Yellow
      exit 0
    }
  }
  Remove-Item -LiteralPath $OutFile -Force
}

$cred = Get-HpeCredential

# Validate (must contain username + non-empty password)
if (-not $cred -or [string]::IsNullOrWhiteSpace($cred.UserName)) {
  throw "No valid username entered. Aborting."
}
$pwLen = $cred.GetNetworkCredential().Password.Length
if ($pwLen -lt 1) {
  throw "Empty password entered. Aborting."
}

# Save (DPAPI current user)
$cred | Export-Clixml -Path $OutFile

# Harden permissions
Set-SafeAcl -Path $OutFile

# Verify (do NOT print password)
$c = Import-Clixml $OutFile
$verifyLen = $c.GetNetworkCredential().Password.Length

Write-Host ("OK. Credential saved (DPAPI encrypted): {0}" -f $OutFile) -ForegroundColor Green
Write-Host ("UserName in file: {0}" -f $c.UserName) -ForegroundColor Green
Write-Host ("Password length:  {0}" -f $verifyLen) -ForegroundColor Green
Write-Host "Tip: Scheduled Task must run as SAME Windows user to decrypt DPAPI." -ForegroundColor Cyan
