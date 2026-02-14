<#
00_UnblockFiles.ps1
If you downloaded the ZIP from the internet, Windows may block scripts (Mark-of-the-Web).
Run this ONCE after extracting the ZIP:

  cd C:\hpe-casebot\FINAL
  powershell -ExecutionPolicy Bypass -File .\00_UnblockFiles.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference="Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

Get-ChildItem -Path $root -Recurse -File | Unblock-File -ErrorAction SilentlyContinue
Write-Host "✅ Unblocked all files under: $root"
