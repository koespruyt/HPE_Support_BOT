<#
Uninstall-HPECaseBotScheduledTask.ps1
#>
[CmdletBinding()]
param(
  [string]$TaskName = "HPE CaseBot Daily"
)
$ErrorActionPreference = "Stop"
try {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop | Out-Null
  Write-Host "✅ Removed scheduled task: $TaskName"
}
catch {
  # fallback
  cmd.exe /c "schtasks /Delete /F /TN `"$TaskName`"" | Out-Null
  Write-Host "✅ Removed scheduled task: $TaskName (schtasks fallback)"
}
