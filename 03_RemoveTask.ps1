param(
  [string]$TaskName = "HPE CaseBot - Poll"
)

$ErrorActionPreference = "Stop"

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
  Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
  Write-Host "✅ Removed scheduled task: $TaskName"
} else {
  Write-Host "ℹ️ Scheduled task not found: $TaskName"
}
