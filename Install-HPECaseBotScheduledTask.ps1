<#
Install-HPECaseBotScheduledTask.ps1
Creates/updates a daily scheduled task that runs Run-HPECaseBot.ps1.

IMPORTANT:
- Do setup first:  .\00_Setup.ps1
- Do login once:  .\01_Login.ps1
Then install the task.

Usage:
  powershell.exe -ExecutionPolicy Bypass -File .\Install-HPECaseBotScheduledTask.ps1 -Time "07:00"
#>

[CmdletBinding()]
param(
  [string]$TaskName = "HPE CaseBot Daily",
  [string]$Time = "07:00",
  [switch]$RunHighest = $true
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$runner = Join-Path $root "Run-HPECaseBot.ps1"
if (-not (Test-Path $runner)) { throw "Run-HPECaseBot.ps1 not found in $root" }

$vbs = Join-Path $root "Run-HPECaseBot_hidden.vbs"

# Always (re)write hidden VBS wrapper (robust quoting) so upgrades overwrite older/broken versions.
$VbsContent = @'
' Run-HPECaseBot_hidden.vbs
' Doel: Run-HPECaseBot.ps1 volledig verborgen (geen console window) voor Scheduled Task.
' Versie: v5 (robuste quoting)

Option Explicit

Dim sh, fso, here, ps, runner, cmd, rc, q
Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

here   = fso.GetParentFolderName(WScript.ScriptFullName)
ps     = sh.ExpandEnvironmentStrings("%SystemRoot%") & "\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"
runner = here & "\\Run-HPECaseBot.ps1"
q      = Chr(34)

sh.CurrentDirectory = here
cmd = q & ps & q & " -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File " & q & runner & q

On Error Resume Next
rc = sh.Run(cmd, 0, True)
If Err.Number <> 0 Then
  WScript.Quit 2
End If
WScript.Quit rc
'@

Set-Content -LiteralPath $vbs -Encoding ASCII -Value $VbsContent

$wscript = "$env:SystemRoot\System32\wscript.exe"
$arg = "//B //NoLogo `"$vbs`""

Write-Host "TaskName: $TaskName"
Write-Host "Time:     $Time"
Write-Host "Action:   $wscript $arg"
Write-Host ""

try {
  Import-Module ScheduledTasks -ErrorAction Stop
  $action    = New-ScheduledTaskAction -Execute $wscript -Argument $arg -WorkingDirectory $root
  $trigger   = New-ScheduledTaskTrigger -Daily -At ([datetime]::Parse($Time))
  $userId    = "$env:USERDOMAIN\$env:USERNAME"
  $principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel (if($RunHighest){"Highest"}else{"Limited"})
  $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2)
  $task      = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings
  Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null
  Write-Host "✅ Scheduled task created/updated."
}
catch {
  Write-Warning "ScheduledTasks module not usable. Falling back to schtasks.exe"
  $rl = if($RunHighest){"/RL HIGHEST"}else{""}

  # schtasks.exe expects the entire command line as ONE /TR argument.
  # Put wscript + its switches + vbs path into a single string.
  $tr = "`"$wscript`" //B //NoLogo `"$vbs`""

  $args = @("/Create","/F","/SC","DAILY","/ST",$Time,"/TN",$TaskName,"/TR",$tr)
  if ($RunHighest) { $args += @("/RL","HIGHEST") }

  Write-Host ("Running: schtasks.exe " + ($args -join " "))
  $out = & schtasks.exe @args 2>&1
  $code = $LASTEXITCODE
  if ($code -ne 0) {
    throw ("schtasks.exe failed (exit {0}). Output: {1}" -f $code, ($out -join "`n"))
  }

  Write-Host "✅ Scheduled task created/updated (schtasks)."
}
