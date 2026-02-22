param(
  [string]$RootPath = "",
  [string]$TaskName = "HPE CaseBot (10min)",
  [int]$Minutes = 10
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($RootPath)) {
  $RootPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$Root   = (Resolve-Path $RootPath).Path
$Script = Join-Path $Root "Run-HPECaseBot.ps1"
$Vbs    = Join-Path $Root "Run-HPECaseBot_hidden.vbs"

if (!(Test-Path $Script)) { throw "Missing Run-HPECaseBot.ps1 in $Root" }

# Always (re)write hidden VBS wrapper so upgrades overwrite older/broken versions.
# Note: VBScript does NOT use backslash escaping for quotes; we build the command with Chr(34).
$VbsContent = @'
' Run-HPECaseBot_hidden.vbs
' Doel: Run-HPECaseBot.ps1 volledig verborgen (geen console window) voor Scheduled Task.
' Versie: v8 (timeout + robust)

Option Explicit

Dim sh, fso, here, ps, runner, cmd, rc, q
Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

here   = fso.GetParentFolderName(WScript.ScriptFullName)
ps     = sh.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
runner = here & "\Run-HPECaseBot.ps1"
q      = Chr(34)

sh.CurrentDirectory = here
cmd = q & ps & q & " -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File " & q & runner & q & " -Headless -TimeoutSeconds 900"

On Error Resume Next
rc = sh.Run(cmd, 0, True)
If Err.Number <> 0 Then
  WScript.Quit 2
End If
WScript.Quit rc
'@

Set-Content -LiteralPath $Vbs -Encoding ASCII -Value $VbsContent

$WScript = "$env:SystemRoot\System32\wscript.exe"
$UserId  = "$env:USERDOMAIN\$env:USERNAME"

# Build a task XML with a CalendarTrigger + Repetition (most compatible way).
# IMPORTANT: schtasks.exe expects the XML file encoding to match the XML header.
# We write UTF-16LE (PowerShell: -Encoding Unicode) and declare UTF-16.
$startBoundary = (Get-Date).AddMinutes(1).ToString("yyyy-MM-dd'T'HH:mm:ss")
$intervalIso   = "PT{0}M" -f $Minutes

# XML-escape any values that could contain '&' etc.
$esc = { param([string]$s) [System.Security.SecurityElement]::Escape($s) }
$RootXml = & $esc $Root
$VbsXml  = & $esc $Vbs
$CmdXml  = & $esc $WScript
$Author  = & $esc $UserId

$xml = @"
<?xml version='1.0' encoding='UTF-16'?>
<Task version='1.4' xmlns='http://schemas.microsoft.com/windows/2004/02/mit/task'>
  <RegistrationInfo>
    <Author>$Author</Author>
    <Description>HPE CaseBot - run every $Minutes minutes (hidden via VBS)</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$startBoundary</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
      <Repetition>
        <Interval>$intervalIso</Interval>
        <Duration>P1D</Duration>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id='Author'>
      <UserId>$Author</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT15M</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context='Author'>
    <Exec>
      <Command>$CmdXml</Command>
      <Arguments>//B //NoLogo &quot;$VbsXml&quot;</Arguments>
      <WorkingDirectory>$RootXml</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

$tmp = Join-Path $env:TEMP ("hpe_casebot_{0}.xml" -f ([guid]::NewGuid().ToString("N")))
Set-Content -LiteralPath $tmp -Encoding Unicode -Value $xml

try {
  $out = & schtasks.exe /Create /TN $TaskName /XML $tmp /F 2>&1
  $code = $LASTEXITCODE
  if ($code -ne 0) {
    $msg = ($out | Out-String).Trim()
    throw "schtasks.exe failed (exit $code). Output: $msg"
  }
}
finally {
  Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}

Write-Host "Scheduled Task created/updated: $TaskName"
Write-Host "  Root:     $Root"
Write-Host "  Interval: every $Minutes minutes (starts ~1 minute from now)"
Write-Host ('To remove: schtasks /Delete /TN "{0}" /F' -f $TaskName)
