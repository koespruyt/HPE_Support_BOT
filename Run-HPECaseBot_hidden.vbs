' Run-HPECaseBot_hidden.vbs
' Doel: Run-HPECaseBot.ps1 volledig verborgen (geen console window) voor Scheduled Task.
' Versie: v6 (robuste quoting)

Option Explicit

Dim sh, fso, here, ps, runner, cmd, rc, q
Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

here   = fso.GetParentFolderName(WScript.ScriptFullName)
ps     = sh.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
runner = here & "\Run-HPECaseBot.ps1"
q      = Chr(34)

sh.CurrentDirectory = here
cmd = q & ps & q & " -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File " & q & runner & q & " -Headless"

On Error Resume Next
rc = sh.Run(cmd, 0, True)
If Err.Number <> 0 Then
  WScript.Quit 2
End If
WScript.Quit rc
