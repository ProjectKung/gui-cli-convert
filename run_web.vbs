Option Explicit

Dim fso, shell, scriptDir, command
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
command = "powershell -NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File """ _
    & scriptDir & "\run_web_launcher.ps1"" -AutoInstallPython"

' Run in hidden console mode, but show the custom loading window from launcher script.
shell.Run command, 0, False
