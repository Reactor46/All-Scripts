Option Explicit
Dim answer
Dim args
Dim script_path

script_path = "C:\Documents and Settings\E\Desktop\Powershell\boinc_task_saver.ps1"

answer = MsgBox("Click 'YES' to backup and shutdown" + vbCrLf + "Click 'NO' to backup and continue" + vbCrLf + "Click 'CANCEL' to do nothing", _
                  vbYesNoCancel, "Boinc Task Saver")

If answer = vbNo Then
  args = ""
ElseIf answer = vbYes Then
  args = " -shutdown"  
Else
  WScript.Quit  
End if

CreateObject("Wscript.Shell").Run "powershell.exe -noprofile -noexit -WindowStyle Hidden . '" & script_path & "'" & args, 0, False