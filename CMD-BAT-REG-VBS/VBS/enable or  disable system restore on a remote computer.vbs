'Enable or disable system restore on a remote system

strComputer = InputBox("Enter the COMPUTER NAME or the IP address of the system you would like to enable or disable System Restore on. (use localhost for current system)","Remotely Enable or Disable System Restore","localhost")

If strComputer = "" Then
  WScript.Quit
End If

strChoice = Inputbox ("System Restore:" & vbCrLf & vbCrLf & "1: Enable"& vbCrLf &"2: Disable"  & vbCrLf& vbCrLf & vbCrLf & "Please enter 1 or 2 (leave blank to quit):","Enter 1 to Enable or 2 to Disable","")


If strChoice = "" Then
  WScript.Quit
End If


if strChoice = 1 then 
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")

Set objItem = objWMIService.Get("SystemRestore")
errResults = objItem.Enable("")

'Inform User that the task is done.
Mybox = MsgBox("System Restore is now enabled on "& strComputer &""  & vbCRLF ,vbOkOnly,"System Restore Enabled")
End If


if strChoice = 2 then 
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")

Set objItem = objWMIService.Get("SystemRestore")
errResults = objItem.Disable("")

'Inform User that the task is done.
Mybox = MsgBox("System Restore is now disabled on "& strComputer &""  & vbCRLF ,vbOkOnly,"System Restore Disabled")
End If