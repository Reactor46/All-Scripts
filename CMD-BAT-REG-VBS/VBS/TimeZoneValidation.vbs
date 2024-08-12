On Error Resume Next
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("output.txt", True)
Set objTextFile = objFSO.OpenTextFile("computersprod.txt", ForReading)
Do Until objTextFile.AtEndOfStream 
    strComputer = objTextFile.Readline
const HKEY_LOCAL_MACHINE = &H80000002
Set oReg=GetObject( _
   "winmgmts:{impersonationLevel=impersonate}!\\" &_
    strComputer & "\root\default:StdRegProv")
strKeyPath = "System\CurrentControlSet\Control\TimeZoneInformation"
strValueName = "DayLightName"
oReg.GetStringValue _
   HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
objFile.WriteLine "Server: " & strComputer & " Value: " & strValue
Loop
objTextFile.Close

