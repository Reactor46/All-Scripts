On Error Resume Next

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objGetComputerList = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("nicspeed.txt", True)
Set fsoReadComputerList = objGetComputerList.OpenTextFile("computers.txt", 1, TristateFalse)
aryComputers = Split(fsoReadComputerList.ReadAll, vbCrLf)
fsoReadComputerList.Close

const HKEY_LOCAL_MACHINE = &H80000002
For Each strComputer In aryComputers
	'strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
 strComputer & "\root\default:StdRegProv")
 
strKeyPath = "System\Currentcontrolset\Control\Class\{4D36E972-E325-11CE-BFC1-08002be10318}\0007"
strValueName = "*SpeedDuplex"
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

objFile.WriteLine "========================================="
objFile.WriteLine "Host Name:" & strComputer
if strValue = 0 Then objFile.WriteLine "Current NIC Speed: Auto"
if strValue = 1 Then objFile.WriteLine "Current NIC Speed: 10 Half"
if strValue = 2 Then objFile.WriteLine "Current NIC Speed: 10 Full"
if strValue = 3 Then objFile.WriteLine "Current NIC Speed: 100 Half"
if strValue = 4 Then objFile.WriteLine "Current NIC Speed: 100 Full"

next