'=*=*=*=*=*=*=*=*=*=*=*=
' Coded By Assaf Miron 
' Date : 21/03/06
'=*=*=*=*=*=*=*=*=*=*=*=

On Error Resume Next
Const ForReading = 1

Function CheckPerm(strComputer)
On Error Resume Next
Err = 0
set objFSO = CreateObject("Scripting.FileSystemObject")
Set objClass = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\cimv2:Win32_Service")

if Err.Number = -2147217405 then
  ObjFile.WriteLine "No Permission On " & StrComputer
  CheckPerm = 2
  Exit Function
End If
if Err.Number = 462 Then
  ObjFile.WriteLine "No Such Computer " & StrComputer
  CheckPerm = 1
  Exit Function
End If
If Err.Number = 0 Then   CheckPerm = 0
ObjFile.WriteLine "Unknown Error On " & StrComputer
CheckPerm = 3
End Function

Function ShutDown(strComputer)
Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2:Win32_Process")

E = objWMIService.Create("cmd /c shutdown -r -t 30 ", null, null, intProcessID)

Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'OUTLOOK.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next

Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'Excel.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next


Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'WinWORD.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next


Set colProcessList = objWMIService.ExecQuery _
    ("Select * from Win32_Process Where Name = 'PowerPnt.exe'")
For Each objProcess in colProcessList
    objProcess.Terminate()
Next

wscript.sleep 30000
Set WMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\" & strComputer & "\root\cimv2")
	Set OperatingSystems = WMIService.ExecQuery("Select * From Win32_OperatingSystem")
	For Each OperatingSystem in OperatingSystems
		OperatingSystem.Reboot()
	Next
if E = 0 then
  ShutDown  = "OK"
end if

End Function

' Opening File

Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "Text Files|*.txt|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
Else
    FileLoc = objDialog.FileName
End If

'Output Location And Name
Set WshShell = WScript.CreateObject("WScript.Shell")
fname="Boot Log.txt"
set fso = CreateObject ("Scripting.FileSystemObject")
If fso.FileExists (fname) THEN
	set objFile = fso.GetFile (fname)
	objFile.Delete
end If 
Set objFile = fso.CreateTextFile (fname, True)

objFile.WriteLine "Log Started : " & Now
objFile.WriteLine

'Get List Of Computers
set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
set objReadFile = objFSO.OpenTextFile(FileLoc, ForReading)
Do Until objReadFile.AtEndOfStream
    Err = 0
    strNextLine = objReadFile.Readline
    arrComputerList = Split(strNextLine , ",")
    For i = 0 to Ubound(arrComputerList)
        a=0
	strComputer = arrComputerList(i)
	If CheckPerm(strComputer) = 0 Then
		If ShutDown(strComputer) = "OK" Then
		  objFile.WriteLine "Computer: " & arrComputerList(i) & " Booted OK"
		Else
		  objFile.WriteLine "Computer: " & arrComputerList(i) & " Did not Boot"
		End If
	End If
    Next
Loop

objReadFile.Close
objFile.WriteLine
objFile.WriteLine "Log Ended : " & Now
objFile.Close
wscript.echo "Done Booting All Computers!"