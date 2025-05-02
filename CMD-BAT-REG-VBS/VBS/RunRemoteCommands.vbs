'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' Http://assaf.miron.googlepages.com
' Date : 10/03/2009
' RunRemoteCommands.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' This Script Runs any Command on a List of Remote Computers.
' The Computer List is Retreived from an Excel File, and a Log File is saved
' With Details on Each Computer.
' Examples for Commads :
' Const strCommand = "msiexec.exe /x {35C03C04-3F1F-42C2-A989-A757EE691F65} REMOVE=ALL REBOOT=R /q"
' Const strCommand = "ShutDown -t 10 -c " & chr(34) & "This Computer Is Restarting Now" & chr(34) & " -r"
' Const strCommand = "Net LocalGroup Administrators > C:\LocalAdmins.txt "
On Error Resume Next

Const strCommand = "Echo Hello World, This is Computer & HostName"
Const LOG_FILE = "C:\RemoteCommand.txt"
Const ForReading = 1
Const ForWriting = 2
intNOP = 0
intNOC = 0
intERR = 0

Function CheckPerm(strComputer)
On Error Resume Next
Err = 0
set objFSO = CreateObject("Scripting.FileSystemObject")
Set objClass = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\cimv2:Win32_Service")

if Err.Number = -2147217405 then
  ObjFile.WriteLine "No Permission On " & StrComputer
  intNOP = intNOP + 1
  CheckPerm = 1
  Exit Function
End If
if Err.Number = 462 Then
  ObjFile.WriteLine "No Such Computer " & StrComputer
  intNOC = intNOC + 1
  CheckPerm = 1
  Exit Function
End If
If Err.Number = 0 Then
  CheckPerm = 0
  Exit Function
End If

ObjFile.WriteLine "Unknown Error On " & StrComputer
  intERR = intERR + 1
  CheckPerm = 1
  Exit Function
End Function

Sub RunCommand(strComputer,Command)

Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2:Win32_Process")
Err = objWMIService.Create("cmd /c " & Command, null, null, intProcessID)
If Err.Number = 0 Then
	objFile.WriteLine "Command Has begun on " & strComputer
Else
	objFile.WriteLine "Command Failed to run on " & strComputer
End If

End Sub

Set objDialog = CreateObject("UserAccounts.CommonDialog")
set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
set objExcel=createobject("Excel.Application")


'Locate Computers File

objDialog.Filter = "Excel Files|*.xls|CSV Files|*.csv"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
Else
   FileLoc = objDialog.FileName
End If

'Output Location And Name

If objFso.FileExists (LOG_FILE) THEN
	Set objFile = objFso.GetFile (LOG_FILE)
	objFile.Delete
end If 
Set objFile = objFso.CreateTextFile (LOG_FILE, True)

objFile.WriteLine "Log Started : " & Now
'Get List Of Computers

objexcel.workbooks.open(FileLoc)

IntRow = 1

Do Until objExcel.cells(introw,1).value=""
	
	strComputer = objExcel.cells(introw,1).value
	objFile.WriteLine "Computer: " & strComputer
	If CheckPerm(strComputer) = 0 Then
		RunCommand strComputer,strCommand
	End If
	IntRow = IntRow + 1
Loop

objFile.writeLine
objFile.WriteLine "Summary :"
objFile.WriteLine "Sum All Computers :" & intRow-1
objFile.WriteLine "Sum All Computers With No Permmision :" & intNOP
objFile.WriteLine "Sum All Computers With Unknown Error :" & intERR
objFile.WriteLine "Sum All Computers that Dont exist :" & intNOC
objFile.writeLine
objFile.WriteLine "The Script Ended : " & Now
objFile.Close
objExcel.WorkBook.Close
objExcel.Close
wscript.echo "Done !"