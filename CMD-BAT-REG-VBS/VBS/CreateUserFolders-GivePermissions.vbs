'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' http://sites.google.com/site/assafmiron/
' Date : 24/01/2010
' CreateUserFolders-GivePermissions.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
Option Explicit
On Error Resume Next

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim objFSO, objFile
Dim objShell, objDialog, objExcel
Dim LogFile, strServerPath, strUser, FileLoc
Dim intResult, intRow

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = createObject("WScript.Shell")
' Disalog Works on XP and Above with WSH 5.6 - Does not work on Server 2003
Set objDialog = CreateObject("UserAccounts.CommonDialog")
set objExcel=createobject("Excel.Application")

' Set the Server's Path
strServerPath = "\\Server\UserData\"
LogFile = "MyLogFile.txt" ' Set the Log File Path


Sub CreateFolder (strUserName, bFCAdminRights)
' Function will Create a User Folder and give it the appropriate Permissions
' Input  : strUser - User Name wich will be the Folder Name, and Will give Full Control to that User from Active Directory
'		   bFCAdminRights - Boolean, Give Administrators Full Controll on the Directory? (True/False)
' Output : NONE
	On Error Resume Next
	' Create the Folder
	If objFSO.FileExists(strServer & strUserName)=False Then
		objFSO.CreateFolder(strServerPath & strUserName)
	End If
	
	' Give Permissions
	If bFCAdminRights Then
		objShell.run "cmd /C xcacls " & strServerPath & "\" & strUserName & " /g " & strUserName & "-R:f administrators:f /t /y /c", 0
	Else
		objShell.run "cmd /C xcacls " & strServerPath & "\" & strUserName & " /g " & strUserName & "-R:f /t /y /c", 0
	End If
End Sub

'-----------
' Code Start
'-----------

' Dialog Works on XP and Above with WSH 5.6 - Does not work on Server 2003
objDialog.Filter = "Excel Files|*.xls|CSV Files|*.csv|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
Else
    FileLoc = objDialog.FileName
End If

' For Server Version Use This:
' FileLoc = InputBox("Enter User File Name", "Enter User File Name","C:\UsersFile.xls")

' Create the Log File
If objFSO.FileExists(LogFile) Then
	Set objFile = objFSO.OpenTextFile(LogFile,ForAppending) ' Appending The Log File
Else
	Set objFile = objFSO.CreateTextFile(LogFile,ForWriting) ' Creating The Log File
End If

' Open the Excel File
objexcel.workbooks.open(FileLoc)
intRow=2 ' Set 2 to Skip Header Line

' Start the Log File
ObjFile.WriteLine "Script Started on " & Now
ObjFile.WriteLine "File Path: " & FileLoc

' Create the Folders from the File
Do Until objExcel.Cells(IntRow,1).Value=""
	strUser = objExcel.Cells(IntRow,1).Value
	objFile.WriteLine "Creating User Folder: " & strUser
	' Create the User Folder with Admin Full Controll Permissions
	CreateFolder strUser, True
	intRow=intRow+1
Loop

' Close the Excel File
objExcel.Workbooks.Close

' Close the Log File
objFile.Close

' Clean up
Set objFile = Nothing
Set objFSO = Nothing
wscript.echo "Done"