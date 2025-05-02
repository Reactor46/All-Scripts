'=*=*=*=*=*=*=*=*=*=*=*=*=
' Created By : Assaf Miron
' Date : 14/02/2007
'=*=*=*=*=*=*=*=*=*=*=*=*=
On Error Resume Next
Const MY_DOMAIN = "MyDomain.com"
Sub CheckPerm(strComputer)

On Error Resume Next
Err = 0
set objFSO = CreateObject("Scripting.FileSystemObject")
Set objClass = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\cimv2:Win32_Service")

if Err.Number = -2147217405 then
  ObjFile.WriteLine "No Permission On " & StrComputer
  ObjFile.WriteLine "<br>"
  Exit Sub
End If
if Err.Number = 462 Then
  ObjFile.WriteLine "No Such Computer " & StrComputer
  ObjFile.WriteLine "<br>"  
  Exit Sub
End If
If Err.Number = 0 Then Exit Sub
  ObjFile.WriteLine "Unknown Error On " & StrComputer
  ObjFile.WriteLine "<br>"

End Sub

Sub AddAdmin(User,Computer)

strComputer = Computer
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2:Win32_Process")

Error = objWMIService.Create("cmd /c net localgroup administrators /add " & MY_DOMAIN & "\" & User , null, null, intProcessID)
If Error = 0 Then
    objFile.WriteLine "<p>The User " & User & " was added To Admin Group On Computer " & Computer & ".</p><br>"
Else
    objFile.WriteLine "<p>Error User: " & User & " " & Error & ".</p><br>"
End If
	objFile.WriteLine "<br><br>"
End Sub

Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "Excel Files|*.xls|CSV Files|*.csv|All Files|*.*"
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
fname=WshShell.SpecialFolders("Desktop")& "\" & "DisplayLog.html"
set fso = CreateObject ("Scripting.FileSystemObject")
If fso.FileExists (fname) THEN
	set objFile = fso.GetFile (fname)
	objFile.Delete
end If 
Set objFile = fso.CreateTextFile (fname, True)

set objExcel=createobject("excel.application")
objexcel.workbooks.open(FileLoc)
intRow=2

'write the HTML
	objFile.WriteLine "<html>"
	objFile.WriteLine "<head>"
	objFile.WriteLine "<meta http-equiv='Content-Language' content='he'>"
	objFile.WriteLine "<meta http-equiv='Content-Type' content='text/html; charset=windows-1255'>"
	objFile.WriteLine "<title>Admin</title>"
	objFile.WriteLine "</head>"
	objFile.WriteLine "<body>"

 do while objexcel.cells(introw,1).value <> ""
 	' Excel File needs to have to columns - First with the user name, Socond the computer name
 	' Don't leave empty lines
	User=objexcel.cells(introw,1).value
	Computer=objexcel.cells(introw,2).value
	CheckPerm(Computer)
	AddAdmin User,Computer

	introw=introw+1

 loop
objexcel.workbooks.close

objFile.WriteLine "</body>"
objFile.WriteLine "</html>"
objFile.Close

wscript.echo "done !"