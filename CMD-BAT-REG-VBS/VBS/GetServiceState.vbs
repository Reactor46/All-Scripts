'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' Http://assaf.miron.googlepages.com
' Date : 26/02/2009
' GetServiceState.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
Const SERVICE_NAME = "lanmanworkstation"
Const LOG_FILE = "C:\Service status.csv"

Set objDialog = CreateObject("UserAccounts.CommonDialog")
set objFSO = CreateObject("Scripting.FileSystemObject")
set objExcel=createobject("Excel.Application")

Sub GetServiceStatus (strComputer)
' This sub Checks the Service State and outputs it to the Log file
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
	Set colServiceList = objWMIService.ExecQuery _
	    ("Select * from Win32_Service where Name='" & SERVICE_NAME & "'")
	
	For Each objService in colServiceList
	    objFile.WriteLine strComputer & "," & objService.State
	Next
	
End Sub

'Locate Computers File (Excel or CSV File)
objDialog.Filter = "Excel Files|*.xls|CSV Files|*.csv"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\"
intResult = objDialog.ShowOpen
If intResult = 0 Then
    Wscript.Quit
Else
   FileLoc = objDialog.FileName
End If

' Open The Computers file for reading
objExcel.workbooks.open(FileLoc)

' Check if the Log File exists - Delete it to Create a new one
If objFso.FileExists (LOG_FILE) THEN
	set objFile = objFso.GetFile (LOG_FILE)
	objFile.Delete
End If 
' Create a new Log File
Set objFile = objFso.CreateTextFile (LOG_FILE, True)

IntRow = 1 ' Set to 2 if there is a Header in The Excel File

' Write the Headers on the Log File (For The CSV)
objFile.WriteLine "Computer Name,Service State"

' Loop on the Excel File until no Computers left in the First Column
Do Until objExcel.cells(introw,1).value=""
	strComputer = objExcel.cells(introw,1).value
	GetServiceStatus strComputer
	introw = introw+1
Loop

' Close and Cleanup
objExcel.Quit
objFile.Close
