'=*=*=*=*=*=*=*=*=*=*=*=*=
' Author : Assaf Miron 
' Http://assaf.miron.googlepages.com
' Date : 09/08/10
' ShowWorkSheetNames.vbs
' Description : Echos the Sheet Name from Excel
'=*=*=*=*=*=*=*=*=*=*=*=*=
'Option Explicit
'On Error Resume Next

Dim objExcelIn, objDialog, objWorkbook
Dim objWorksheet, WS
Dim SheetName, SheetID
Dim FileLoc

Set objExcelIn  = CreateObject("Excel.Application")
Set objDialog = CreateObject("UserAccounts.CommonDialog")

' Open a File Using a Open Dialog Box
objDialog.Filter = "Excel Files|*.xls|CSV Files|*.csv|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\"
intError = objDialog.ShowOpen
 
If intError = 0 Then
    Wscript.Quit
Else
    FileLoc = objDialog.FileName
End If

' Read an Excel Spreadsheet
Set objWorkbook = objExcelIn.Workbooks.Open(FileLoc)

WScript.Echo objExcelIn.Worksheets.Count

'For Each WS In objExcelIn.Worksheets
'	WScript.Echo WS.Name
'Next

SheetID = "1"
' Activate the WorkSheet in the Input Excel
Set objWorksheet = objExcelIn.Worksheets(Int(SheetID))
' Activate Each WorkSheet and Read the Data from it
ObjWorkSheet.Activate
SheetName = ObjWorkSheet.Name

WScript.Echo SheetName
objExcelIn.Quit