'TITLE: Enumerate Software 
'CREATED: 01/02/2012
'AUTHOUR: Darren Mayes
'CHANGES: DM | 01/02/2012 | - UPDATED CODE TO UTILISE PROMPT BOX FOR REMOTE ENUMERATION

'Pop-up box to inform you that the script is running
WScript.Echo "The Script has started"

' Declare Variables
Dim dTtle
dTtle = "Enumerate Software"
Dim dHost
'Provide prompt to specify audit system
dHost = InputBox("Enter the I.P. or the computer " & _
                       "you would like to check the installed software " & _
                       "on." & vbcrlf & vbcrlf & "Remote enumeration " & _
                       "will be performed in the context of the current " & _
                       "logged on user", dTtle)
                       
If IsEmpty(dHost) Then WScript.Quit
dHost = Trim(dHost)
If dHost = "" Then dHost = "."
                    
'Provide prompt to specify save location & further declaration
Dim dFdl
dFdl = InputBox("Enter the pathname to output the details of this audit " , dTtle)
Dim dFle
dFle = dFdl & dHost & ".xls"
 
If IsEmpty(dFdl) Then WScript.Quit

'Create & launch Excel workbook   
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)
x = 2

objExcel.Cells(1, 1).Value = "Name"
objExcel.Cells(1, 2).Value = "Vendor"
objExcel.Cells(1, 3).Value = "Version"
objExcel.Cells(1, 4).Value = "InstallLocation"
objExcel.Cells(1, 5).Value = "Description"

'Access the required WMI namespace to perform query
Set objWMIService = _
    GetObject("winmgmts:\\" & dHost & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_Product")
'Perform FOR EACH loop and update newly created spreadsheet
For Each objItem in colItems
    objWorksheet.Cells(x, 1) = objItem.Name
    objWorksheet.Cells(x, 2) = objItem.Vendor
    objWorksheet.Cells(x, 3) = objItem.Version 
    objWorksheet.Cells(x, 4) = objItem.InstallLocation
    objWorksheet.Cells(x, 5) = objItem.Description    
    x = x + 1
Next

Set objRange = objWorksheet.UsedRange
objRange.EntireColumn.Autofit()

'Save & close updated spreadsheet.
objExcel.ActiveWorkbook.SaveAs dFle
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit

WScript.Echo "Software Enumeration Complete."
WScript.Quit








