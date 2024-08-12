'=*=*=*=*=*=*=*=*=*=*=*=
' Coded By Assaf Miron 
'=*=*=*=*=*=*=*=*=*=*=*=

'This Script Disabels a List of Services from a text file

Const ForReading=1

Sub DisableService(ServiceName)
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colServiceList = objWMIService.ExecQuery _
    ("Select * from Win32_Service where DisplayName = " & ServiceName)
For Each objService in colServiceList
    errReturnCode = objService.StopService()    
    errReturnCode = objService.Change( , , , , "Disabled")   
Next
End Sub

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

set objReadFile = objFSO.OpenTextFile(FileLoc, ForReading)
Do Until objReadFile.AtEndOfStream
    Err = 0
    strNextLine = objReadFile.Readline
    arrServiceList = Split(strNextLine , ",")
    For i = 0 to Ubound(arrServiceList)

	DisableService arrServiceList(i)

    Next
Loop