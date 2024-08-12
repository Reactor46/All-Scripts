'=*=*=*=*=*=*=*=*=*=*=*=
' Coded By Assaf Miron 
' 19/03/2006
' Update : 22/03/06
'=*=*=*=*=*=*=*=*=*=*=*=
On Error Resume Next

Const ALL_USERS = True
Const ForReading = 1
Const ForWriting = 2
intNOP = 0
intNOC = 0
intERR = 0

Sub CheckPerm(strComputer)
On Error Resume Next
Err = 0
set objFSO = CreateObject("Scripting.FileSystemObject")
Set objClass = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\cimv2:Win32_Service")

if Err.Number = -2147217405 then
  ObjFile.WriteLine "No Permission On " & StrComputer
  intNOP = intNOP + 1
  Exit Sub
End If
if Err.Number = 462 Then
  ObjFile.WriteLine "No Such Computer " & StrComputer
  intNOC = intNOC + 1
  Exit Sub
End If
If Err.Number = 0 Then Exit Sub

ObjFile.WriteLine "Unknown Error On " & StrComputer
  intERR = intERR + 1
End Sub

Sub InstMSI(strComputer,MSIFile)
Set objService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objSoftware = objService.Get("Win32_Product")
errReturn = objSoftware.Install(MSIFile,"SMSSITECODE=ABA DISABLESITEOPT=True DISABLECACHEOPT=True" , ALL_USERS)

if errReturn <> 0 then
	objFile.WriteLine errRutern.num & "On Computer - " & strComputer & vbcrlf & errReturn.description
End if

IF errReturn = 0 then
	objFile.WriteLine "The installation has completed succfully on compuer " & strComputer
End If
End Sub

Sub ClearSMS(strComputer,CCmClean)

Set objWMIService = GetObject("winmgmts:" _
      & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2:Win32_Process")
Set fso = CreateObject ("Scripting.FileSystemObject")

    objFSO.CopyFile CCmlean,"\\" & strComputer & "\Admin$", True
    E = objWMIService.Create("cmd /c %windir%\ccmclean.exe /q ", null, null, intProcessID)

End Sub

Set objDialog = CreateObject("UserAccounts.CommonDialog")
set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")
set objExcel=createobject("excel.application")


' Locating Msi File
wscript.echo "Locate MSI File"

objDialog.Filter = "Msi Files|*.msi"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
Else
    MSIFile = objDialog.FileName
End If

' Locating CCmClean File
wscript.echo "Locate CCmClean File"

objDialog.Filter = "EXE Files|*.exe"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
Else
    CCmClean = objDialog.FileName
End If

'Locate Computers File
wscript.echo "Locate Computers File"

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

fname= "SMS Instalation Log.txt"

If objFso.FileExists (fname) THEN
	set objFile = objFso.GetFile (fname)
	objFile.Delete
end If 
Set objFile = objFso.CreateTextFile (fname, True)

objFile.WriteLine "Log Started : " & Now
'Get List Of Computers

objexcel.workbooks.open(FileLoc)

IntRow = 1

Do Until objExcel.cells(introw,1).value=""
	
	strComputer = objExcel.cells(introw,1).value
	objFile.WriteLine "Computer: " & strComputer
	CheckPerm strComputer
	ClearSMS strComputer,CCmClean
 	InstMSI strComputer,MSIFile
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