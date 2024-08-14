On Error Resume Next
Const ForReading = 1
Const ForWriting = 2
'========================================================
Dim objLog, objFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Reads from C:\hostname.txt and writes to C:\temp\FileVer.txt
Set objFile = objFSO.OpenTextFile("c:\hostname.txt", ForReading)
Set objLog = objFSO.OpenTextFile("c:\temp\FileVer.txt", 8, True)
'=========================================================

'Reads the contents of C:\hostname.txt and stores in strContents
strContents = objFile.ReadAll
objFile.Close

' Creates an array arrLines and stores contents of C:\hostname.txt, splitting using Carriage Return
arrLines = Split(strContents, vbCrLf)

For i = 0 to UBound(arrLines)
strComputer = arrLines(i)
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("Select * from CIM_Datafile Where Name = 'C:\\Program Files\\Internet Explorer\\iexplore.exe'")

For Each objFile in colFiles
    objLog.WriteLine  strComputer & chr(32) & objFile.Version
Next
Next


