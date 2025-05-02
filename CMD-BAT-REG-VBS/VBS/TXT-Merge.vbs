Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutputFile = objFSO.CreateTextFile("output.txt")

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set FileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='z:\Scripts\Test'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In FileList
    Set objTextFile = objFSO.OpenTextFile(objFile.Name, ForReading) 
    strText = objTextFile.ReadAll
    objTextFile.Close
    objOutputFile.WriteLine strText
Next

objOutputFile.Close