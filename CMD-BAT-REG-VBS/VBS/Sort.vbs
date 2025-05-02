'On Error Resume Next

Dim objShell
Set objShell = wscript.CreateObject("WScript.Shell")
Dim oFSO, oFile, oFSO2, oFile2
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oFSO2 = CreateObject("Scripting.FileSystemObject")
Dim destPath, unitList, unit, groupList, group, output, logFile
Set wshShell = WScript.CreateObject("WScript.Shell")
set wshFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set wshEnv = WshShell.Environment("Process")

unitList = "units.txt"
groupList = "groups.txt"
logfile = "product.csv"

If oFSO.FileExists(unitList) Then
	Set oFile = oFSO.OpenTextFile(unitList, 1)
		Do While Not oFile.AtEndOfStream
			If oFSO2.FileExists(groupList) Then
				Set oFile2 = oFSO2.OpenTextFile(groupList, 1)
					unit = oFile.ReadLine
				Do While Not oFile2.AtEndOfStream			
					group = oFile2.ReadLine
					output = unit & ", " & group
'msgbox(output)
		
					Dim objLog
					Set objLog = wshFSO.OpenTextFile(logfile, 8, True)
					objLog.Writeline output
					objLog.Close
					Err.Clear
				Loop
			End If
		Loop
	oFile.Close
	oFile2.Close	
End If

msgbox "Complete"
WScript.quit