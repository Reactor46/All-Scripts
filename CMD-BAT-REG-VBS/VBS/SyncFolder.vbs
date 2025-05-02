Option Explicit

ForceScriptEngine("cscript")

Dim strFolder1, strFolder2, strLogDest, objFSO, objLogFile
dim strDate, thisDay, thisMonth, thisYear
thisDay = Day(date)
thisMonth = Month(date)
thisYear = Year(date)
strDate = thisMonth & "-" & thisDay & "-" & thisYear & "-" & Timer

strLogDest = "C:\Users\testuser\Desktop\"     ' Destination for the log file

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile(strLogDest & strDate & ".log", 8, True, 0)
objLogFile.WriteLine Now() & " :: " & "Backup Started"

strFolder1 = "C:\Users\testuser\Desktop\BackupScript\testFolder1\"
strFolder2 = "C:\Users\testuser\Desktop\BackupScript\testFolder2\"

Call SyncFolders(strFolder1, strFolder2)
Call SyncFolders(strFolder2, strFolder1)

Sub SyncFolders(strFolder1, strFolder2)
	Dim objFileSys
	Dim objFolder1
	Dim objFolder2
	Dim objFile1
	Dim objFile2
	Dim objSubFolder
	Dim arrFolders
	Dim i
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	arrFolders = Array(strFolder1, strFolder2)
	For i = 0 To 1 ' Make sure that missing folders are created first:
		If objFileSys.FolderExists(arrFolders(i)) = False Then
			wscript.echo("Creating folder " & arrFolders(i))
			objLogFile.WriteLine Now() & " :: " & "Creating folder " & arrFolders(i)
			objFileSys.CreateFolder(arrFolders(i))
		End If
	Next
	Set objFolder1 = objFileSys.GetFolder(strFolder1)
	Set objFolder2 = objFileSys.GetFolder(strFolder2)
	For i = 0 To 1
		If i = 1 Then ' Reverse direction of file compare in second run
			Set objFolder1 = objFileSys.GetFolder(strFolder2)
			Set objFolder2 = objFileSys.GetFolder(strFolder1)
		End If
		For Each objFile1 in objFolder1.files
			If Not objFileSys.FileExists(objFolder2 & "\" & objFile1.name) Then
				Wscript.Echo("Copying " & objFolder1 & "\" & objFile1.name & " :: to :: " & objFolder2 & "\" & objFile1.name)
				objLogFile.WriteLine Now() & " :: " & "Copying " & objFolder1 & "\" & objFile1.name & " :: to :: " & objFolder2 & "\" & objFile1.name
				objFileSys.CopyFile objFolder1 & "\" & objFile1.name, objFolder2 & "\" & objFile1.name
			Else
				Set objFile2 = objFileSys.GetFile(objFolder2 & "\" & objFile1.name)
				If objFile1.DateLastModified > objFile2.DateLastModified Then
					Wscript.Echo("Overwriting " & objFolder2 & "\" & objFile1.name & " :: with :: " & objFolder1 & "\" & objFile1.name)
					objLogFile.WriteLine Now() & " :: " & "Overwriting " & objFolder2 & "\" & objFile1.name & " :: with :: " & objFolder1 & "\" & objFile1.name
					objFileSys.CopyFile objFolder1 & "\" & objFile1.name, objFolder2 & "\" & objFile1.name    
				End If
			End If
		Next
	Next
	
	For Each objSubFolder in objFolder1.subFolders
		Call SyncFolders(strFolder1 & "\" & objSubFolder.name, strFolder2 & "\" & objSubFolder.name)
	Next
	Set objFileSys = Nothing
End Sub

Sub ForceScriptEngine(strScriptEng)
	' Forces this script to be run under the desired scripting host.
	' Valid arguments are "wscript" or "cscript".
	' The command line arguments are passed on to the new call.
	Dim arrArgs
	Dim strArgs
	For Each arrArgs In WScript.Arguments
		strArgs = strArgs & " " & Chr(34) & arrArgs & Chr(34)
	Next
	If Lcase(Right(Wscript.FullName, 12)) = "\wscript.exe" Then
		If Instr(1, Wscript.FullName, strScriptEng, 1) = 0 Then
			CreateObject("Wscript.Shell").Run "cscript.exe //Nologo " & Chr(34) & Wscript.ScriptFullName & Chr(34) & strArgs
			Wscript.Quit
		End If
	Else
		If Instr(1, Wscript.FullName, strScriptEng, 1) = 0 Then
			CreateObject("Wscript.Shell").Run "wscript.exe " & Chr(34) & Wscript.ScriptFullName & Chr(34) & strArgs
			Wscript.Quit
		End If
	End If
End Sub

objLogFile.WriteLine Now() & " :: " & "Backup Completed Successfully"
objLogFile.Close