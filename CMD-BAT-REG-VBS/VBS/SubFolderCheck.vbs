Option Explicit

Dim StrComputer, ObjNetwork, ObjWMI, ColSubFolders, ObjFolder
Dim ObjDelFSO, StrFilePath

Set ObjNetwork = CreateObject("WScript.Network")
StrComputer = Trim(ObjNetwork.ComputerName)
Set ObjNetwork = Nothing

Set ObjDelFSO = CreateObject("Scripting.FileSystemObject")
StrFilePath = Trim(ObjDelFSO.GetFile(WScript.ScriptFullName).ParentFolder)
Set ObjDelFSO = Nothing

Set ObjWMI = GetObject("WinMgmts:" & "{impersonationLevel=impersonate}!\\" & StrComputer & "\Root\CIMV2")
Set ColSubFolders = ObjWMI.ExecQuery("Associators of {Win32_Directory.Name='C:\Sky\Users'} " & "Where AssocClass = Win32_Subdirectory ResultRole = PartComponent")
DoThisDeleteJob:	WScript.Echo
For Each ObjFolder In ColSubFolders
	WScript.Echo ">> Checking Folder: " & Trim(ObjFolder.Name)
	ShowSubFolders(Trim(ObjFolder.Name))
Next
Set ColSubFolders = Nothing:	Set ObjWMI = Nothing
WScript.Echo
WScript.Echo "Task Completed"
WScript.Echo "Check Log File -- " & StrFilePath & "\LogFile.txt -- For Details"
WScript.Quit

Private Sub ShowSubFolders(StrFolder)

	Dim ObjFSO, ObjThisFolder, ColFiles, ObjFile
	Dim ChkDate, ChkAge
	
	' -- On Error Resume Next
	
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	Set ObjThisFolder = ObjFSO.GetFolder(StrFolder)
	Set ColFiles = ObjThisFolder.Files
	If ColFiles.Count = 0 Then
		Call WriteToLogFile("No File", StrFolder, "None", "None", "None")
	End If
	For Each ObjFile In ColFiles			
		ChkDate = ObjFile.DateLastModified
		ChkAge = DateDiff("d", ChkDate, Date)
		If ChkAge <= 90 Then
			Call WriteToLogFile(Trim(ObjFile.Name), StrFolder, FormatDateTime(ChkDate, 0), Trim(ObjFile.Name), ChkAge)
		Else
			Call WriteToLogFile("NotModified", StrFolder, FormatDateTime(ChkDate, 0), Trim(ObjFile.Name), "None")
		End If
	Next
	Set ColFiles = Nothing:	Set ObjThisFolder = Nothing:	Set ObjFSO = Nothing
End Sub

Private Sub WriteToLogFile(StrWhat, StrPath, ModDate, StrName, NumDays)
	
	Dim ObjWriteFSO, WriteHandle, ChkParent
	
	Set ObjWriteFSO = CreateObject("Scripting.FileSystemObject")
	If ObjWriteFSO.FileExists(StrFilePath & "\LogFile.txt") = False Then
		Set WriteHandle = ObjWriteFSO.OpenTextFile(StrFilePath & "\LogFile.txt", 8, True, 0)
		WriteHandle.Close:	Set WriteHandle = Nothing
	End If	
	If ObjWriteFSO.FileExists(StrFilePath & "\LogFile.txt") = True Then
		Set WriteHandle = ObjWriteFSO.OpenTextFile(StrFilePath & "\LogFile.txt", 8, True, 0)
		If StrComp(StrWhat, "No File", vbTextCompare) = 0 Then
			WriteHandle.WriteLine "## Checked The Folder -- " & StrPath
			WriteHandle.WriteLine vbTab & "There is NO FILE in the Folder -- " & StrPath
			WriteHandle.WriteLine vbNullString
		Else
			If StrComp(StrWhat, "NotModified", vbTextCompare) = 0 Then
				WriteHandle.WriteLine "## Checked The Folder: " & StrPath			
				WriteHandle.WriteLine vbTab & "Files In This Folder NOT MODIFIED in the Last 90 Days."			
				WriteHandle.WriteLine vbNullString
			Else
				WriteHandle.WriteLine "## Checked The Folder: " & StrPath
				Set ChkParent = ObjWriteFSO.GetFolder(StrPath)
				WriteHandle.WriteLine "## Parent Folder Name: " & UCase(ChkParent.ParentFolder)
				Set ChkParent = Nothing
				WriteHandle.WriteLine vbTab & "File Path and File Name: " & StrPath & "\" & StrName
				WriteHandle.WriteLine vbTab & "This File Has Been Modified in the Last 90 Days."
				WriteHandle.WriteLine vbTab & "Last Modified Date: " & ModDate
				WriteHandle.WriteLine vbTab & "This File Has Been Modified " & Numdays & " days ago."
				WriteHandle.WriteLine vbNullString
			End If			
		End If
		WriteHandle.Close:	Set WriteHandle = Nothing
	End If	
	Set ObjWriteFSO = Nothing
	
End Sub

Private Sub DoThisDeleteJob
	Set ObjDelFSO = CreateObject("Scripting.FileSystemObject")
	If ObjDelFSO.FileExists(StrFilePath & "\LogFile.txt") = True Then
		ObjDelFSO.DeleteFile StrFilePath & "\LogFile.txt", True
	End If
	Set ObjDelFSO = Nothing
End Sub