'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FilesInFolders.vbs
'Version 1.1
'Writen By: Dave Gilmore
'Created On: 07/30/2013
'Last Modified On: 08/5/2013
' 
'Purpose: Given a base directory, this script will organize all files in that directory to year\month\day format 
'  		based on the date the file was last modified

'Parameters: 	/d		Directory to run in. If ommitted, will run in current directory
'				/s		Minimum size of file to archive. Files smaller than this will be deleted. If ommitted, all
'						files will be archived
'				/a		Directory to archive to. If ommited, will create archive in the directory specified by /d
'				/o		Last Date of files to archive.
'CHANGELOG
'
'v1.1
' - added /a switch in order to specify where you want the archive to go
' - added /0 switch in order to specify last modified date of files to archive. This is so you can keep some of the logs
'    in the "live" directory
'
Option explicit


Dim objFSO
Dim objBaseDir
Dim objFolder
Dim objNewYearFolder
Dim currentFolder
Dim objFile
Dim filePath
Dim strYear
Dim strMonth
Dim strDate
Dim intMinFileSize
Dim colNamedArguments
Dim objArchiveDir
Dim dtArchiveDate

Set objFSO = CreateObject("scripting.filesystemobject")
Set colNamedArguments = WScript.Arguments.Named

'If /d is not specified, use the current directory
If colNamedArguments.Exists("d") Then
 Set objBaseDir = objFSO.GetFolder(colNamedArguments.Item("d"))
Else
 Set objBaseDir = objFSO.GetAbsolutePathName(".")
End If

'If /s is not specified, get all files
If colNamedArguments.Exists("s") Then
 intMinFileSize = colNamedArguments.Item("s")
Else
 intMinFileSize = 0
End If

'If /a is not specified, archive to objBaseDir folder
If colNamedArguments.Exists("a") Then
 Set objArchiveDir = objFSO.GetFolder(colNamedArguments.Item("a"))
Else
 Set objArchiveDir = objBaseDir
End If

'If /o is not specified, archive all files
If colNamedArguments.Exists("o") Then
 dtArchiveDate = CDate(colNamedArguments.Item("o"))
Else
 dtArchiveDate = Date()
End If

	For each objFile in objBaseDir.Files
		
		If objFile.DateLastModified < dtArchiveDate then
			If objFile.Size > Cint(intMinFileSize) then
				strYear = Year(objFile.DateLastModified)
				strMonth = Month(objFile.DateLastModified)
				strDate = Day(objFile.DateLastModified)
		
				'Check if the archive folders exist. If not, create them
				If not objFSO.FolderExists(objArchiveDir & "\" & strYear) Then
					Set objNewYearFolder = objFSO.CreateFolder(objArchiveDir & "\" & strYear)
				End if
				If not objFSO.FolderExists(objArchiveDir & "\" & strYear & "\" & strMonth) Then
					Set objNewYearFolder = objFSO.CreateFolder(objArchiveDir & "\" & strYear & "\" & strMonth)
				End if
				If not objFSO.FolderExists(objArchiveDir & "\" & strYear & "\" & strMonth & "\" & strDate) Then
					Set objNewYearFolder = objFSO.CreateFolder(objArchiveDir & "\" & strYear & "\" & strMonth & "\" & strDate)
				End if
		
				filePath = objArchiveDir & "\" & strYear & "\" & strMonth & "\" & strDate & "\" & objFile.name 
				Wscript.Echo "Moving " & objFile.Name & " " & objFile.Size
				objFSO.MoveFile objFile.Path,filePath
			Else
				Wscript.Echo "Deleting " & objFile.Name & " " & objFile.Size
				objFile.Delete True
			End If
		End If
		 
			
	Next