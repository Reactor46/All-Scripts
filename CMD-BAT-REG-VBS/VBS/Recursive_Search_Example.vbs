'Recursive search example
'------------------------

'The directory to search.
Const startDir = "c:\scripts"

'Variables
Dim objFSO, outFile, objFile, strFolder, strFile, subFolder

'Use the FileSystemObject to search directories and create the text file.
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create a text file.
Set outFile = objFSO.CreateTextFile("C:\DirectoryList.txt",True)

'Write a 'heading' to the text file.
outFile.WriteLine "Directory Listing of " & startDir
outFile.WriteLine String(25,"-")
outFile.WriteBlankLines 1

'Call the Search subroutine to start the recursive search.
Search objFSO.GetFolder(startDir)

Sub Search(sPath)
	
	'Assign the value of sPath to the strFolder variable.
	strFolder = sPath
	
	'Loop through each file in the sPath folder.
	For Each objFile In sPath.Files
	
		'Assign each file name to the strFile variable.
		strFile = strFile & objFile.Name & vbCrLf
	Next
	
	'Write the folder and all file names to the text file.
	outFile.Write strFolder & vbCrLf & strFile
	outFile.WriteBlankLines 1
	
	strFolder = ""
	strFile = ""
	
	'Find EACH SUBFOLDER.
	For Each subFolder In sPath.SubFolders
	
		'Call the Search subroutine to start the recursive search on EACH SUBFOLDER.
		Search objFSO.GetFolder(subFolder.Path)
	Next
End Sub

'Close the text file write stream.
outFile.Close

'Notify the user that the script is done.
WScript.Echo "Done."