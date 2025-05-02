'Declare all varibles in advance
Dim curPath,srcDir,dstDir,myDir,objFSO,objFolder,objFiles,objFile
Const OverwriteExisting = True

'Create a file sytem object as a base for manipulating files and directories
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Process command line arguments
Select Case WScript.Arguments.Count
'If only one argument is given, take it as source directory
'then create a destination directory in the current working directory
    Case 1
       		srcDir = WScript.Arguments.Item(0)
       		curPath = objFSO.GetAbsolutePathName(".")
     		dstDir = curPath & "\Split_Move_" & Day(Now()) & Month(Now()) & Year(Now()) & Hour(Now()) & Minute(Now()) & Second(Now())
	Case 2
'Two arguments given, first one is source directory
'Second one is destination directory
       		srcDir = WScript.Arguments.Item(0)
     		dstDir = WScript.Arguments.Item(1)
'In all other cases display Usage and exit script
    Case Else
       		Wscript.Echo "Usage: " & Wscript.ScriptName & "<Source Directory> [<Destination Directory>]"
       		WScript.Quit 1
End Select

'If source directory does not exist, display error message and exit script
If ( Not objFSO.FolderExists(srcDir) ) Then
   	Wscript.Echo "Invalid Source Directory"
   	WScript.Quit 1
End If

'If destination directory does not exist, create it.
If Not objFSO.FolderExists(dstDir) Then
   	objFSO.CreateFolder(dstDir)
End If

'Get list of files in source directory
Set objFolder = objFSO.getFolder( srcDir )
Set objFiles = objFolder.files

'For each file in source directory, destination should be:
'a directory under destination directory,
'that directory's name should constitute month and year of the file's creation time. 
For Each objFile in objFiles
'Construct destination directory name with month and year of file's creation time
myDir = dstDir & "\" & Year(objFile.DateCreated) & " - " & Month(objFile.DateCreated)
'If directory exists, move the file there.
If objFSO.FolderExists(myDir) Then
	objFSO.CopyFile objFile.Path , myDir & "\" & objFile.Name
	objFSO.DeleteFile objFile
'Otherwise create the directory and then move the file there.
Else
   	objFSO.CreateFolder(myDir)
	objFSO.CopyFile objFile.Path , myDir & "\" & objFile.Name
	objFSO.DeleteFile objFile
End If
Next