'------------------------
' Author : Assaf Miron
' Date 	 : 25/11/2008
' Description : Copies a file or a folder to a specific Location 
'				Using Drag and Drop
'------------------------

Const MyDestinationFolder = "C:\Temp\"
Const OverwriteExisting = True

Dim objFile,objFolder
Dim Arg

Set objFSO = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count > 0 Then
	For Each Arg in Wscript.Arguments
		Arg =  Trim(Arg)
    If InStr(Arg,".") Then
    ' Assume a File
      Set objFile = objFSO.GetFile(Arg)
      ' Copy file to the Dest Folder using the same name
      objFile.Copy MyDestinationFolder & objFile.Name,OverwriteExisting
    Else
    'Assume a Folder
      Set objFolder = objFSO.GetFolder(Arg)
      ' Copy Folder to the Dest Folder
      objFolder.Copy MyDestinationFolder, OverwriteExisting
    End If
	Next
End If