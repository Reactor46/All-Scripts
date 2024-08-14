Set oFileSys = WScript.CreateObject("Scripting.FileSystemObject")
sRoot = "C:\Downloads\"			
today = Date
nMaxFileAge = 7					

DeleteFiles(sRoot)

Function DeleteFiles(ByVal sFolder)

	Set oFolder = oFileSys.GetFolder(sFolder)
	Set aFiles = oFolder.Files
	Set aSubFolders = oFolder.SubFolders

	For Each file in aFiles
		dFileCreated = FormatDateTime(file.DateCreated, "2")
		If DateDiff("d", dFileCreated, today) > nMaxFileAge Then
			file.Delete(True)
		End If
	Next

	For Each folder in aSubFolders
		DeleteFiles(folder.Path)
	Next

End Function