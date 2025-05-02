'By K8l0
'Day numbers paramiter
strNumberDays = 90
'

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
If  ObjFSO.FolderExists("C:\Documents and Settings\") Then Set ObjFolder = ObjFSO.GetFolder("C:\Documents and Settings\")
If  ObjFSO.FolderExists("C:\Users\") Then Set ObjFolder = ObjFSO.GetFolder("C:\Users\")

On error resume next

For each ObjFolder in ObjFolder.SubFolders
	If not isexception(ObjFolder.name) and DateDiff("d", ObjFolder.DateLastModified,Now) > strNumberDays  then
		objFSO.DeleteFolder ObjFolder.path, True
	End if
Next

'Attention to Folders Exception
Function isException(byval foldername)
	select case foldername
		case "All Users"
			isException = True
		case "Default User"
			isException = True
		case "Default"
			isException = True
		case "LocalService"
			isException = True
		case "NetworkService"
			isException = True
		case "Administrator"
			isException = True
		case "Adm-Pass"
			isException = True
		case "AppData"
			isException = True
		case "Classic .NET AppPool"
			isException = True
		case "Public"
			isException = True
		case Else
			isException = False
	End Select
End Function