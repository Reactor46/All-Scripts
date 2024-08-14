Dim objFSO   : Set objFSO = CreateObject( "Scripting.FileSystemObject" )
  Dim strRoot  : strRoot    = "C:\Data"
  Dim aFolders : aFolders   = Array( "Test1", "Test2", "Test3", "Test4" )
  Dim strUInp  : strUInp    = "userinput"
  Dim sFolder  : sFolder    = strRoot & "\" & strUInp
 
  If objFSO.FolderExists( sFolder ) Then
    WScript.Echo "Folder '" & sFolder & "' already exists"
  Else
    objFSO.CreateFolder sFolder ' easier than Subfolder.Add
  End If
 
  Dim oSubF    : Set oSubF  = objFSO.GetFolder( sFolder ).SubFolders
  Dim sSubF
  For Each sSubF In aFolders
     If objFSO.FolderExists( sFolder & "\" & sSubF ) Then
        WScript.Echo "Folder '" & sSubF & "' already exists"
     Else
        oSubF.Add sSubF
     End If
  Next