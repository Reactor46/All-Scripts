'Credit goes to Eric G. for writing this script.
'More info at http://frontslash.wordpress.com/2010/03/01/removing-internet-explorer-and-windows-media-player-from-taskbar/#comment-178

Option Explicit

Const CSIDL_COMMON_PROGRAMS = &H17
Const CSIDL_PROGRAMS = &H2
Const CSIDL_STARTMENU = &HB

Dim objShell, objFSO
Dim objCurrentUserStartFolder
Dim strCurrentUserStartFolderPath
Dim objAllUsersProgramsFolder
Dim strAllUsersProgramsPath
Dim objFolder
Dim objFolderItem
Dim colVerbs
Dim objVerb

Set objShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objCurrentUserStartFolder = objShell.NameSpace (CSIDL_STARTMENU)
strCurrentUserStartFolderPath = objCurrentUserStartFolder.Self.Path
Set objAllUsersProgramsFolder = objShell.NameSpace(CSIDL_COMMON_PROGRAMS)
strAllUsersProgramsPath = objAllUsersProgramsFolder.Self.Path

'Wait 120 seconds for Explorer to finish loading and the pins to get created (if it's the first time the user's logging in)
Wscript.sleep 120

'Internet Explorer
If objFSO.FileExists(strCurrentUserStartFolderPath & "\Programs\Internet Explorer.lnk") Then
    Set objFolder = objShell.Namespace(strCurrentUserStartFolderPath & "\Programs")
    Set objFolderItem = objFolder.ParseName("Internet Explorer.lnk")
    Set colVerbs = objFolderItem.Verbs
    For Each objVerb in colVerbs
        If Replace(objVerb.name, "&", "") = "Unpin from Taskbar" Then objVerb.DoIt
    Next
End If
'Windows Explorer
If objFSO.FileExists(strCurrentUserStartFolderPath & "\Programs\Accessories\Windows Explorer.lnk") Then
    Set objFolder = objShell.Namespace(strCurrentUserStartFolderPath & "\Programs\Accessories")
    Set objFolderItem = objFolder.ParseName("Windows Explorer.lnk")
    Set colVerbs = objFolderItem.Verbs
    For Each objVerb in colVerbs
        If Replace(objVerb.name, "&", "") = "Unpin from Taskbar" Then objVerb.DoIt
    Next
End If
'Windows Media Player
If objFSO.FileExists(strAllUsersProgramsPath & "\Windows Media Player.lnk") Then
    Set objFolder = objShell.Namespace(strAllUsersProgramsPath)
    Set objFolderItem = objFolder.ParseName("Windows Media Player.lnk")
    Set colVerbs = objFolderItem.Verbs
    For Each objVerb in colVerbs
        If Replace(objVerb.name, "&", "") = "Unpin from Taskbar" Then objVerb.DoIt
    Next
End If