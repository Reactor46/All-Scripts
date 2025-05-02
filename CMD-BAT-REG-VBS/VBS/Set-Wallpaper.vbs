Option Explicit

On Error Resume Next

Dim WSHShell, RegKey, RegKey2, Wallpaper, Wallpaper2, objFSO, usrProf

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set WSHShell = CreateObject("WScript.Shell")

usrProf = WSHShell.expandEnvironmentStrings("%USERPROFILE%")

RegKey = "HKCU\Control Panel\Desktop\Wallpaper"
RegKey2 = "HKCU\Control Panel\Desktop\OriginalWallpaper"

Wallpaper = usrProf&"\Local Settings\Application Data\Microsoft\Wallpaper1.bmp"
Wallpaper2 = usrProf&"\Local Settings\Application Data\Microsoft\Wallpaper2.bmp"

objFso.copyFile "\\servername\logon$\Wallpaper1.bmp", wallpaper
objFso.copyFile "\\servername\logon$\Wallpaper1.bmp", wallpaper2


WSHShell.RegWrite regkey , Wallpaper

WSHShell.RegWrite regkey2 , Wallpaper   


' End code