'#########################################################################
'  Script name:		GetWallpaperData.vbs
'  Created on:		08/05/2011
'  Author:			Dennis Hemken
'  Purpose:			Returns the Wallpaper data
'#########################################################################

Dim strComputer

    strComputer = "."
    fct_GetWallpaperData(strComputer)

Public Function fct_GetWallpaperData(strComputer)

Dim strOutput
Dim objWMIS
Dim colWMI
Dim objWallpaper

On Error Resume Next

    Set objWMIS = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colWMI = objWMIS.ExecQuery("SELECT Name, Wallpaper, WallpaperStretched, WallpaperTiled FROM Win32_Desktop")
    
    For Each objWallpaper In colWMI
        strOutput = strOutput & "Username: " & objWallpaper.Name & vbCrLf _
					& "Wallpaper: " & objWallpaper.Wallpaper & vbCrLf _
					& "WallpaperStretched: " & objWallpaper.WallpaperStretched & vbCrLf _
					& "WallpaperTiled: " & objWallpaper.WallpaperTiled & vbCrLf & vbCrLf
    Next
    
    wscript.echo strOutput
    
End Function

'Retrieving WallpaperData