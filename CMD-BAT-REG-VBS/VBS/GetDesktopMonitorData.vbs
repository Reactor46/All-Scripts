'#########################################################################
'  Script name:		GetDesktopMonitorData.vbs
'  Created on:		07/08/2011
'  Author:			Dennis Hemken
'  Purpose:			Returns the Desktop Monitor data
'					of all used monitors
'#########################################################################

Dim strComputer

    strComputer = "."
    fct_GetDesktopMonitorData(strComputer)

Public Function fct_GetDesktopMonitorData(strComputer)

Dim strOutput
Dim objWMIS
Dim colWMI
Dim objDesktopMonitor

On Error Resume Next

    Set objWMIS = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colWMI = objWMIS.ExecQuery("SELECT SystemName, Caption, Name, DeviceID, PnPDeviceID, ScreenHeight, ScreenWidth, PixelsPerXLogicalInch, PixelsPerYLogicalInch, Status FROM Win32_DesktopMonitor")
    
    For Each objDesktopMonitor In colWMI
        strOutput = strOutput & "Systemname: " & objDesktopMonitor.SystemName & vbCrLf _
					& "Name: " & objDesktopMonitor.Name & vbCrLf _
					& "Caption: " & objDesktopMonitor.Caption & vbCrLf _
					& "DeviceID: " & objDesktopMonitor.DeviceID & vbCrLf _
					& "PnPDeviceID: " & objDesktopMonitor.PnPDeviceID & vbCrLf _
					& "ScreenHeight: " & objDesktopMonitor.ScreenHeight & vbCrLf _
					& "ScreenWidth: " & objDesktopMonitor.ScreenWidth & vbCrLf _
					& "PixelsPerXLogicalInch: " & objDesktopMonitor.PixelsPerXLogicalInch & vbCrLf _
					& "PixelsPerYLogicalInch: " & objDesktopMonitor.PixelsPerYLogicalInch & vbCrLf _
					& "Status: " & objDesktopMonitor.Status & vbCrLf _
					& vbCrLf
    Next
    
    wscript.echo strOutput
    
End Function

'Retrieving WallpaperData