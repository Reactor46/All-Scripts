Option Explicit
Dim oFS	: Set oFS = CreateObject("Scripting.FileSystemObject")
Dim sScriptDirectory : sScriptDirectory =  Left(Wscript.ScriptFullName, Len(Wscript.ScriptFullName)-Len(Wscript.ScriptName))

Dim aIncFiles, sIncFile, oFile, sText
aIncFiles = Array("Settings.inc","Global.inc")
For Each sIncFile In aIncFiles
	If oFS.FileExists(sScriptDirectory & sIncFile) Then
		Set oFile = oFS.OpenTextFile(sScriptDirectory & sIncFile,1)
		sText = oFile.ReadAll
		oFile.close
		ExecuteGlobal sText
	Else
		Wscript.Quit
	End If
Next

iDebugLogLevel = iSubProcessDebugLogLevel
Dim cNamedArguments
Set cNamedArguments = WScript.Arguments.Named
sLanNic = cNamedArguments.Item("lan")
sWlanNic = cNamedArguments.Item("wlan")
iParentProcessID = cNamedArguments.Item("ParentPID")

Debug "Script Parent ProcessID: " & iParentProcessID, 0
Debug "NotificationQueryDeviceChangeEvent" & vbTab & "Monitoring for device changes (configuration, arrivals, removals or docking).", 0
Debug "NotificationQueryDeviceChangeEvent" & vbTab & "On device change events, the existance of all monitored adapters (LAN: " & sLanNic & " & WLAN: " & sWlanNic & ") will be checked.", 0
Debug "NotificationQueryDeviceChangeEvent" & vbTab & "If either are missing, all adapters will be enabled.", 0

Set colMonitoredEvents = oWMIRootCimv2.ExecNotificationQuery("Select * from Win32_DeviceChangeEvent") 

Do While IsProcessRunning(iParentProcessID)
	Set sLatestEvent = colMonitoredEvents.NextEvent
	Select Case sLatestEvent.EventType
		Case 1		Debug "NotificationQueryDeviceChangeEvent" & vbTab & "A Win32_DeviceChangeEvent has occurred with EventType 1: Configuration Changed", 0
		Case 2		Debug "NotificationQueryDeviceChangeEvent" & vbTab & "A Win32_DeviceChangeEvent has occurred with EventType 2: Device Arrival", 0
		Case 3		Debug "NotificationQueryDeviceChangeEvent" & vbTab & "A Win32_DeviceChangeEvent has occurred with EventType 3: Device Removal", 0
		Case 4		Debug "NotificationQueryDeviceChangeEvent" & vbTab & "A Win32_DeviceChangeEvent has occurred with EventType 4: Docking", 0
		Case Else	Debug "NotificationQueryDeviceChangeEvent" & vbTab & "A Win32_DeviceChangeEvent has occurred with EventType 4: Unknown", 0
	End Select
	If DoesAdapterExist(sLanNic) And DoesAdapterExist(sWlanNic) Then
		Debug "NotificationQueryDeviceChangeEvent" & vbTab & "It doesn't appear the Win32_DeviceChangeEvent made a difference to the monitored interfaces.", 0
	Else
		Debug "NotificationQueryDeviceChangeEvent" & vbTab & "Win32_DeviceChangeEvent appears to be to do with one of the monitored interfaces! Enabling all adapters and quitting.", 0
		EnableAllAdapters()
		If bIsPartOfDomain Then sAdInfo = RefreshAdSite()
		Wscript.quit
	End If
Loop
Debug "NotificationQueryDeviceChangeEvent" & vbTab & "Parent Process ID " & iParentProcessID & " does not appear to be running. Enabling all adapters and quitting.", 0
EnableAllAdapters()