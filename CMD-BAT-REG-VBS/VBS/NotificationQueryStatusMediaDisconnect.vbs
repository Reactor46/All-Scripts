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
Debug "NotificationQueryStatusMediaDisconnect" & vbTab & "Monitoring network connection on LAN: " & sLanNic, 0
Debug "NotificationQueryStatusMediaDisconnect" & vbTab & "On media disconnect events, WLAN will be enabled: " & sWlanNic, 0

Set colMonitoredEvents = oWmiRootWmi.ExecNotificationQuery("Select * from MSNdis_StatusMediaDisconnect") 

Do While IsProcessRunning(iParentProcessID)
	Set sLatestEvent = colMonitoredEvents.NextEvent
	Debug "NotificationQueryStatusMediaDisconnect" & vbTab & "A network connection has disconnected: " & sLatestEvent.InstanceName, 0
	If sLatestEvent.InstanceName = sLanNic Then
		EnableAdapter(sWlanNic)
		Popup sLanNic & " has disconnected." & vbcrlf & sWlanNic & " is being enabled."
		If bIsPartOfDomain Then sAdInfo = RefreshAdSite()
		Set oRsAdapterStatus = GetAdapterStatus()
		sEventMessage = "MSNdis_StatusMediaDisconnect: A network connection has disconnected: " & sLatestEvent.InstanceName & vbcrlf &_
			"Wlan adapter has been enabled: " & sWlanNic & sAdInfo & vbcrlf & "--" & vbcrlf & "Network Adapter Status" & vbcrlf & "--" & vbcrlf & RsToString(oRsAdapterStatus)
		LogEvent iEventTypeWarning,sEventMessage
	End If
Loop
Debug "NotificationQueryStatusMediaDisconnect" & vbTab & "Parent Process ID " & iParentProcessID & " does not appear to be running. Enabling all adapters and quitting.", 0
EnableAllAdapters()