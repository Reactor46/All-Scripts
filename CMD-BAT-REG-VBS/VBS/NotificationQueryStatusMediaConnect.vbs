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
Debug "NotificationQueryStatusMediaConnect" & vbTab & "Monitoring network connection on LAN: " & sLanNic, 0
Debug "NotificationQueryStatusMediaConnect" & vbTab & "On media connect events, WLAN will be disabled: " & sWlanNic, 0

Set colMonitoredEvents = oWmiRootWmi.ExecNotificationQuery("Select * from MSNdis_StatusMediaConnect") 

Do While IsProcessRunning(iParentProcessID)
	Set sLatestEvent = colMonitoredEvents.NextEvent
	Debug "NotificationQueryStatusMediaConnect" & vbTab & "A network connection has connected: " & sLatestEvent.InstanceName, 0
	If sLatestEvent.InstanceName = sLanNic Then
		DisableAdapter(sWlanNic)
		Popup sLanNic & " is connected." & vbcrlf & sWlanNic & " is being disabled."
		If bIsPartOfDomain Then sAdInfo = RefreshAdSite()
		Set oRsAdapterStatus = GetAdapterStatus()
		sEventMessage = "MSNdis_StatusMediaConnect: A network connection has connected: " & sLatestEvent.InstanceName & vbcrlf &_
			"Wlan adapter has been disabled: " & sWlanNic & sAdInfo & vbcrlf & "--" & vbcrlf & "Network Adapter Status" & vbcrlf & "--" & vbcrlf & RsToString(oRsAdapterStatus)
		LogEvent iEventTypeWarning,sEventMessage
	End If
Loop
Debug "NotificationQueryStatusMediaConnect" & vbTab & "Parent Process ID " & iParentProcessID & " does not appear to be running. Enabling all adapters and quitting.", 0
EnableAllAdapters()