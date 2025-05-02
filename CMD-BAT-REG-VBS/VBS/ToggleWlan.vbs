Option Explicit
Dim sScriptHost	: sScriptHost = LCase(Mid(WScript.FullName, InstrRev(WScript.FullName,"\")+1))
If Not sScriptHost="cscript.exe" Then
	Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
	WshShell.Run "cscript.exe " & Chr(34) & WScript.ScriptFullName & Chr(34)
	WScript.Quit
End If

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

If iScriptHostProcessID = 0 Or IsEmpty(iScriptHostProcessID) Then
	sEventMessage = "It doesn't seem the script could find its script host process ID. Quitting to avoid unintended consequences."
	LogEvent iEventTypeWarning,sEventMessage
	Debug "Main" & vbTab & sEventMessage, 0
	Wscript.quit
End If

Do While True
	iLanNICs = 0
	iWLanNICs = 0
	Do While iLanNICs <> 1 Or iWLanNICs <> 1
		Debug "Main" & vbTab & "Calling EnableAllAdapters." & vbCrlf & vbTab &_
			"This enables any disabled adapters so linking " & Chr(34) & "InstanceName From MSNdis_PhysicalMediumType" & Chr(34) & " with " & Chr(34) & "Name From Win32_NetworkAdapter" & Chr(34) & " is successful.", 0
		EnableAllAdapters()
		Debug "Main" & vbTab & "Sleeping for " & (lToggleAdapterSleepTime/1000) & " seconds to give the adapters a chance to get into a " & Chr(34) & "Connected" & Chr(34) & " state.", 0
		Wscript.Sleep lToggleAdapterSleepTime
		Debug "Main" & vbTab & "Calling GetAdapterStatus." & vbCrlf & vbTab &_
			"This uses WMI (MSNdis_PhysicalMediumType & Win32_NetworkAdapter) and returns the oRsAdapterStatus Recordset containing all network adapters.", 0
		Set oRsAdapterStatus = GetAdapterStatus()
		Debug "Main" & vbTab & "oRsAdapterStatus.RecordCount: " & oRsAdapterStatus.RecordCount, 0
		Debug "Main" & vbTab & "Filtering the oRsAdapterStatus Recordset returned by GetAdapterStatus for none WLAN and WLAN adapters respectively.", 0
		iTotalNICs = oRsAdapterStatus.RecordCount
		oRsAdapterStatus.Filter = "Wired=True"
		iLanNICs = oRsAdapterStatus.RecordCount
		Debug "Main" & vbTab & "Filtering for none WLAN adapters returned: " & iLanNICs, 0
		If iLanNICs = 1 Then
			sLanNic = oRsAdapterStatus("Name")
			sLanNicStatus = oRsAdapterStatus("NetConnectionStatusDisplayName")
			Debug "Main" & vbTab & "sLanNic: " & sLanNic, 0
			Debug "Main" & vbTab & "sLanNicStatus: " & sLanNicStatus, 0
		ElseIf iLanNICs > 1 Then
			Debug "Main" & vbTab & "More than one none WLAN adapters returned. Filtering for connected adapters...", 0
			iConnectedLanNICs = 0
			oRsAdapterStatus.MoveFirst
			Do Until oRsAdapterStatus.EOF
				If oRsAdapterStatus("NetConnectionStatusDisplayName") = "Connected" Then
					iConnectedLanNICs = iConnectedLanNICs + 1
					sLanNic = oRsAdapterStatus("Name")
					sLanNicStatus = oRsAdapterStatus("NetConnectionStatusDisplayName")
					Debug "Main" & vbTab & "sLanNic: " & sLanNic, 0
					Debug "Main" & vbTab & "sLanNicStatus: " & sLanNicStatus, 0
				End If
				oRsAdapterStatus.MoveNext
			Loop
			If iConnectedLanNICs = 1 Then iLanNICs = 1
		End If
		oRsAdapterStatus.Filter = "Wireless=True"
		iWLanNICs = oRsAdapterStatus.RecordCount
		Debug "Main" & vbTab & "Filtering for WLAN adapters returned: " & iWLanNICs, 0
		If iWLanNICs = 1 Then
			sWLanNic = oRsAdapterStatus("Name")
			sWLanNicStatus = oRsAdapterStatus("NetConnectionStatusDisplayName")
			Debug "Main" & vbTab & "sWLanNic: " & sWLanNic, 0
			Debug "Main" & vbTab & "sWLanNicStatus: " & sWLanNicStatus, 0
		End If
		If iLanNICs <> 1 Or iWLanNICs <> 1 Then
			Debug "Main" & vbTab & iLanNICs & " wired and " & iWLanNICs & " wireless adapter(s) found. The script needs one of each to run. Sleeping for 5 minutes to see if this situation changes.", 0
			Wscript.Sleep 300000
		End If
	Loop
	
	oRsAdapterStatus.Filter = adFilterNone
	sEventMessage = "Toggling the WLAN adapter state based on the ethernet adapter connection has started." & vbcrlf &_
		"Monitoring the following adapters: " & vbcrlf & vbTab &_
			"Lan Adapter: " & sLanNic & vbcrlf & vbTab &_
			"Wlan Adapter: " & sWLanNic & vbcrlf &_
		"Here's the current network adapter status" & vbcrlf & "--" & vbcrlf & RsToString(oRsAdapterStatus)
	LogEvent iEventTypeInformation,sEventMessage

	Debug "Main" & vbTab & "bOnlyRunWithDomainConnection: " & bOnlyRunWithDomainConnection, 0
	Debug "Main" & vbTab & "bIsPartOfDomain: " & bIsPartOfDomain, 0
	Debug "Main" & vbTab & "bIsConnectedToDomain: " & bIsConnectedToDomain, 0
	If ((bOnlyRunWithDomainConnection And bIsConnectedToDomain) Or (bOnlyRunWithDomainConnection = False Or bIsPartOfDomain = False)) Then
		If sLanNicStatus = "Connected" Then
			Debug "Main" & vbTab & sLanNic & " is " & sLanNicStatus & ". " & sWLanNic & " will be disabled.", 0
			DisableAdapter(sWLanNic)
	   	Popup sLanNic & " is connected." & vbcrlf & sWlanNic & " is being disabled."
			If bIsPartOfDomain Then sAdInfo = RefreshAdSite()
			Set oRsAdapterStatus = GetAdapterStatus()
			sEventMessage = "Wlan adapter has been disabled: " & sWLanNic & sAdInfo & vbcrlf & "--" & vbcrlf & "Network Adapter Status" & vbcrlf & "--" & vbcrlf & RsToString(oRsAdapterStatus)
			LogEvent iEventTypeSuccess,sEventMessage
		End If
		If IsProcessRunning(iMediaConnectPID) Then
		Else
			Debug "Main" & vbTab & "Starting NotificationQueryStatusMediaConnect.vbs." & vbCrlf & vbTab &_
				"This will wait for MediaConnect events from the LAN adapter (" & sLanNic & ") and" & vbCrlf & vbTab &_
				"disable the WLAN adapter (" & sWLanNic & ") when they occur.", 0
			iMediaConnectPID = CreateProcess("cscript " & Chr(34) & sScriptDirectory & "NotificationQueryStatusMediaConnect.vbs" & Chr(34) & " /lan:" & Chr(34) & sLanNic & Chr(34) & " /wlan:" & Chr(34) & sWLanNic & Chr(34) & " /ParentPID:" & Chr(34) & iScriptHostProcessID & Chr(34) & " > ExecNotificationQueryStatusMediaConnect.log")
			Debug "Main" & vbTab & "NotificationQueryStatusMediaConnect.vbs running under ProcessID: " & iMediaConnectPID, 0
		End If
		If IsProcessRunning(iMediaDisconnectPID) Then
		Else
			Debug "Main" & vbTab & "Starting NotificationQueryStatusMediaDisconnect.vbs." & vbCrlf & vbTab &_
				"This will wait for MediaDisconnect events from the LAN adapter (" & sLanNic & ") and" & vbCrlf & vbTab &_
				"enable the WLAN adapter (" & sWLanNic & ") when they occur.", 0
			iMediaDisconnectPID = CreateProcess("cscript " & Chr(34) & sScriptDirectory & "NotificationQueryStatusMediaDisconnect.vbs" & Chr(34) & " /lan:" & Chr(34) & sLanNic & Chr(34) & " /wlan:" & Chr(34) & sWLanNic & Chr(34) & " /ParentPID:" & Chr(34) & iScriptHostProcessID & Chr(34) & " > ExecNotificationQueryStatusMediaDisconnect.log")
			Debug "Main" & vbTab & "NotificationQueryStatusMediaDisconnect.vbs running under ProcessID: " & iMediaDisconnectPID, 0
		End If
		If IsProcessRunning(iDeviceChangePID) Then
		Else
			Debug "Main" & vbTab & "Starting NotificationQueryDeviceChangeEvent.vbs." & vbCrlf & vbTab &_
				"This will wait for DeviceChangeEvent's and enable all adapters if they involve the LAN adapter (" & sLanNic & ") or the WLAN adapter (" & sWLanNic & ").", 0
			iDeviceChangePID = CreateProcess("cscript " & Chr(34) & sScriptDirectory & "NotificationQueryDeviceChangeEvent.vbs" & Chr(34) & " /lan:" & Chr(34) & sLanNic & Chr(34) & " /wlan:" & Chr(34) & sWLanNic & Chr(34) & " /ParentPID:" & Chr(34) & iScriptHostProcessID & Chr(34) & " > ExecNotificationQueryDeviceChangeEvent.log")
			Debug "Main" & vbTab & "NotificationQueryDeviceChangeEvent.vbs running under ProcessID: " & iDeviceChangePID, 0
		End If
		
		If bOnlyRunWithDomainConnection And bIsPartOfDomain Then
			Do While (bIsConnectedToDomain And IsProcessRunning(iMediaConnectPID) And IsProcessRunning(iMediaDisconnectPID) And IsProcessRunning(iDeviceChangePID))
				Debug "Main" & vbTab & "Sleeping for " & (lDomainLoopTimer/1000) & " seconds (based on lDomainLoopTimer: " & lDomainLoopTimer & ")." & vbCrlf & vbTab &_
					"Domain connectivity will then be checked via IsConnectedToDomain and" & vbCrlf & vbTab &_
					"IsProcessRunning will then be called against the ProcessIDs for: " & vbCrlf & vbTab & vbTab &_
					"NotificationQueryStatusMediaDisconnect (" & iMediaDisconnectPID & ")" & vbCrlf & vbTab & vbTab &_
					"NotificationQueryStatusMediaConnect (" & iMediaConnectPID & ")" & vbCrlf & vbTab & vbTab &_
					"NotificationQueryDeviceChangeEvent (" & iDeviceChangePID & ")" & vbCrlf & vbTab &_
					"Loop will exit if that is unsuccessful.", 0
				Wscript.Sleep lDomainLoopTimer
				bIsConnectedToDomain = IsConnectedToDomain()
				Debug "Main" & vbTab & "bIsConnectedToDomain: " & bIsConnectedToDomain, 0
			Loop
		Else
			Do While (IsProcessRunning(iMediaConnectPID) And IsProcessRunning(iMediaDisconnectPID) And IsProcessRunning(iDeviceChangePID))
				Debug "Main" & vbTab & "Sleeping for " & (lNoDomainLoopTimer/1000) & " seconds (based on lNoDomainLoopTimer: " & lNoDomainLoopTimer & ")." & vbCrlf & vbTab &_
					"IsProcessRunning will then be called against the ProcessIDs for: " & vbCrlf & vbTab & vbTab &_
					"NotificationQueryStatusMediaDisconnect (" & iMediaDisconnectPID & ")" & vbCrlf & vbTab & vbTab &_
					"NotificationQueryStatusMediaConnect (" & iMediaConnectPID & ")" & vbCrlf & vbTab & vbTab &_
					"NotificationQueryDeviceChangeEvent (" & iDeviceChangePID & ")" & vbCrlf & vbTab &_
					"Loop will exit if that is unsuccessful.", 0
				Wscript.Sleep lNoDomainLoopTimer
			Loop
		End If
		TerminateProcess iMediaConnectPID
		TerminateProcess iMediaDisconnectPID
		TerminateProcess iDeviceChangePID
	ElseIf (bOnlyRunWithDomainConnection And bIsConnectedToDomain = False) Then
		Do Until bIsConnectedToDomain
			Debug "Main" & vbTab & "Sleeping for " & (lDomainLoopTimer/1000) & " seconds (based on lDomainLoopTimer: " & lDomainLoopTimer & ")." & vbCrlf & vbTab &_
				"Domain connectivity will then be checked via IsConnectedToDomain." & vbCrlf & vbTab &_
				"Loop will exit if that is successful.", 0
			Wscript.Sleep lDomainLoopTimer
			bIsConnectedToDomain = IsConnectedToDomain()
			Debug "Main" & vbTab & "bIsConnectedToDomain: " & bIsConnectedToDomain, 0
		Loop
	End If
Loop