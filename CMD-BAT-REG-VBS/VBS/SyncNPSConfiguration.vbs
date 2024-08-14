'The sample scripts are not supported under any Microsoft standard support 
'program or service. The sample scripts are provided AS IS without warranty  
'of any kind. Microsoft further disclaims all implied warranties including,  
'without limitation, any implied warranties of merchantability or of fitness for 
'a particular purpose. The entire risk arising out of the use or performance of  
'the sample scripts and documentation remains with you. In no event shall 
'Microsoft, its authors, or anyone else involved in the creation, production, or 
'delivery of the scripts be liable for any damages whatsoever (including, 
'without limitation, damages for loss of business profits, business interruption, 
'loss of business information, or other pecuniary loss) arising out of the use 
'of or inability to use the sample scripts or documentation, even if Microsoft 
'has been advised of the possibility of such damages.

Option Explicit

Const WshRunning = 0
Dim objShell,objNetwork,objFSO

Set objShell = CreateObject("WScript.Shell")
Set objNetwork = CreateObject("WScript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Function GetUTCDateTime(ComputerName)
	'Get UTC Date Time from local or remote compurter.
	'In this script, this function is used to generate part of the WMI query.
	Dim strDateTime,objWMIService,colItems,objItem,intBias
	Set objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * From Win32_UTCTime")
	For Each objItem In colItems
		With objItem
			strDateTime = .Year
			If Len(.Month) = 1 Then
				strDateTime = strDateTime & "0" & .Month
			Else
				strDateTime = strDateTime & .Month
			End If
			If Len(.Day) = 1 Then
				strDateTime = strDateTime & "0" & .Day
			Else
				strDateTime = strDateTime & .Day
			End If
			If Len(.Hour) = 1 Then
				strDateTime = strDateTime & "0" & .Hour
			Else
				strDateTime = strDateTime & .Hour
			End If
			If Len(.Minute) = 1 Then
				strDateTime = strDateTime & "0" & .Minute
			Else
				strDateTime = strDateTime & .Minute
			End If
			If Len(.Second) = 1 Then
				strDateTime = strDateTime & "0" & .Second
			Else
				strDateTime = strDateTime & .Second
			End If
		End With
		strDateTime = strDateTime & ".000000+000"
		GetUTCDateTime = CStr(strDateTime)
	Next
End Function

Function GetOSCADObjectDN(SamAccountName,ObjectCategory)
	'This function is used to get distinguishedName of an AD object.
	Dim strDefaultNamingContext,objRootDSE
	Dim objConnection,objCommand,objRecordSet,strDN
	Set objRootDSE = GetObject("LDAP://RootDSE")
	strDefaultNamingContext = objRootDSE.Get("defaultNamingContext")	
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	If ObjectCategory = "computer" Then
		objCommand.CommandText = _
		"<LDAP://" & strDefaultNamingContext & _
		">;(&(objectCategory=" & ObjectCategory & ")(sAMAccountName=" & SamAccountName & _
		"$));distinguishedName;subtree"	
	Else
		objCommand.CommandText = _
		"<LDAP://" & strDefaultNamingContext & _
		">;(&(objectCategory=" & ObjectCategory & ")(sAMAccountName=" & SamAccountName & _
		"));distinguishedName;subtree"
	End If
	Set objRecordSet = objCommand.Execute
	If objRecordSet.RecordCount = 1 Then
		While Not objRecordSet.EOF
			GetOSCADObjectDN = objRecordSet.Fields("distinguishedName")
			objRecordSet.MoveNext
		Wend
	Else
		GetOSCADObjectDN = Null
	End If
	objConnection.Close
End Function

Function GetOSCNPSConfiguration(ComputerName,NPSServerType,SharedFolder,Timeout)
	'This function is used to get the NPS configuration file from Hub server and proxy server. 
	Dim strLocalComputerName,blnIsLocalComputer,strFullConfigurationFileName,strNewConfigurationFileName
	Dim objExec,strCommand,objWMIWin32Process,intReturnValue,intPID,strRemoteNPSConfigFile
	Dim intCounter
	strLocalComputerName = objNetwork.ComputerName
	If ComputerName = strLocalComputerName Then
		blnIsLocalComputer = True
	Else
		blnIsLocalComputer = False
	End If
	If Not objFSO.FolderExists(SharedFolder) Then
		WScript.Echo SharedFolder & "does not exist."
		WScript.Quit(1)
	End If
	'This function will generate the configuration file name based on the NPS server type.
	Select Case UCase(NPSServerType)
		Case "HUB"
			If Left(SharedFolder,1) <> "\" Then
				strFullConfigurationFileName = SharedFolder & "\" & "NPS-Hub-Server-Configuration.xml"
			Else
				strFullConfigurationFileName = SharedFolder & "NPS-Hub-Server-Configuration.xml"
			End If
		Case "PROXY"
			If Left(SharedFolder,1) <> "\" Then
				strFullConfigurationFileName = SharedFolder & "\" & "NPS-Proxy-Server-Configuration.xml"
			Else
				strFullConfigurationFileName = SharedFolder & "NPS-Proxy-Server-Configuration.xml"
			End If
		Case Else
			WScript.Echo "Please enter a valid NPS server type."
			WScript.Quit(1)
	End Select
	'Execute the command, use WScript.Shell for running the command in a local computer.
	'Use WMI for running the command in a remote computer.
	strCommand = "netsh nps export filename=" & strFullConfigurationFileName & " exportPSK=YES"
	If blnIsLocalComputer Then
		Set objExec = objShell.Exec("netsh nps export filename=" & strFullConfigurationFileName & " exportPSK=YES")
		WScript.Echo "[" & ComputerName & "] Exporting " & NPSServerType & " NPS configuration."
		Do While (objExec.Status <> WshRunning)
			Call WScript.Sleep(Timeout)
		Loop
		If objExec.ExitCode <> 0 Then
			WScript.Echo "[" & ComputerName & "] Cannot export the NPS configuration file."
			WScript.Quit(1)
		Else
			WScript.Echo "[" & ComputerName & "] " & NPSServerType & " NPS configuration is exported."
		End if
	Else
		'Cannot direct export configuration file to the shared folder.
		'You will get an Access Denied Exception.
		'So you need to use a local folder instead, after the configuration file is created, copy it to the shared folder.
		strCommand = Replace(strCommand,SharedFolder,"C:\Windows\Temp\")
		strNewConfigurationFileName = Replace(strFullConfigurationFileName,SharedFolder,"")
		Set objWMIWin32Process = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2:Win32_Process")
		intReturnValue = objWMIWin32Process.Create(strCommand, Null, Null, intPID)
		If intReturnValue <> 0 Then
			WScript.Echo "[" & ComputerName & "] Cannot start netsh process. The return value is " & intReturnValue & "."
			WScript.Quit(1)
		Else
			WScript.Echo "[" & ComputerName & "] Exporting " & NPSServerType & " NPS configuration."
		End If
		strRemoteNPSConfigFile = "\\" & ComputerName & "\c$\Windows\Temp\" & strNewConfigurationFileName
		Do Until objFSO.FileExists(strRemoteNPSConfigFile)
			Call WScript.Sleep(Timeout)
			intCounter = intCounter + 1
			Wscript.Echo "[" & ComputerName & "] Waiting netsh nps export operation to complete.(" & intCounter & "/5)"			
			If intCounter > 4 Then 
				Exit Do
			End If
		Loop
		If objFSO.FileExists(strRemoteNPSConfigFile) Then
			Call objFSO.CopyFile(strRemoteNPSConfigFile,SharedFolder,True)
			If Err.Number = 0 Then
				WScript.Echo "[" & ComputerName & "] " & NPSServerType & " NPS configuration has been exported."
			End If
		Else
			WScript.Echo "[" & ComputerName & "] Cannot export NPS configuration file."
			WScript.Quit(1)
		End If
	End If
	GetOSCNPSConfiguration = strFullConfigurationFileName
End Function

Function GetOSCNPSProxyServer(NPSHubServerConfigurationFile)
	'This function is used to get the RADIUS proxy server list
	'from the configuration file of NPS hub server.
	Dim blnIsLoaded,objXML,objRootNode,objNPSProxyServersNode,objNPSProxyServerNode
	Dim strXPath,i
	Dim arrProxyServers()
	strXPath = "/Root/Children/Microsoft_Internet_Authentication_Service/Children/Protocols/" & _
	"Children/Microsoft_Radius_Protocol/Children/Clients/Children/*"
	Set objXML = CreateObject("Msxml2.DOMDocument.3.0")
	objXML.async = False
	If Not objFSO.FileExists(NPSHubServerConfigurationFile) Then
		WScript.Echo "Cannot find the configuration file which is exported from the NPS hub server."
	Else
		blnIsLoaded = objXML.load(NPSHubServerConfigurationFile)
	End If
	Set objRootNode = objXML.documentElement
	Set objNPSProxyServersNode = objRootNode.selectNodes(strXPath)
	For Each objNPSProxyServerNode In objNPSProxyServersNode
		ReDim Preserve arrProxyServers(i)
		arrProxyServers(i) = objNPSProxyServerNode.nodeName
		i = i + 1
	Next
	If Not IsNull(arrProxyServers) Then
		GetOSCNPSProxyServer = arrProxyServers
	Else
		WScript.Echo "Cannot get proxy servers from configuration file."
	End If
	Set objXML = Nothing
End Function

Function SyncOSCNPSConfiguration(NPSProxyServer,BackupFolder,ProxyServerConfiguration,ProxyServerExclusionList,Timeout)
	'This function is used to synchronize the configuration among different RADIUS proxy servers.
	On Error Resume Next
	Dim i,j,objRegExp,blnCheckExclusionList,strNPSProxyServerName,objWMIService,objWMIWin32Process,strCommand
	Dim arrNPSTargetProxyServer(),intPID,intReturnValue,strRemoteNPSConfigFile,dtmCmdStartDT,colItems
	Dim intCounter,objDicSucceed,objDicFail,objDicItems
	Set objDicSucceed = CreateObject("Scripting.Dictionary")
	Set objDicFail = CreateObject("Scripting.Dictionary")
	
	'Processing the exclusion list.
	If ProxyServerExclusionList<> "" Then
		Set objRegExp = New RegExp
		objRegExp.Global = True
		objRegExp.IgnoreCase = True
		If InStr(ProxyServerExclusionList,",") > 0 Then
			objRegExp.Pattern = Replace(ProxyServerExclusionList,",","|")
		Else
			objRegExp.Pattern = ProxyServerExclusionList
		End If
		For i = 0 To UBound(NPSProxyServer)
			If Not objRegExp.Test(NPSProxyServer(i)) Then
				ReDim Preserve arrNPSTargetProxyServer(j)
				arrNPSTargetProxyServer(j) = NPSProxyServer(i)
				j = j + 1
			End If
		Next	
	Else
		For i = 0 To UBound(NPSProxyServer)
			ReDim Preserve arrNPSTargetProxyServer(i)
			arrNPSTargetProxyServer(i) = NPSProxyServer(i)
		Next
	End If
	
	'Begin the synchronization process
	For i = 0 To UBound(arrNPSTargetProxyServer)
		'Backup configuration
		intCounter = 0
		strNPSProxyServerName = arrNPSTargetProxyServer(i)
		strCommand = "netsh nps export filename=c:\Windows\Temp\" & strNPSProxyServerName & ".xml exportPSK=YES"
		Set objWMIWin32Process = GetObject("winmgmts:\\" & strNPSProxyServerName & "\root\cimv2:Win32_Process")
		If Err.Number <> 0 Then
			Call objDicFail.Add(strNPSProxyServerName,strNPSProxyServerName)
			WScript.Echo "[" & strNPSProxyServerName & "] Cannot run netsh due to following reason: " & vbCrLf & Err.Description
			Err.Clear
		Else
			intReturnValue = objWMIWin32Process.Create(strCommand, Null, Null, intPID)
			If intReturnValue <> 0 Then
				WScript.Echo "[" & strNPSProxyServerName & "] Cannot start netsh process. The return value is " & intReturnValue & "."
			Else
				WScript.Echo "[" & strNPSProxyServerName & "] Backing up NPS configuration to " & BackupFolder
			End If
			strRemoteNPSConfigFile = "\\" & strNPSProxyServerName & "\c$\Windows\Temp\" & strNPSProxyServerName & ".xml"
			Do Until objFSO.FileExists(strRemoteNPSConfigFile)
				Call WScript.Sleep(Timeout)
				intCounter = intCounter + 1
				Wscript.Echo "[" & strNPSProxyServerName & "] Waiting netsh nps export operation to complete.(" & intCounter & "/5)"			
				If intCounter > 4 Then
					Exit Do
				End If
			Loop
			If objFSO.FileExists(strRemoteNPSConfigFile) Then
				Call objFSO.CopyFile(strRemoteNPSConfigFile,BackupFolder,True)
			Else
				Call objDicFail.Add(strNPSProxyServerName,strNPSProxyServerName)
				WScript.Echo "[" & strNPSProxyServerName & "] Cannot backup NPS configuration file."
				WScript.Quit(1)
			End If
			'Copy new configuration file
			If objFSO.FileExists(ProxyServerConfiguration) Then
				Call objFSO.CopyFile(ProxyServerConfiguration,"\\" & strNPSProxyServerName & "\c$\Windows\Temp\NPS_New_Configuration.xml",True)
			Else
				Call objDicFail.Add(strNPSProxyServerName,strNPSProxyServerName)
				WScript.Echo "[" & strNPSProxyServerName & "] Cannot find the configuration file which will be imported."
				WScript.Quit(1)
			End If
			'Reset NPS configuration on a remote NPS server.
			intCounter = 0
			dtmCmdStartDT = GetUTCDateTime(strNPSProxyServerName)
			intReturnValue = objWMIWin32Process.Create _
			("cmd /c netsh nps reset config && eventcreate /T Information /ID 1 /L Application /D ""Succeeded to reset NPS configuration""", Null, Null, intPID)
			Set objWMIService = GetObject("winmgmts:\\" & strNPSProxyServerName & "\root\cimv2")
			Set colItems = objWMIService.ExecQuery("Select * From Win32_NTLogEvent Where " & _
			"LogFile='Application' And EventCode='1' And SourceName='EventCreate' And TimeGenerated > '" & dtmCmdStartDT & "'")
			Do Until colItems.Count = 1
				Set colItems = objWMIService.ExecQuery("Select * From Win32_NTLogEvent Where " & _
				"LogFile='Application' And EventCode='1' And SourceName='EventCreate' And TimeGenerated > '" & dtmCmdStartDT & "'")
				Call WScript.Sleep(Timeout)
				intCounter = intCounter + 1
				Wscript.Echo "[" & strNPSProxyServerName & "] Waiting netsh nps reset configuration operation to complete.(" & intCounter & "/5)"			
				If intCounter > 4 Then 
					Exit Do
				End If
			Loop
			If colItems.Count = 0 Then
				Call objDicFail.Add(strNPSProxyServerName,strNPSProxyServerName)			
				WScript.Echo "[" & strNPSProxyServerName & "] Cannot reset NPS configuration."
			Else
				'Import NPS configuration on a remote NPS server.
				intCounter = 0
				dtmCmdStartDT = GetUTCDateTime(strNPSProxyServerName)
				intReturnValue = objWMIWin32Process.Create _
				("cmd /c netsh nps import c:\Windows\Temp\NPS_New_Configuration.xml " & _
				"&& eventcreate /T Information /ID 2 /L Application /D ""Succeeded to import NPS configuration.""", Null, Null, intPID)
				Set objWMIService = GetObject("winmgmts:\\" & strNPSProxyServerName & "\root\cimv2")
				Set colItems = objWMIService.ExecQuery("Select * From Win32_NTLogEvent Where " & _
				"LogFile='Application' And EventCode='2' And SourceName='EventCreate' And TimeGenerated > '" & dtmCmdStartDT & "'")
				Do Until colItems.Count = 1
					Set colItems = objWMIService.ExecQuery("Select * From Win32_NTLogEvent Where " & _
					"LogFile='Application' And EventCode='2' And SourceName='EventCreate' And TimeGenerated > '" & dtmCmdStartDT & "'")
					Call WScript.Sleep(Timeout)
					intCounter = intCounter + 1
					Wscript.Echo "[" & strNPSProxyServerName & "] Waiting netsh nps import configuration operation to complete.(" & intCounter & "/5)"			
					If intCounter > 4 Then 
						Exit Do
					End If
				Loop
				If colItems.Count = 0 Then
					Call objDicFail.Add(strNPSProxyServerName,strNPSProxyServerName)				
					WScript.Echo "[" & strNPSProxyServerName & "] Cannot import NPS configuration."
				Else
					Call objDicSucceed.Add(strNPSProxyServerName,strNPSProxyServerName)				
					WScript.Echo "[" & strNPSProxyServerName & "] New RADIUS proxy server configuration has been applied."
				End If
			End If
		End If
	Next
	'Display Report
	WScript.Echo ""
	WScript.Echo "There are " & UBound(arrNPSTargetProxyServer) + 1 & " RADIUS proxy server(s) need to apply new configuration. " & _
	objDicSucceed.Count & " server(s) applied the new configuration. " & objDicFail.Count & " server(s) did not apply the new configuration."
	If objDicFail.Count > 0 Then
		WScript.Echo "Here is the server list which you need to import the configuration manually:"
		objDicItems = objDicFail.Items
		intCounter = 0
		For intCounter = 0 To objDicFail.Count
			WScript.Echo objDicItems(intCounter)
		Next
	End If
End Function

Sub OSCScriptUsage
	WScript.Echo "How to use this script:" & vbCrLf & vbCrLf _
	& "1. Logon to one NPS Hub Server." & vbCrLf _
	& "2. Run the script with following command:" & vbCrLf  _
	& "cscript //nologo SyncNPSConfiguration.vbs /NPSHubServerName:""NPSHubServerName"" /NPSProxyServerName:""NPSProxyServerName"" " _
	& "/NPSProxyServerExclusionList:""NPSProxyServerName01,NPSProxyServerName02"" /SharedFolder:""\\server\share\"" /Timeout:5000"
End Sub

Sub Main
	Dim strNPSHubServerName,strNPSProxyServerName,strSharedFolder,intTimeout
	Dim arrProxyServers,strNPSProxyServerConfigurationFilePath,strNPSHubServerConfigurationFilePath
	Dim strNPSProxyServerExclusionList,objRegExp,objArgs,i
	Set objRegExp = New RegExp
	Set objArgs = WScript.Arguments
	objRegExp.Global = True
	objRegExp.IgnoreCase = True
	objRegExp.Pattern = "npshubservername|npsproxyservername|npsproxyserverexclusionlist|sharedfolder|timeout"
	'Please run this script with cscript.
	If InStr(WScript.FullName,"cscript") = 0 Then
		Call OSCScriptUsage()
		WScript.Quit(1)
	End If	
	'Verify Arguments
	If WScript.Arguments.Named.Count <> 5 Then
		Call OSCScriptUsage()
		WScript.Quit(1)
	Else
		For i = 0 To objArgs.Count - 1
			If Not objRegExp.Test(objArgs(i)) Then
				Call OSCScriptUsage()
				WScript.Quit(1)
			Else
				With objArgs.Named
					If .Exists("npshubservername") Then strNPSHubServerName = .Item("npshubservername")
					If .Exists("npsproxyservername") Then strNPSProxyServerName = .Item("npsproxyservername")
					If .Exists("npsproxyserverexclusionlist") Then strNPSProxyServerExclusionList = .Item("npsproxyserverexclusionlist")
					If .Exists("sharedfolder") Then strSharedFolder = .Item("sharedfolder")
					If .Exists("timeout") Then intTimeout = .Item("timeout")
				End With
			End If
		Next
	End If
	If Not objFSO.FolderExists(strSharedFolder) Then
		WScript.Echo "Please enter a valid path name."
		WScript.Quit(1)
	Elseif Right(strSharedFolder,1) <> "\" Then
		strSharedFolder = strSharedFolder + "\"
	End If
	If IsNull(GetOSCADObjectDN(strNPSHubServerName,"computer")) Then
		WScript.Echo "Please enter a valid NPS hub server name which belongs to current domain."
		WScript.Quit(1)
	End If
	If IsNull(GetOSCADObjectDN(strNPSProxyServerName,"computer")) Then
		WScript.Echo "Please enter a valid RADIUS proxy server name which belongs to current domain."
		WScript.Quit(1)
	End If
	If intTimeout = "" Then intTimeout = 5000
	'Call functions
	strNPSHubServerConfigurationFilePath = GetOSCNPSConfiguration(strNPSHubServerName,"Hub",strSharedFolder,intTimeout)
	strNPSProxyServerConfigurationFilePath = GetOSCNPSConfiguration(strNPSProxyServerName,"Proxy",strSharedFolder,intTimeout)
	arrProxyServers = GetOSCNPSProxyServer(strNPSHubServerConfigurationFilePath)
	Call SyncOSCNPSConfiguration(arrProxyServers,strSharedFolder,strNPSProxyServerConfigurationFilePath,strNPSProxyServerExclusionList,intTimeout)	
End Sub

Call Main()