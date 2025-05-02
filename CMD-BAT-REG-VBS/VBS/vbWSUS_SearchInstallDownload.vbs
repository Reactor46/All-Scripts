' ---------------------------------------------------------------------
' ABOUT
'
' ABOUT::Description
' Script to detect, download and install updates
' based on : http://msdn.microsoft.com/en-us/library/aa387102%28VS.85%29.aspx

' ABOUT::ExpectedArgs (may received arguments from vbWSUS_exec.vbs)
' strWUAUConfig;strWUServerConfig;strWUProxyConfig;strWUSearch;strWUFilter
'
' ABOUT::TellMeMore
' see README.txt and VERSION.txt
'
' /ABOUT
' ---------------------------------------------------------------------

bDEBUG = False
bPROMPT = False ' True = confirm update installation, don't set bPROMPT to True if you are running from the task scheduler without Desktop Interaction of the script will get stuck !

On Error Resume Next
' ---------------------------------------------------------------------
' CONFIGURATION
'
' CONFIGURATION::Constants
' Registry
Const HKEY_CURRENT_USER 	= &H80000001
Const HKEY_LOCAL_MACHINE 	= &H80000002
Const REG_SZ 				= 1
Const REG_EXPAND_SZ 		= 2
Const REG_BINARY 			= 3
Const REG_DWORD 			= 4
Const REG_MULTI_SZ 			= 7
' Files
Const ForAppending 			= 8
Const ForWriting 			= 2
Const ForReading 			= 1
Const CreateIfNotExists 	= True
Const Overwrite 			= True
' vbWSUS
Const sep 					= ";"
Const LogLevel 				= 3
Const ERROR_DEFAULT 		= 1
Const vbRegeditSet			= 0
Const vbRegeditRestore		= 1

' CONFIGURATION::Objects
Set objArgs 		= WScript.Arguments
Set objShell 		= WScript.CreateObject("WScript.Shell") 
Set objFSO 			= Wscript.CreateObject("Scripting.FileSystemObject")
Set objReg			= GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

' CONFIGURATION::Objects::Registry

' Reset computer name
strComputer = objShell.ExpandEnvironmentStrings("%Computername%")

' CONFIGURATION::Files
' Open logfile ' DO NOT EDIT OR vbWSUS_exec.vbs WILL NOT BE ABLE TO FETCH LOG
strLogFile = objFSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\vbWSUS_SearchInstallDownload.log"
Set objLog = objFSO.OpenTextFile (strLogfile, ForWriting, CreateIfNotExists)

' CONFIGURATION::Parameters
' This values have to be edited in you plan on using vbWSUS_SearchInstallDownload.vbs without vbWSUS_scheduler.vbs
' Else ignore them !
log "INFO", "no arguments found, setting default values"
strWUSearch 		= "IsAssigned=1 and IsHidden=0 and IsInstalled=0 and Type='Software'"
strWUFilter 		= "Critical,Important"
strWUAUConfig 		= "_noWUAUConfig_"
strWUServerConfig 	= "_noWUServerConfig_"
strWUProxyConfig 	= "_noWUProxyConfig_"

' CONFIGURATION::Parameters::Registry
' Uncomment these lines to update the registry settings
' Windows Update Automatic Updates configuration, don't forget to set
'strWUAUConfig 		= HKEY_LOCAL_MACHINE & "|SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update|" & _
'												"dw=AUOptions=1," & _
'												"dw=ElevateNonAdmins=1," & _
'												"dw=IncludeRecommendedUpdates=0," & _
'												"dw=UseWUServer=0"

' Windows Update WSUS Server configuration, don't forget to enable strWUAUConfig and to set UseWUServer=1 above !
'strWUServerConfig 	= HKEY_LOCAL_MACHINE & "|SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate|" & _
'												"dw=ElevateNonAdmins=1," & _
'												"str=WUServer=http://new.wsusserver.lan," & _
'												"str=WUStatusServer=http://my.wsusserver.lan"

' Windows WinHTTP Proxy
'strWUProxyConfig	= HKEY_LOCAL_MACHINE & "|SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections|" & _
'												"bin=WinHttpSettings=http://my.proxyserver.lan"

'
' /CONFIGURATION
' ---------------------------------------------------------------------

' ---------------------------------------------------------------------
' MAIN
'

'
' MAIN::PARSE_ARGS
If objArgs.Count = 5 Then
	strWUAUConfig 		= objArgs(0)
	strWUServerConfig 	= objArgs(1)
	strWUProxyConfig 	= objArgs(2)
	strWUSearch 		= objArgs(3)
	strWUFilter 		= objArgs(4)

	If strWUFilter = "" Then
		strWUFilter= "All"
	End If
Else

End If
'
' /MAIN::PARSE_ARGS
'
restartService("wuauserv")
restartService("BITS")

'
' MAIN::CONFIGURE_WU
'
boolRestartWUAUSERV = False
boolRestartBITS 	= False

If strWUAUConfig <> "_noWUAUConfig_" Then
	If bDEBUG Then log "DEBUG", "strWUAUConfig=" & strWUAUConfig
	
	strWUAUConfig = vbRegedit(strWUAUConfig, vbRegeditSet)
	boolRestartWUAUSERV = True
End If

If strWUServerConfig <> "_noWUServerConfig_" Then
	If bDEBUG Then log "DEBUG", "strWUServerConfig=" & strWUServerConfig
	
	strWUServerConfig = vbRegedit(strWUServerConfig, vbRegeditSet)	
	boolRestartWUAUSERV = True
End If

If strWUProxyConfig <> "_noWUProxyConfig_" Then
	' Get the proxy URL from strWUProxyConfig
	arrWUProxyConfig = Split(strWUProxyConfig, "=")
	strWUProxyURL = arrWUProxyConfig(2)
	' Convert the proxy URL chars to their decimal value
	decWUProxyURL = proxy2dec(strWUProxyURL)
	' Replace the string URL in strWUProxyConfig by the decimal values
	If bDEBUG Then log "DEBUG", "strWUProxyConfig=" & strWUProxyConfig
	strWUProxyConfig = Replace(strWUProxyConfig, strWUProxyURL, decWUProxyURL)
	If bDEBUG Then log "DEBUG", "strWUProxyConfig=" & strWUProxyConfig
	
	log "INFO", "Setting proxy to " & strWUProxyURL
	strWUProxyConfig = vbRegedit(strWUProxyConfig, vbRegeditSet)
	boolRestartBITS = True
	boolRestartWUAUSERV = True
End If

If boolRestartWUAUSERV Then
	restartService("wuauserv")
End If

If boolRestartBITS Then
	restartService("BITS")
End If

'
' /MAIN::CONFIGURE_WU
'

'
' MAIN::SEARCH_FOR_UPDATES
'
log "INFO", "Searching for updates matching " & strWUSearch
Set updateSession 	= CreateObject("Microsoft.Update.Session")
checkError Err, 1, "Failed to create session object (Set updateSession = CreateObject(Microsoft.Update.Session))"
Set updateSearcher 	= updateSession.CreateupdateSearcher()
checkError Err, 1, "Failed to create updateSearcher (Set updateSearcher = updateSession.CreateupdateSearcher())"
Set searchResult = updateSearcher.Search(strWUSearch)
checkError Err, 1, "Failed to search updates matching " & strWUSearch & "(Set searchResult = updateSearcher.Search(" & strWUSearch & "))"


log "INFO", "List of applicable updates:"
For I = 0 To searchResult.Updates.Count-1
	Set update = searchResult.Updates.Item(I)
	strSeverity = update.MsrcSeverity
	If strSeverity = "" Then
		strSeverity = "NoSeverity"
	End If
	log "INFO", "Found " & getKBNumber(update.Title) & " :: " & strSeverity & " :: " & update.Title
Next

If searchResult.Updates.Count = 0 Then
	strMessage = "No applicable updates found matching search scope " & chr(34) & strWSUSearch & chr(34)
	log "INFO", strMessage & ", exiting..."
	objLog.WriteLine "RESULT_KB:" & strMessage
	runPostUpdateActions 2, 0
End If
'
' /MAIN::SEARCH_FOR_UPDATES
'

'
' MAIN::BUILD_DOWNLOAD_LIST
' Add updates to download list given 'strWUFilter'
log "INFO",  "Creating collection of updates to download matching filter '" & strWUFilter & "'"
Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")

For I = 0 to searchResult.Updates.Count-1
	Set update = searchResult.Updates.Item(I)
	strSeverity = update.MsrcSeverity
	If strSeverity = "" Then
		strSeverity = "NoSeverity"
	End If
	If strWUFilter = "All" Then
		log "INFO", "Adding " & getKBNumber(update.Title) & " :: " & strSeverity & " :: " & update.Title
		updatesToDownload.Add(update)
	ElseIf InStr(strWUFilter, strSeverity) > 0 Then
		log "INFO", "Adding " & getKBNumber(update.Title) & " :: " & strSeverity & " :: " & update.Title
		updatesToDownload.Add(update)
	End If
Next

'If downloader.Updates.Count = 0 Then
If updatesToDownload.Count = 0 Then
	strMessage = "No updates matching filter " & chr(34) & strWUFilter & chr(34)
	log "INFO", strMessage & ", exiting..."
	objLog.WriteLine "RESULT_KB:" & strMessage
	runPostUpdateActions 2, 0
End If
'
' /MAIN::BUILD_DOWNLOAD_LIST
'

'
' MAIN::DOWNLOAD_UPDATES_AND_PREPARE_INSTALL
'
log "INFO", "Downloading updates..."

Set downloader = updateSession.CreateUpdateDownloader()
checkError Err, 1, "Failed to create downloader object (Set downloader = updateSession.CreateUpdateDownloader())"
downloader.Updates = updatesToDownload
downloader.Download()

' And add them to the collection of updates to install
Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")

log "INFO", "List of downloaded updates:"

For I = 0 To downloader.Updates.Count-1
	Set update = downloader.Updates.Item(I)
	If update.IsDownloaded Then
		strSeverity = update.MsrcSeverity
		If strSeverity = "" Then
			strSeverity = "NoSeverity"
		End If
		log "INFO", "Downloaded/Ready to install " & getKBNumber(update.Title) & " :: " & strSeverity & " :: " & update.Title
		updatesToInstall.Add(update)
	End If
Next
'
' /MAIN::DOWNLOAD_UPDATES_AND_PREPARE_INSTALL
'

If bPROMPT Then 
	WScript.Echo vbCRLF & "Would you like to install updates now? (Y/N)"
	strInput = WScript.StdIn.Readline
	WScript.Echo
Else
	strInput = "Y"
End If

If (strInput = "N" or strInput = "n") Then
	WScript.Quit
ElseIf (strInput = "Y" or strInput = "y") Then
'
' MAIN::INSTALL_UPDATES
'
	WScript.Echo "Installing updates..."
	' The installation of all updates is done here
	log "INFO", "Installing updates..."
	
	Set installer = updateSession.CreateUpdateInstaller()
	installer.Updates = updatesToInstall
	Set installationResult = installer.Install()
	
	log "INFO", "Finished installing updates."
	' The updates are installed
'
' /MAIN::INSTALL_UPDATES
'	
'
' MAIN::CHECK_INSTALL_RESULT
'
	log "INFO", "Windows Update result :"
	' List installation results
	intOverallResult=0
	For I = 0 to updatesToInstall.Count - 1
		Set update = updatesToInstall.Item(I)
		strKBNumber = getKBNumber(update.Title)
		strSeverity = update.MsrcSeverity
		If strSeverity = "" Then
			strSeverity = "NoSeverity"
		End If
		strTitle = update.Title
		intResultCode = installationResult.GetUpdateResult(I).ResultCode
		If intResultCode > intOverallResult Then
			intOverallResult = intResultCode
		End If
		
		If bDEBUG Then log "DEBUG", "strKBNumber=" & strKBNumber & ", intResultCode=" & intResultCode
		strResultCode = convertResultCode(intResultCode)
		objLog.Writeline("RESULT_KB:" & strKBNumber & sep & strSeverity & sep & strTitle & sep & strResultCode)
	Next
	
	boolRebootRequired = installationResult.RebootRequired
	If boolRebootRequired Then
		boolRebootRequired = 1
	Else
		boolRebootRequired = 0
	End If
	If bDEBUG Then log "DEBUG", "boolRebootRequired=" & boolRebootRequired
	If bDEBUG Then log "DEBUG", "intOverallResult=" & intOverallResult
'
' /MAIN::CHECK_INSTALL_RESULT
'
	' and now end it !
	runPostUpdateActions intOverallResult, boolRebootRequired
End If

objLog.Close
wscript.quit
'
' /MAIN
' ---------------------------------------------------------------------

' ---------------------------------------------------------------------
' SUBS
'

' SUBS::log = log a message to objLog and or console
Sub log(strLevel, strMessage)
	dateNow = Now
	If LogLevel = 1 or LogLevel = 3 Then
		wscript.echo(dateNow & " > " & strLevel & ": " & strMessage)
	End If
	If LogLevel > 1 Then
		objLog.WriteLine(dateNow & " > " & strLevel & ": " & strMessage)
	End If
End Sub

' SUBS::restartService = stop then start strServiceName (this function is dirty and should be improved)
Sub restartService(strServiceName)
	' Restart Windows Update Service
	log "INFO", "Restarting service: " & strServiceName
	objShell.run "net stop " & strServiceName, 0, True
	wscript.sleep 1000
	objShell.run "net start " & strServiceName, 0, True
	wscript.sleep 1000
	log "INFO", "Restarted service: " & strServiceName
End Sub

' SUBS::checkError
Sub checkError(objErr, intCriticity, strMessage)
	If VarType(objErr) = 2 Then 
		ErrorNumber = objErr
		ErrorDescription = "ERROR_DEFAULT"
	Else 
		ErrorNumber = objErr.Number
		ErrorDescription = objErr.Description
	End If

	If ErrorNumber <> 0 Then
		log "ERROR", strMessage & " (ErrorNumber=" & ErrorNumber & ", ErrorDescription=" & ErrorDescription & ")"
		If intCriticity > 0 Then
			log "FATAL", "Cannot continue, exiting in 5 seconds."
			wscript.sleep 5000
			wscript.quit
		End If
	End if
End Sub

'
' SUBS::convertResultCode
Function convertResultCode(intCode)
	convertResultCode = "Unknown"
	Select Case intCode
		Case 0
			convertResultCode = "Not Started"
		Case 1
			convertResultCode = "In-Progress"
		Case 2
			convertResultCode = "Succeeded"
		Case 3
			convertResultCode = "Succeeded With Errors"
		Case 4
			convertResultCode = "Failed"
		Case 5
			convertResultCode = "Aborted"
	End Select
End Function

'
' SUBS::getKBNumber = return the KB number contained in strTitle
Function getKBNumber(strTitle)
	intLength = Len(strTitle)
	intIndex = InStrRev(strTitle, "(KB")
	If intIndex = 0 Then
		' No occurence found
		getKBNumber = "KB_UNKNOWN_"
	Else
		strKB = Mid(strTitle, intIndex + 1) ' + 1 removes leading "("
		getKBNumber = Left(strKB, Len(strKB) - 1) ' Len(strKB - 1) removes trailing ")"
	End If
End Function

'
' SUBS::runPostUpdateActions = finalize the log file, restore previous registry settings and 
Sub runPostUpdateActions(intOverallResult, boolRebootRequired)
	On Error Resume Next
	' Finalize log file
	strOverallResult = convertResultCode(intOverallResult)
	objLog.WriteLine("RESULT_OVERALL:" & strOverallResult)
	objLog.WriteLine("REBOOT_REQUIRED:" & boolRebootRequired)

	If strWUProxyConfig <> "_noWUProxyConfig_" Then
		strWUProxyConfig = vbRegedit(strWUProxyConfig, vbRegeditRestore)
		If bDEBUG Then log "DEBUG", "Restored proxy settings " & strWUProxyConfig
	End If
	
	' Start reboot if needed
	If boolRebootRequired = 1 Then
		log "INFO", "Launching reboot in 10 minutes (600 seconds)"
		strReboot = "shutdown /r /t 600 /c " & chr(34) & "WindowsUpdate requires this computer to reboot" & vbCrLf & "Reboot starting in 10 minutes..." & chr(34)
		errorCode = objShell.run(strReboot, 1, False)
		If bDEBUG Then log "DEBUG", "Started reboot process, errorCode=" & errorCode
	End If

	objLog.Close
	wscript.quit
End Sub

'
' SUBS::proxy2dec = prepare strWUProxyURL to be stored as binary value in the registry
' bin=WinHttpSettings="40,0,0,0,0,0,0,0,3,0,0,0," & Len(strWUProxyURL) & ",0,0,0," & str2dec(strWUProxyURL) & ",0,0,0,0"
Function proxy2dec(strWUProxyURL)
	intLen = Len(strWUProxyURL)-1
	ReDim arrChars(intLen)
	proxy2dec = "40:0:0:0:0:0:0:0:3:0:0:0:" & intLen + 1 & ":0:0:0"

	' Convert each char to its decimal value
	For i = 0 to intLen
		proxy2dec = proxy2dec & ":" & Asc(Mid(strWUProxyURL, i + 1,1))
	Next
	
	proxy2dec = proxy2dec & ":0:0:0:0"
End Function

'
' SUBS::vbRegedit = modify the registry given the intMode parameter (0 = set, 1 = restore)
Function vbRegedit(strRegParams, intMode)
	boolBackupExists = False
	If bDEBUG Then log "DEBUG", "In vbRegedit, intMode= " & intMode & ", strRegParams=" & strRegParams
	arrRegParams = Split(strRegParams, "|")
	strKeyRoot = CLng(arrRegParams(0))
	strKeyPath = arrRegParams(1)
	arrValues = Split(arrRegParams(2), ",")
	
	If intMode = vbRegeditRestore Then
		If UBound(arrRegParams) = 3 Then ' A backup of the keys exist
			boolBackupExists = True
			arrBackupValues = Split(arrRegParams(3), ",")
			If bDEBUG Then log "DEBUG", "A backup exists for " & strKeyRoot & "\" & strKeyPath
		Else
			log "ERROR", "no backup available for " & strKeyRoot & "\" & strKeyPath & " " & strValue
		End If
	End If
	errorCode = objReg.CreateKey(strKeyRoot, strKeyPath)
	
	For I = 0 To UBound(arrValues)
	'For Each strValue In arrValues
		arrValue = Split(arrValues(I), "=")
		strValueType = arrValue(0)
		strValueName = arrValue(1)
		strValueData = arrValue(2)
		
		If intMode = vbRegeditRestore and boolBackupExists Then ' Backup required and a backup of the keys exist
			log "INFO", "found backup for " & strKeyRoot & "\" & strKeyPath & ", restoring..."
			intBackupExists = 1
			arrBackupValue = Split(arrBackupValues(I), "=")
			strBackupValueType = arrBackupValue(0)
			strBackupValueName = arrBackupValue(1)
			strBackupValueData = arrBackupValue(2)
		End If
		
		If bDEBUG Then log "DEBUG", "strValueType=" & strValueType & ", strValueName=" & strValueName & ", strValueData=" & strValueData
		Select Case strValueType
			Case "dw"
				objReg.GetDWORDValue strKeyRoot,strKeyPath,strValueName,strValueRead
				Select Case intMode
					Case vbRegeditSet ' set values
						If IsNull(strValueRead) Then
							strValueRead = "_NOT_SET_"
						End If
						objReg.SetDWORDValue strKeyRoot,strKeyPath,strValueName,strValueData
					Case vbRegeditRestore ' restore values
						If Cstr(strBackupValueData) <> "_NOT_SET_" Then
							objReg.SetDWORDValue strKeyRoot,strKeyPath,strBackupValueName,strBackupValueData
						Else
							objReg.DeleteValue strKeyRoot,strKeyPath,strBackupValueName
						End If
					Case Else
						log "ERROR", "unknown intMode " & intMode & _
								", no modification on " & strKeyRoot & "\" & strKeyPath & " " & strValue
				End Select
			Case "str"
				objReg.GetStringValue strKeyRoot,strKeyPath,strValueName,strValueRead
				
				Select Case intMode
					Case vbRegeditSet ' set values
						If IsNull(strValueRead) Then
							strValueRead = "_NOT_SET_"
						End If
						objReg.SetStringValue strKeyRoot,strKeyPath,strValueName,strValueData
					Case vbRegeditRestore ' restore values
						If Cstr(strBackupValueData) <> "_NOT_SET_" Then
							objReg.SetStringValue strKeyRoot,strKeyPath,strBackupValueName,strBackupValueData
						Else
							objReg.DeleteValue strKeyRoot,strKeyPath,strBackupValueName
						End If
					Case Else
						log "ERROR", "unknown intMode " & intMode & _
								", no modification on " & strKeyRoot & "\" & strKeyPath & " " & strValue
				End Select
			Case "bin"
				objReg.GetBinaryValue strKeyRoot,strKeyPath,strValueName,arrValueRead
				'strWUProxyConfig bin=WinHttpSettings="40,0,0,0,0,0,0,0,3,0,0,0," & Len(strWUProxyURL) & ",0,0,0," & str2dec(strWUProxyURL) & ",0,0,0,0"
				Select Case intMode
					Case vbRegeditSet ' set values
						If IsNull(arrValueRead) Then
							strValueRead = "_NOT_SET_"
						Else
							strValueRead = Join(arrValueRead, ":")
						End If
						arrValueData = Split(strValueData, ":")
						objReg.SetBinaryValue strKeyRoot,strKeyPath,strValueName,arrValueData
					Case vbRegeditRestore ' restore values
						If Cstr(strBackupValueData) <> "_NOT_SET_" Then
							arrBackupValueData = Split(strBackupValueData, ":")
							objReg.SetBinaryValue strKeyRoot,strKeyPath,strBackupValueName,arrBackupValueData
						Else
							objReg.DeleteValue strKeyRoot,strKeyPath,strBackupValueName
						End If
					Case Else
						log "ERROR", "unknown intMode " & intMode & _
								", no modification on " & strKeyRoot & "\" & strKeyPath & " " & strValue
				End Select
		End Select
		If intMode = vbRegeditSet Then ' Backup current values
			If bDEBUG Then log "DEBUG", "Backuping " & strValueType & "=" & strValueName & "=" & strValueRead
			strBackupValues = strBackupValues & strValueType & "=" & strValueName & "=" & strValueRead & ","
		End If
	Next
	If intMode = vbRegeditSet Then ' Add backuped values to the return string
		strBackupValues = Left(strBackupValues, Len(strBackupValues) - 1) ' Remove trailing ","
		vbRegedit = strRegParams & "|" & strBackupValues
	Else 
		vbRegedit = strRegParams
	End If
	If bDEBUG Then log "DEBUG", "Updated registry, vbRegedit=" & vbRegedit
End Function

'
' /SUBS
' ---------------------------------------------------------------------