' ---------------------------------------------------------------------
' ABOUT
'
' ABOUT::Description
' Script to remotely run vbWSUS_SearchInstallDownload.vbs using PSExec
' Fetch remote log and parse it to a centralized log folder
'
' ABOUT::ExpectedArgs (must receive arguments from vbWSUS_scheduler.vbs)
' runDate;strHostname;strHostAddress;strPSExecUsername;strPSExecPassword;strPSExec;strRemotePath;strWUAUConfig;strWUServerConfig;strWUProxyConfig;strWUSearch;strWUFilter
'
' ABOUT::TellMeMore
' see README.txt and VERSION.txt
'
' /ABOUT
' ---------------------------------------------------------------------

bDEBUG = False

' ---------------------------------------------------------------------
' vbWSUS_CONFIGURATION
'
Set objFSO = wscript.CreateObject("Scripting.FileSystemObject")
strScriptDir = objFSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\"

includeFile strScriptDir & "conf\vbWSUS.conf"
includeFile strScriptDir & "lib\common.inc"
On Error Resume Next
'
' /vbWSUS_CONFIGURATION
' ---------------------------------------------------------------------

' ---------------------------------------------------------------------
' SCRIPT_CONFIGURATION
'
' SCRIPT_CONFIGURATION::LogFile
' LogFile : <hostname>.log will be appended to strLogFile later in the script
strLogFile = objFSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\logs\exec\vbWSUS_exec."
'
' /SCRIPT_CONFIGURATION
' ---------------------------------------------------------------------

' ---------------------------------------------------------------------
' MAIN
'
' MAIN::CHECK_CONFIGURATION
If boolConfigurationLoaded Then
	If bDEBUG Then log "DEBUG", "Successfully loaded configuration file vbWSUS.conf"
End If
' /MAIN::CHECK_CONFIGURATION
'
' MAIN::PARSE_ARGS
' Let's be sure we get all the command line parameters in a single string
If objArgs.Count = 12 Then
	runDate 			= objArgs(0)
	strHostname 		= objArgs(1)
	strHostAddress 		= objArgs(2)
	strPSExecUsername 	= objArgs(3)
	strPSExecPassword 	= objArgs(4)
	strPSExec 			= objArgs(5)
	strRemoteScript 	= objArgs(6)
	strWUAUConfig 		= objArgs(7)
	strWUServerConfig 	= objArgs(8)
	strWUProxyConfig 	= objArgs(9)
	strWUSearch 		= objArgs(10)
	strWUFilter 		= objArgs(11)
Else
	checkError ERROR_DEFAULT, 1, "Missing arguments !"
End If

'
' /MAIN::PARSE_ARGS
'

'
' MAIN::RUN
'
'objLog.Close
strLogFile = strLogFile & strHostName & ".log"
Set objLog = objFSO.OpenTextFile (strLogfile, ForAppending, CreateIfNotExists)

If Not objFSO.FileExists(strPSExec) Then 
	log "ERROR", "cannot find psexec.exe (" & strPSExec & ")"
	wscript.quit
End If

' we re ready
strPSExecParams = "\\" & strHostAddress & " -s " ' -s = Run the remote process with SYSTEM account
If strPSExecUsername <> "" Then
	strPSExecParams = strPSExecParams & "-u " & chr(34) & strPSExecUsername & chr(34) & " -p " & chr(34) & strPSExecPassword & chr(34) & " "
End If
strPSExecParams = strPSExecParams & "cscript.exe " & chr(34) & strRemoteScript & chr(34)

strPSExecArgs = quote(strWUAUConfig) 		& sep & _
				quote(strWUServerConfig) 	& sep & _
				quote(strWUProxyConfig) 	& sep & _
				quote(strWUSearch) 			& sep & _
				quote(strWUFilter)
					
log "INFO", "starting " & strRemoteScript & " (remote address: " & strHostAddress & ")"
If bDEBUG Then log "DEBUG", "		strPSExecUsername : " 	& strUsername
If bDEBUG Then log "DEBUG", "		strWUAUConfig : " 		& strWUAUconfig
If bDEBUG Then log "DEBUG", "		strWUServerConfig : " 	& strWUServerConfig
If bDEBUG Then log "DEBUG", "		strWUProxyConfig : " 	& strWUProxyConfig
If bDEBUG Then log "DEBUG", "		strWUSearch : " 		& strWUSearch
If bDEBUG Then log "DEBUG", "		strWUFilter : " 		& strWUFilter
	

' Send mail if required
If boolMailOnStart Then
	strMailSubject = strMailSubjectPrefix & " [ INFO ] [ " & strHostname & " ] starting update processus scheduled on " & runDate
	strMailContent = 	"Parameters are :"	& vbCrLf & _
						"	strPSExecUsername	=" & strPSExecUsername 	& vbCrLf & _
						"	strWUAUConfig		=" & strWUAUConfig		& vbCrLf & _
						"	strWUServerConfig	=" & strWUServerConfig	& vbCrLf & _
						"	strWUProxyConfig	=" & strWUProxyConfig	& vbCrLf & _
						"	strWUSearch			=" & strWUSearch		& vbCrLf & _
						"	strWUFilter			=" & strWUFilter		& vbCrLf
	sendMail strMailSubject, strMailContent, ""
End If

' Running the update on strHostname
If bDEBUG Then log "DEBUG", "strPSExecParams=" 	& strPSExecParams
If bDEBUG Then log "DEBUG", "strPSExecArgs=" 	& strPSExecArgs
errorCode = objShell.run(strPSExec & " " & strPSExecParams & " " & strPSExecArgs, 1, True)
If errorCode <> 0 Then
	' errorCode 2250 = ERROR_NOT_CONNECTED (This network connection does not exist.)
	' errorCode 1326 = ERROR_LOGON_FAILURE (Incorrect user name or password.)
	
	errorMessage = "Failed to run " & strRemoteScript & " on " & strHostname
	
	' Send mail if required
	If boolMailOnError Then
		strMailSubject = strMailSubjectPrefix & " [ ERROR ] [ " & strHostName & " ] strRemoteScript"
		strMailContent = errorMessage
		sendMail strMailSubject, strMailContent, ""
	End If
	
	' Exit script and raise error
	checkError errorCode, 1, errorMessage
Else
	log "INFO", "execution of " & strRemoteScript & " on " & strHostname & " ended with error code " & errorCode
End If
	
' Get the remote log
strRemoteLog = getParentFolder(strRemoteScript) 'objFSO.GetFile doesn't work as we can reach the strRemoteScript file from here
strRemoteLog = "\\" & strHostAddress & "\" & Replace(strRemoteLog, ":", "$") & "vbWSUS_SearchInstallDownload.log"
If bDEBUG Then log "DEBUG", "strRemoteLog=" & strRemoteLog
	
If Not objFSO.FileExists(strRemoteLog) Then
	errorCode = ERROR_DEFAULT
	errorMessage = "Failed to fetch " & strRemoteLog & " on " & strHostname
	
	' Send mail if required
	If boolMailOnError Then
		strMailSubject = strMailSubjectPrefix & " [ ERROR ] [ " & strHostName & " ] strRemoteLog"
		strMailContent = errorMessage
		sendMail strMailSubject, strMailContent, ""
	End If
	
	' Exist script and raise error
	checkError errorCode, 1, errorMessage
Else 
	'strDate = Replace(runDate, "/", "-")
	'strDate = Replace(strDate, ":", "-")
	'strDate = Replace(strDate, " ", "_")
	strDate = customDate(runDate, "YYYYMMDD_HHmm")
	strLocalLog = objFSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\logs\results\" & strHostname & "." & strDate & ".log"
	If bDEBUG Then log "DEBUG", "strLocalLog=" & strLocalLog
	
	log "INFO", "fetching remote log " & strRemoteLog & " to " & strLocalLog
	objFSO.CopyFile strRemoteLog, strLocalLog
	wscript.sleep 5000
	' Close network connection established by vbWSUS_scheduler.vbs
	objShell.Run "net use \\" & strHostAddress & "\ipc$ /DELETE", 1, True
	' Parse for reboot and/or errors
	If bDEBUG Then log "DEBUG", "Parsing " & strLocalLog
	Set objResultLog = objFSO.OpenTextFile(strLocalLog, ForReading)
	i = 0
	Do Until objResultLog.AtEndOfStream
		strLine = objResultLog.Readline
		If strLine <> "" Then
			If Left(strLine, 15) = "RESULT_OVERALL:" Then
				log "INFO", "updates finished with overall result " & strLine
				strResultOverall = Mid(strLine, 16)
				objDict.Add i, strLine
				i = i + 1
				' TODO: insert into DB
			ElseIf Left(strLine, 16) = "REBOOT_REQUIRED:" Then
				If Mid(strLine, 17) = "0" Then
					log "INFO", "no reboot required"
				Else
					log "INFO", "reboot required"
				End If
				objDict.Add i, strLine
				i = i + 1
				' TODO: insert into DB
			ElseIf Left(strLine, 10) = "RESULT_KB:" Then
				objDict.Add i, strLine
				i = i + 1
				'TODO: Insert into DB
			ElseIf InStr(strLine, "ERROR:") > 0 Then
				errorMessage = Mid(strLine, InStr(strLine, "ERROR:") + 6)
				log "ERROR", "Found error in " & strLocalLog & ": " & errorMessage
				objDict.Add i, "ERROR: " & strMessage
				i = i + 1
			End If
		End If
	Loop
	objResultLog.Close	
	
	If bDEBUG Then log "DEBUG", "Finished parsing " & strLocalLog

	' Send result email
	If boolMailOnResult Then
		strMailContent = ""
		strMailSubject = strMailSubjectPrefix & " [ INFO ] [ " & strHostName & " ] updates finished with status : " & strResultOverall
		' Get the values
		If bDEBUG Then log "DEBUG", "Parsing objDict"
		For Each strLine In objDict.Items
			'On Error Resume Next
			If bDEBUG Then log "DEBUG", "Adding line to strMailContent : " & strLine
			strMailContent = strMailContent & vbCrLf & strLine
		Next
		sendMail strMailSubject, strMailContent, ""	
	End If
End If
'
' /MAIN::RUN
'
objLog.Close
wscript.quit
'
' /MAIN
' ---------------------------------------------------------------------



' ---------------------------------------------------------------------
' SUBS
' subs specific to vbWSUS_exec.vbs

'
' SUB::includeFile
' Thanks to http://stackoverflow.com/questions/316166/how-do-i-include-a-common-file-in-vbscript-similar-to-c-include
sub includeFile (fSpec)
    dim fileSys, file, fileData
    set fileSys = createObject ("Scripting.FileSystemObject")
    set file = fileSys.openTextFile (fSpec)
    fileData = file.readAll ()
    file.close
    executeGlobal fileData
    set file = nothing
    set fileSys = nothing
end sub

'
' /SUBS
' ---------------------------------------------------------------------