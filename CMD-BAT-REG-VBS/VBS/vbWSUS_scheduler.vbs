' ---------------------------------------------------------------------
' ABOUT
'
' ABOUT::Description
' Main script of vbWSUS, checks serverlist.csv and if "Now" matches the specified cron in the input file
' - copy vbWSUS_SearchInstallDownload.vbs to remote host
' - start vbWSUS_exec.vbs which will handle the execution of vbWSUS_SearchInstallDownload.vbs
'
' ABOUT::ExpectedArgs
' _no arguments required_, see README.txt for more informations on command line arguments
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
										
' SCRIPT_CONFIGURATION::Logfile
strLogFile 			= objFSO.GetFile(Wscript.ScriptFullName).ParentFolder & "\logs\vbWSUS_scheduler." & customDate(Now, "YYYYMMDD") & ".log"
Set objLog 			= objFSO.OpenTextFile (strLogfile, ForAppending, CreateIfNotExists)
												
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
strArgs = ""

If objArgs.Count > 0 Then
	If bDEBUG Then log "DEBUG", "found " & objArgs.Count & " argument(s), using custom configuration."
	For I = 0 To objArgs.Count - 1
		' source = url
		If Left(objArgs(I), 11) = "source=http" Then
			strSourceType = "URL"
			strServerList = Mid(objArgs(I), 8)
		' source = file
		ElseIf Left(objArgs(I), 7) = "source=" Then
			strServerList = Mid(objArgs(I), 8)
		' PSExec account
		ElseIf Left(objArgs(I), 5) = "user=" Then
			strPSExecUser = Mid(objArgs(I), 6)
		ElseIf Left(objArgs(I), 9) = "password=" Then
			strPSExecPassword = Mid(objArgs(I), 10)
		' search parameters
		ElseIf Left(objArgs(I), 7) = "search=" Then
			strWUSearch = Mid(objArgs(I), 8)
			strWUSearch = Replace(strWUSearch, "&", " and ")
			strWUSearch = Replace(strWUSearch, "|", " or ")
		' filter parameters
		ElseIf Left(objArgs(I), 7) = "filter=" Then
			strWUFilter = Mid(objArgs(I), 8)
		' configure Automatic Updates
		ElseIf Left(objArgs(I), 13) = "config_au=yes" or Left(objArgs(I), 11) = "config_au=y" Then
			boolWUAU = True
		' WSUS server
		ElseIf Left(objArgs(I), 7) = "server=" Then
			strWUServerURL = Mid(objArgs(I), 8)
		' proxy server
		ElseIf Left(objArgs(I), 6) = "proxy=" Then
			strWUProxyURL = Mid(objArgs(I), 7)
		End If
	Next
Else
	If bDEBUG Then log "DEBUG", "no arguments found, using default configuration."
End If

If bDEBUG Then log "DEBUG",	"		strServerList=" 		& strServerList
If bDEBUG Then log "DEBUG", "		strPSExecUser=" 		& strPSExecUser
If bDEBUG Then log "DEBUG", "		strPSExecPassword=" 	& strPSExecPassword
If bDEBUG Then log "DEBUG", "		boolWUAU=" 				& boolWUAU
If bDEBUG Then log "DEBUG", "		strWUServerURL=" 		& strWUServerURL
If bDEBUG Then log "DEBUG", "		strWUProxyURL=" 		& strWUProxyURL
If bDEBUG Then log "DEBUG", "		strWUSearch=" 			& strWUSearch
If bDEBUG Then log "DEBUG", "		strWUFilter=" 			& strWUFilter
If bDEBUG Then log "DEBUG", "		strRemoteScript=" 		& strRemoteScript
'
' /MAIN::PARSE_ARGS
'


'
' MAIN::LOAD_SOURCE
'
If strSourceType = "URL" Then
	strServerList = getURL(strServerList)
	If Left(strServerList, 6) = "ERROR:" Then
		errorMessage = Mid(strServerList, 7)
		
		' Send mail if required
		If boolMailOnError Then
			strMailSubject = strMailSubjectPrefix & " [ ERROR ] load source"
			strMailContent = errorMessage
			sendMail strMailSubject, strMailContent, ""
		End If

		' Exit script and raise error
		checkError ERROR_DEFAULT, 1, errorMessage
	End If
	
	arrServerList = Split(strServerList, vbLf)
	i = 0
	For Each strLine in arrServerList
		' Skip empty line and line starting with "#" (comments)
		If strLine <> "" And Left(strLine, 1) <> "#" Then
			objDict.Add i, strLine
			i = i + 1
		End If	
	Next
	
Else
	If Not objFSO.FileExists(strServerList) Then
		errorMessage = "Could not find file " & strServerList
		' Send mail if required
		If boolMailOnError Then
			strMailSubject = strMailSubjectPrefix & " [ ERROR ] load source"
			strMailContent = errorMessage
			sendMail strMailSubject, strMailContent, ""
		End If
		checkError ERROR_DEFAULT, 1, errorMessage
	End If
	' Read CSV File
	Set objCSV = objFSO.OpenTextFile(strServerList, ForReading)
	i = 0
	Do Until objCSV.AtEndOfStream
		strLine = objCSV.Readline
		' Skip empty line and line starting with "#" (comments)
		If strLine <> "" And Left(strLine, 1) <> "#" Then
			objDict.Add i, strLine
			i = i + 1
		End If
	Loop
	objCSV.Close
End If
If bDEBUG Then log "DEBUG", "found " & i & " schedule(s)"
'
' /MAIN::LOAD_SOURCE
'

'
' MAIN::CHECK_CRON
'
For Each strLine in objDict.Items

	' Backup global configuration
	If bDEBUG Then log "DEBUG", "Restoring global configuration"
	backup_strPSExecUser		= strPSExecUser
	backup_strPSExecPassword	= strPSExecPassword
	backup_boolWUAU				= boolWUAU
	backup_strWUServerURL		= strWUServerURL
	backup_strWUProxyURL		= strWUProxyURL
	backup_strWUSearch			= strWUSearch
	backup_strWUFilter			= strWUFilter
	backup_strRemoteScript		= strRemoteScript
	
	' Get current equipment configuration
	If bDEBUG Then log "DEBUG", "Looking for custom configuration"
	arrEquipment = Split(strLine, ";")
	cron 			= arrEquipment(0)
	strDesc 		= arrEquipment(1)
	strHostname 	= arrEquipment(2)
	strHostAddress 	= arrEquipment(3)
	strUsername 	= arrEquipment(4)
	strPassword 	= arrEquipment(5)
	strProxy 		= arrEquipment(6)
	strServer 		= arrEquipment(7)
	intConfigAU 	= arrEquipment(8)
	strSearch 		= arrEquipment(9)
	strFilter 		= arrEquipment(10)
	strScript 		= arrEquipment(11)
	' All retrieved values are vbString

'
' MAIN::CHECK_CRON::RUN
'
	runDate = cron2date(cron)
	currentDate = Now
	If ( DateDiff("d", currentDate, runDate) = 0 and _
			DateDiff("m", currentDate, runDate) = 0 and _
			DateDiff("y", currentDate, runDate) = 0 and _
			DateDiff("h", runDate, currentDate) = 0 and _
			DateDiff("n", runDate, currentDate) = 0 ) Then
			
		If bDEBUG then log "DEBUG", "Found schedule to run:" & cron & " ( " & desc & " ) => " & runDate
		
		' Override "global configuration" if different parameters specified for this host
		' strPSExecUsername and strPSExecPassword
		If strUsername <> "" and strPassword <> "" Then
			If bDEBUG Then log "DEBUG", "Overriding default credentials (login/password) with login " & strUsername & " and password " & strPassword
			strPSExecUsername = strUsername
			strPSExecPassword = strPassword
		ElseIf strUsername <> "" and strPassword = "" Then
			log "ERROR", "Missing password, not overriding current credentials (strPSExecUsername=" & strPSExecUsername & ")"
		End If
		
		' strWUServerConfig	
		If strServer <> "" Then
			If bDEBUG Then log "DEBUG", "Overriding default WSUS server with " & strServer
			strWUServerConfig = Replace(strWUServerConfig, "_WSUS_URL_", strServer)
		ElseIf strWUServerURL <> "" Then
			strWUServerConfig = Replace(strWUServerConfig, "_WSUS_URL_", strWUServerURL)
		Else
			strWUServerConfig = "_noWUServerConfig_"
		End If	
		
		' strWUProxyConfig
		If strProxy <> "" Then
			If bDEBUG Then log "DEBUG", "Overriding default proxy with " & strProxy
			strWUProxyConfig = Replace(strWUProxyConfig, "_PROXY_URL_", strProxy)
		ElseIf strWUProxyURL <> "" Then
			strWUProxyConfig = Replace(strWUProxyConfig, "_PROXY_URL_", strWUProxyURL)
		Else
			strWUProxyConfig = "_noWUProxyConfig_"
		End If	
		
		' strWUAUConfig->UseWUServer value depends on strWUServerConfig
		If strWUServerConfig = "_noWUServerConfig_" Then
			strUseWUServer = "0"
		Else
			strUseWUServer = "1"
			boolWUAU = True 'Force config_au:yes
		End If
		strWUAUConfig = Replace(strWUAUConfig, "_USE_WSUS_", strUseWUServer)
		If Not boolWUAU and intConfigAU = "" Then
			strWUAUConfig = "_noWUAUConfig_"
		End If
		
		' strWUSearch
		If strSearch <> "" Then
			strWUSearch = strSearch
		End If
		
		' strWUFilter
		If strFilter <> "" Then
			strWUFilter = strFilter
		End If
		
		' strRemoteScript
		If strScript <> "" Then
			strRemoteScript = strScript
		End If
		
		' Build parameters to be passed to 'cscript.exe strExec'
		strExecParams = quote(runDate) 				& sep & _
						quote(strHostname)			& sep & _
						quote(strHostAddress) 		& sep & _
						quote(strPSExecUsername) 	& sep & _
						quote(strPSExecPassword) 	& sep & _
						quote(strPSExec) 			& sep & _
						quote(strRemoteScript) 		& sep & _
						quote(strWUAUConfig) 		& sep & _
						quote(strWUServerConfig)	& sep & _
						quote(strWUProxyConfig) 	& sep & _
						quote(strWUSearch) 			& sep & _
						quote(strWUFilter)

		If bDEBUG Then log "DEBUG", "strExecParams=" & strExecParams
		
		' Initialize remote connection
		strNetUse = "net use \\" & strHostAddress & "\ipc$ "
		If strPSExecUsername <> "" Then
			strNetUse = strNetUse & strPSExecPassword & " /USER:" & strPSExecUsername
		End If
		errorCode = objShell.run(strNetUse, 1, True)
		If errorCode <> 0 Then
			errorMessage = "Failed to establish IPC session to " & strHostname
			
			' Send mail if required
			If boolMailOnError Then
				strMailSubject = strMailSubjectPrefix & " [ ERROR ] [ " & strHostName & " ] copyFile"
				strMailContent = errorMessage
				sendMail strMailSubject, strMailContent, ""
			End If
			
			' Exit script and raise error
			checkError ERROR_DEFAULT, 1, errorMessage
		End If
		' If the errorCode is 0, go on :
		log "INFO", "Successfully established IPC session to " & strHostname

		' Copy vbWSUS_SearchInstallDownload.vbs to remote host
		strDestFile = "\\" & strHostAddress & "\" & Replace(strRemoteScript, ":", "$")
		strDestPath = getParentFolder(strDestFile) 'objFSO.GetFile doesn't work as we can reach the strRemoteScript file from here (not yet copied)
				
		createDirs strDestPath 'create remote folder (and subfolder(s))
		
		log "INFO", "Copying " & strLocalPath & " to " & strDestFile
		errorCode = objFSO.CopyFile(strLocalPath, strDestFile, Overwrite)
		' If the function succeeds, the return value is nonzero.
		' If the function fails, the return value is zero
		If VarType(errorCode) <> 0 Then
			errorMessage = "Failed to copy " & strLocalPath & " to " & strDestFile
			' Clear IPC connection
			strNetUseDelete = "net use \\" & strHostAddress & "\ipc$ /DELETE"
			errorCode = objShell.run(strNetUseDelete, 1, False) 'Don't wait for completion
			
			' Send mail if required
			If boolMailOnError Then
				strMailSubject = strMailSubjectPrefix & " [ ERROR ] [ " & strHostName & " ] copyFile"
				strMailContent = errorMessage
				sendMail strMailSubject, strMailContent, ""
			End If

			' Exit script and raise error
			checkError ERROR_DEFAULT, 1, errorMessage
		Else
			' Detach strExec from vbWSUS_scheduler.vbs to allow processing of next scheduled host
			If bDEBUG Then log "DEBUG", "strExecParams=" & strExecParams
			log "INFO", "Running updates on " & strHostname
			log "INFO", "Starting " & strExec & " (remote address: " & strHostAddress & ")"
			objShell.run "cscript.exe " & strExec & " " & strExecParams, 1, False
			checkError Err, 0, "Failed to start cscript.exe " & strExec
		End If
	End If
'
' /MAIN::CHECK_CRON::RUN
'
	' Restore "global configuration" for next loop
	strPSExecUser		= backup_strPSExecUser
	strPSExecPassword	= backup_strPSExecPassword
	boolWUAU			= backup_boolWUAU
	strWUServerURL		= backup_strWUServerURL
	strWUProxyURL		= backup_strWUProxyURL
	strWUSearch			= backup_strWUSearch
	strWUFilter			= backup_strWUFilter
	strRemoteScript		= backup_strRemoteScript
Next
'
' /MAIN::CHECK_CRON
'
objLog.Close
wscript.quit
'
' /MAIN
' ---------------------------------------------------------------------



' ---------------------------------------------------------------------
' SUBS
' subs specific to vbWSUS_scheduler.vbs
'

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
' SUBS::cron2date = convert a 'cron' formatted date to 'DD/MM/YYYY HH:mm'
Function cron2date(cron)
	' 0 1 * * 1#3
	arrCron = split(cron, " ")
	intMin = arrCron(0)
	intHour = arrCron(1)
	intDay = arrCron(2)
	intMonth = arrCron(3)
	intWeekday = arrCron(4)	
	
	' If no hour and/or min defined we say it's 00 and/or 00 (prevent to run again and again)
	If intMin = "*" Then
		intMin = "00"
	End If
	If intHour = "*" Then
		intHour = "00"
	End If
	
	' Find day and month
	If intMonth = "*" Then
		intMonth = Month(Date)
	End If
	If intWeekday <> "*" Then
		intWeekday = getDateByWeekday(intWeekday)
		intDay = Day(intWeekday)
		intMonth = Month(intWeekday)
	ElseIf intDay = "L" Then
		intDay = getLastDayOfMonth(intMonth)
	ElseIf intDay = "*" Then
		intDay = Day(Date)
	End If
	
	' Add a leading 0 if needed
	If Len(intMin) = 1 Then
		intMin = 0 & intMin
	End If
	If Len(intHour) = 1 Then
		intHour = 0 & intHour
	End If
	If Len(intMonth) = 1 Then
		intMonth = 0 & intMonth
	End If
	If Len(intDay) = 1 Then
		intDay = 0 & intDay
	End If

	cron2date = intDay & "/" & intMonth & "/" & Year(Date) & " " & intHour & ":" & intMin & ":00" ' 00 = seconds
End Function

'
' SUBS::getLastDayOfMonth = get the last day of 'intMonth'
' Thanks to http://www.tek-tips.com/viewthread.cfm?qid=1164880
Function getLastDayOfMonth(intMonth)
	getLastDayOfMonth = Day(DateSerial(Year(Now), 1 + intMonth, 0))
End Function

'
' SUBS::getDateByWeekDay = get the next date matching 'intWeekday'
Function getDateByWeekday(intWeekday)
	
	boolLast = InStr(intWeekday, "L")
	boolWeek = InStr(intWeekday, "#")
	
	If boolLast > 0 Then ' find last occurence of weekday in month
		arrWeekday = Split(intWeekday, "L")
		intWeekday = Cint(arrWeekday(0))
		
		' Find first occurence of 'weekday' on next month
		dateNextMonth = DateAdd("m", 1, Date)
		intNextMonth = Month(dateNextMonth)
		intYear = Year(dateNextMonth)
		firstdayofnextmonth = intYear & "-" & intNextMonth & "-01"		
		
		firstday = DatePart("w", firstdayofnextmonth) - 1 ' we start week at 0 in cron, not 1 (sunday is 0 in cron, 1 in vbscript)

		If intWeekday >= firstday Then
			delay = intWeekday - firstday
		Else
			delay = intWeekday - firstday + 7
		End If
		' The last intWeekday of a month is always 7 days before first occurence of intWeekday on next month
		getDateByWeekday = DateAdd("d", delay - 7, firstdayofnextmonth)

	ElseIf boolWeek > 0 Then ' find the nth weekday in month
		arrWeekday = Split(intWeekday, "#")
		intWeekday = Cint(arrWeekday(0))
		intWeek = Cint(arrWeekday(1) - 1)
		
		If intWeek > 3 Then
			getDateByWeekday = getDateByWeekday(intWeekday & "L")
		Else
			' Find first occurence of 'weekday'
			intMonth = Month(Date)
			intYear = Year(Date)
			firstdayofmonth = intYear & "-" & intMonth & "-01"

			firstday = DatePart("w", firstdayofmonth) - 1 ' we start week at 0 in cron, not 1 (sunday is 0 in cron, 1 in vbscript)

			If intWeekday >= firstday Then
				delay = intWeekday - firstday
			Else
				delay = intWeekday - firstday + 7
			End If
			' ^ first occurence of 'intWeekday' occurs in 'delay' days
			
			getDateByWeekday = DateAdd("d", intWeek * 7 + delay, firstdayofmonth)
		End If		
	Else ' Each intWeekday of the month, find next closest day to current date 
		currentweekday = DatePart("w", Date) - 1 ' we start week at 0 in cron, not 1 (sunday is 0 in cron, 1 in vbscript)

		If currentweekday = Cint(intWeekday) Then 'it's today
			delay = 0
		ElseIf Cint(intWeekday) > currentweekday Then 'we re early in the week
			delay = intWeekday - currentweekday
		Else ' next weekday is next week
			delay = intWeekday - currentweekday + 7		
		End If
		
		getDateByWeekday = DateAdd("d", delay, Now)
	End If

End Function

'
' SUBS::URLGet = get the content of 'URL'
Function getURL(strURL)
	strURL_METHOD = "MSXML2.ServerXMLHTTP," & _
					"WinHttp.WinHttpRequest.5.1," & _
					"WinHttp.WinHttpRequest.5," & _
					"WinHttp.WinHttpRequest," & _
					"MSXML2.XMLHTTP," & _
					"Microsoft.XMLHTTP"
	arrURL_METHOD = Split(strURL_METHOD, ",")
	For Each strMethod In arrURL_METHOD
		On Error Resume Next
		Set objHttp = CreateObject(strMethod)
		' strMethod works !
		If Err.Number = 0 Then Exit For
	Next
	'Set objHttp = CreateObject(getURL_METHOD)
	' Quit if unable to create the object with the current getURL_METHOD
	'checkError Err, 1, "Could not create objHttp, CreateObject(" & getURL_METHOD & ") returned error code __ERROR_NUMBER__ (description: " & __ERROR_DESCRIPTION__ & ")"
	objHttp.Open "GET",strURL,True
	objHttp.Send
	wscript.sleep 10000 ' wait for response
	strStatus = objHttp.status
	If strStatus <> "200" Then
		getURL = "ERROR:failed to retrieve " & strURL & " (HTTP status=" & strStatus & ")"
	Else
		getURL = objHttp.responseText
	End if
End Function


'
' SUBS::CreateDirs
' Thanks to http://www.robvanderwoude.com
Sub createDirs( MyDirName )
' This subroutine creates multiple folders like CMD.EXE's internal MD command.
' By default VBScript can only create one level of folders at a time (blows
' up otherwise!).
'
' Argument:
' MyDirName   [string]   folder(s) to be created, single or
'                        multi level, absolute or relative,
'                        "d:\folder\subfolder" format or UNC
'
' Written by Todd Reeves
' Modified by Rob van der Woude
' http://www.robvanderwoude.com

    Dim arrDirs, i, idxFirst, objFSO, strDir, strDirBuild

    ' Create a file system object
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ' Convert relative to absolute path
    strDir = objFSO.GetAbsolutePathName( MyDirName )

    ' Split a multi level path in its "components"
    arrDirs = Split( strDir, "\" )

    ' Check if the absolute path is UNC or not
    If Left( strDir, 2 ) = "\\" Then
        strDirBuild = "\\" & arrDirs(2) & "\" & arrDirs(3) & "\"
        idxFirst    = 4
    Else
        strDirBuild = arrDirs(0) & "\"
        idxFirst    = 1
    End If

    ' Check each (sub)folder and create it if it doesn't exist
    For i = idxFirst to Ubound( arrDirs )
        strDirBuild = objFSO.BuildPath( strDirBuild, arrDirs(i) )
        If Not objFSO.FolderExists( strDirBuild ) Then 
            objFSO.CreateFolder strDirBuild
        End if
    Next

    ' Release the file system object
    Set objFSO= Nothing
End Sub

'
' /SUBS
' ---------------------------------------------------------------------
