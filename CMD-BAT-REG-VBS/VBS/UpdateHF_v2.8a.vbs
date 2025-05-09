strScriptVer = "2.8a - Rob Dunn @ http://community.spiceworks.com"
'~~[author]~~
'Rob Dunn
'with some additional improvements made by the WSUS and Spiceworks community
' - thank you!
'~~[/author]~~
'
'~~[ContactInfo]~~
'https://community.spiceworks.com/people/robdunn
'~~[/ContactInfo]~~
'
'~~[website]~~
'For future versions and support:
'https://community.spiceworks.com/scripts/show/82
'~~[/website]~~
'
'~~[scriptType]~~
'vbscript
'~~[/scriptType]~~

'~~[subType]~~
'SystemAdministration
'~~[/subType]~~

'~~[keywords]~~
'wsus, windows, updates, hotfixes, email, windowsupdate, microsoft, wua, sus, 
'~~[/keywords]~~
'
'~~[usage]~~
'****Install updates silently, email you a logfile, then restart the computer****  
'updatehf.vbs action:install mode:silent email:you@yourdomain.com restart:1 
' 
'****Install KB4012213 silently, email you a logfile, then restart the computer****  
'updatehf.vbs action:install mode:silent updateid:80BC2B42-A953-4096-8595-130E9A9C9FB9 email:you@yourdomain.com restart:1 
'
'****Detect missing updates, email you a logfile, then do nothing (no restart)****
'updatehf.vbs action:detect mode:verbose email:you@yourdomain.com restart:0 
' 
'****Prompt user to let them decide whether or not to install updates, email 
' you a logfile, prompt user for restart**** 
'updatehf.vbs action:prompt mode:verbose email:you@yourdomain.com restart:1 
' 
'****Install updates silently, email you a logfile, then shutdown the computer 
' if a reboot is pending**** 
'updatehf.vbs action:install mode:silent email:you@yourdomain.com restart:2 
' 
'****Install updates silently, email you a logfile, then shutdown the computer 
' no matter if a reboot is pending or not****
'updatehf.vbs action:install mode:silent email:you@yourdomain.com restart:2 force:1 
' 
'****Detect missing updates or pending reboot silently, email you a logfile, then 
' restart if there is a pending reboot****
'updatehf.vbs action:detect mode:silent email:you@yourdomain.com restart:1 
' 
'****Detect missing updates or pending reboot silently, email you a logfile, then 
' restart no matter if there is a pending reboot****
'updatehf.vbs action:detect mode:silent email:you@yourdomain.com restart:1 force:1 
'
'~~[/usage]~~
'This script (the core was pulled from Microsoft's website - thank 
' you!) will tell the WU agent to 'detectnow', download and install 
' missing windows updates as compared to it's update server. Works for 
' WSUS and regular Windows Update site. 
'
'This will now reboot the computer if specified after the udpates have 
' been applied (or if there is a reboot pending from a previous update session). 
'
'NOTE: If there are a LOT of downloads to pull, the status window (or log) 
' will say "Downloading" for that entire time. I'm not sure how to get 
' a download progress of each update...maybe someone can help me with that. 
'
'Note on command-line switches: If you don't specify a switch (for 
' example, 'email:') the corresponding variable defined in the script will 
' provide the needed information (command-line switches take precedence). 
'
'Why I put this script together: 
'Our desktop deployment technicians needed a script that would pull 
' updates immediately and install, even if the update configuration is set 
' for 'download only' via Group Policy. 
'
'We have some computers that are sometimes logged on or not (but 
' they run services that must be running almost constantly), and are never rebooted. 
'
'The user ignores the 'you have new updates available' message, so updates are 
' never installed. This script will let you install the updates, and then it 
' tells the WUA to present the 'restart' message - which more users are apt to 
' respond to. 
'
'After the script runs, it will email a recipient the resulting logfile that is 
' produced. 
'
'*******************************************************************************
'You need to edit the following variables: 
'
'strMailFrom - arbitrary reply-to address 
'strMailto - email address you want the report to mail to (this is for manual 
' mode 
' - or if the command-line switch isn't specified). 
'strSMTPServer - the IP address of the email server you are sending the reports 
' through. 
'
'*******************************************************************************
'Optional variables: 
'Silent - 0 = verbose, 1 = silent (no windows or visible information) 
'Intdebug - 0 = off, 1 = 1 (see some variables that are being passed) 
'strAction - prompt|install|detect. Prompt gives users opportunity to install 
' updates or not, install just installs them, detect updates the WU collection 
' and downloads the updates (but does not install them) - useful if you want to 
' have the computer refresh its stats to the stat server but not install the 
' updates. 
'blnEmail - 0 = off|1 = on. If set to 0, the script will not email a log file. 
' If you specify an email address in the command-line, this will force the 
' script to switch blnEmail to '1'. 
'strRestart - 0 = Do nothing|1 = restart|2 = shutdown. Command-switch 'restart:' 
' supercedes this variable. 
'
'*******************************************************************************
'Command line switches: 
'action: prompt|install|detect
'updateid: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx (look up the updateID using the Microsoft Update Catalog) 
'mode: silent|verbose 
'email: you@yourdomain.com 
'restart: 0 (do nothing)| 1 (restart) | 2 (shutdown) 
'force: 0 (do not enforce restart action - this is optional, by default it is 
' set to 0) | 1 (enforce restart action).
'SMTPServer: x.x.x.x or hostname; overrides strSMTPServer above.
'emailsubject: "this is a subject" Overrides default subject. Server name is appended to this text. Use quotes if spaces exist.
'emailifallok: 0|1, where 0 = dont send email if server up to date and no 
' reboot pending, and 1 = always send email
'fulldnsname: 0|1, where 0 = use server name only in subject, and 1 = use full
' dns name in email subject
'authtype: cdoAnonymous|cdoNTLM|cdoBasic - authentication type for SMTP (default
' is 'cdoAnonymous', no creds needed)
'authID: SMTP authentication ID
'authPassword: SMTP authentication password
'Finally, rename the file with .vbs 
'*******************************************************************************
Dim objArgs,strArgs
Dim l
Set objArgs = Wscript.Arguments

'Get all arguments
For Each strArg in objArgs
  strArguments = strArguments & " " & strArg
Next

'Handle UAC
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , WScript.ScriptFullName & " /elevate" & " " & strArguments, "", "runas", 1
  WScript.Quit
End If

Const HKEY_CURRENT_USER 			= &H80000001
Const HKEY_LOCAL_MACHINE 			= &H80000002
Const ForAppending 					= 8
Const ForWriting 					        = 2
Const ForReading 					= 1
Const cdoAnonymous 				= 0 'Do not authenticate
Const cdoBasic 						= 1 'basic (clear-text) authentication
Const cdoNTLM 						= 2 'NTLM
Const cdoSendUsingMethod 			= "http://schemas.microsoft.com/cdo/configuration/sendusing", _
			cdoSendUsingPort 		= 2, _
			cdoSMTPServer 			= "http://schemas.microsoft.com/cdo/configuration/smtpserver", _
			cdoSMTPServerport 		= "http://schemas.microsoft.com/cdo/configuration/smtpserverport", _
			cdoSMTPconnectiontimeout = "http://schemas.microsoft.com/cdo/configuration/Connectiontimeout"

'Web address to refer users for unhandled error codes
strAddr = "https://support.microsoft.com/en-us/kb/938205"

On Error Resume Next

'below variables for progress indicators
Dim objShell, objProcessEnv, objSystemEnv, objNet, objFso, objSwitches
Dim query, item, acounter, blnExtendedWMI, blnProcessEvents
Dim dlgBarWidth, dlgBarHeight, dlgBarTop, dlgBarLeft, dlgProgBarWidth, dlgProgBarHeight 
Dim dlgProgBarTop, dlgProgBarLeft
Dim dlgBar, dlgProgBar, wdBar, objPBar, objBar, blnSearchWildcard
Dim blnProgressMode, blnDebugMode, dbgTitle
Dim dbgToolBar, dbgStatusBar, dbgResizable
Dim IE, objDIV, objDBG, strMyDocPath, strSubFolder, strTempFile, f1, ts, File2Load, objFlash
Dim dbgWidth, dbgHeight, dbgLeft, dbgTop, dbgVisible
'above variables for progress indicators
Dim blnRebootRequired

Dim strAction, regWSUSServer, ws, wshshell, wshsysenv, strMessage, strFrom

Dim strRestart, silenttext, restarttext, blnCallRestart, blnInstall, blnPrompt, strStatus
Dim blnIgnoreError, blnCScript, strLocaleDelim

Set WshShell = WScript.CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("PROCESS")
Set ws = wscript.CreateObject("Scripting.FileSystemObject")
Set objADInfo = CreateObject("ADSystemInfo")

'Try to pick up computername via AD'
strComputer1 = objADInfo.ComputerName

'As a backup, use the environment strings to pick up the computer
' name.
strComputer = wshShell.ExpandEnvironmentStrings("%Computername%")

strUser = WshSysEnv("username")
strDomain = WshSysEnv("userdomain")

'Get computer OU
strOU = "Computer OU: Not detected"

Set objComputer = GetObject("LDAP://" & strComputer1)
If objComputer.Parent <> "" Then  
	strOU = "Computer OU: " & replace(objComputer.Parent,"LDAP://","")
End If

If InStr(ucase(WScript.FullName),"CSCRIPT.EXE") Then
	blnCScript = TRUE
Else
	blnCScript = FALSE
End If

blnCloseIE = true

'*******************************************************************************
' User variables
'*******************************************************************************
'Turn on debugging.  This will show some of the variables that are being passed 
' while the script executes.
Intdebug = 0          

'How long between the time that the script is finished and the IE window stays
' on the screen.  Set to '0' if you don't want the status window to close 
' automatically.
intSleep = 2000

'Whether or not the user will see the status window.
' Possible options are: 
'0 = verbose, progress indicator, status window, etc.
'1 = silent, no progress indicators.  Everything occurs in the background
Silent = 0
      
'The location of the logfile (this is the file that will be parsed
' and the contents will be sent via email.                      
logfile = WshSysEnv("TEMP") & "\" & "vbswsus-status.log"
                                      
'arbitrary email address - reply-to
strMailFrom = "wsus_script@creditone.com"

'who are you mailing to?  Input mode only.  Command-line parameters take 
' precedence
strMailto = "winsysadmin@creditone.com;noc@creditone.com"

'set SMTP email server address here
strSMTPServer = "lasexch01.Contoso.corp"

'set SMTP email server port (default is 25)
iSMTPServerPort = 25

'The computer name will follow this text when the script completes.
strSubject = "[WSUS Update Script] - WSUS Update log file from" 

'Deliminator in above strWUAgentVersion - some locales might have "," instead
' (Non English) - leave as "." if you aren't sure.
strLocaleDelim = "."

'default option for manual run of the script.  Possible options are:
' prompt - (user is prompted to install)
' install - updates will download and install
' detect - updates will download but not install                                       
strAction = "install" 

'Turns email function on/off.  If an email address is specified in the 
' command-line arguments, then this will automatically turn on ('1').
' 0 = off, don't email
' 1 = on, email using default address defined in the var 'strMailto' above.
blnEmail = 1

'strEmailIfAllOK Determines if email always sent or only if updates or reboot 
' needed.
' 0 = off, don't send email if no updates needed and no reboot needed
' 1 = on always send email
strEmailIfAllOK = 1

'strFullDNSName Determines if the email subject contains the full dns name of 
' the server or just the computer name.
' 0 = off, just use computer name
' 1 = on,  use full dns name
strFullDNSName = 1

'tells the script to prompt the user (if running in verbose mode) to input the 
' email address of the recipient they wish to send the script log file to.  The 
' default value in the field is determined by the strMailto variable above.
' 
'This only appears if no command-line arguments are given.  
'0 = do not prompt the user to type in an email address
'1 = prompt user to type in email address to send the log to.
promptemail = 0

'Tells the computer what to do after script execution if the script detects that 
' there is a pending reboot.
'
'Command-prompt supercedes this option.
'0 = do nothing
'1 = reboot
'2 = shutdown
strRestart = 1

'Try to force the script to work through any errors.  Since some are recoverable
' this might be an option for troubleshooting.  Default is 'true'
blnIgnoreError = true

'sets font for display status dialog and sent formatted logfile
strFontStyle = "arial"

'set your SMTP server authentication type.  
' Possible values:cdoAnonymous|cdoBasic|cdoNTLM
' You do not need to configure an id/pass combo with cdoAnonymous
strAuthType = "cdoAnonymous"

'SMTP authentication ID
strAuthID = ""

'Password for the ID
strAuthPassword = ""

'*******************************************************************************
'End of User variables
'*******************************************************************************

'writelog("Arguments: " & wscript.arguments)
writelog("Log file used: " & logfile)
If intdebug = 1 then wscript.echo "Objargs.count = " & objArgs.count

If objArgs.Count > 0 Then
For I = 0 to objArgs.Count - 1
  If objArgs.Count > 0 Then
    if instr(LCase(objargs(i)),"action:") Then
      strArrAction = split(objargs(i),":")
      strAction = strArrAction(1)
      if intdebug = 1 then wscript.echo strAction
	  'wscript.quit
    ElseIf instr(LCase(objargs(i)),"mode:") Then 
      strArrMode = split(objargs(i),":")
      silent = strArrMode(1)
      If lcase(silent) = "silent" then 
        silent = 1
      Elseif lcase(silent) = "verbose" then 
        silent = 0
        blnCloseIE = true
      Else
        strMsg = "Invalid mode switch: " & silent & ".  Now aborting."
        'Call ErrorHandler("Command Switches",strMsg,"true")
      End If
      Silenttext = strArrMode(1)
    ElseIf instr(LCase(objargs(i)),"email:") Then 
	   strArrEmail = split(objargs(i),":") 
	   strMailto = strArrEmail(1) 
	   blnEmail = 1 
	 	ElseIf instr(LCase(objargs(i)),"logfile:") Then 
	   strArrLogfile = split(objargs(i),"logfile:") 
	   LogFile = strArrLogFile(1) 
    ElseIf InStr(LCase(objargs(i)),"restart:") Then 
    	strArrAction = split(objargs(i),":")
    	strRestart = strArrAction(1)
    ElseIf InStr(LCase(objargs(i)),"force:") Then
    	strArrForceAction = split(objargs(i),":")
    	strForceaction = strArrForceAction(1)
    ElseIf InStr(LCase(objargs(i)),"smtpserver:") Then
    	strArrSMTPServer = split(objargs(i),":")
    	strSMTPServer = strArrSMTPServer(1)
    ElseIf InStr(LCase(objargs(i)),"emailifallok:") Then
    	strArrEmailIfAllOK = split(objargs(i),":")
    	strEmailIfAllOK = strArrEmailIfAllOK(1)
    ElseIf InStr(LCase(objargs(i)),"fulldnsname:") Then
    	strArrFullDNSName = split(objargs(i),":")
    	strFullDNSName = strArrFullDNSName(1)
    ElseIf InStr(LCase(objargs(i)),"emailsubject:") Then
    	strArrSubject = split(objargs(i),":")
    	strSubject = strArrSubject(1)
    ElseIf InStr(LCase(objargs(i)),"authtype:") Then
    	strArrAuthType = split(objargs(i),":")
    	strAuthType = strArrAuthType(1)
    ElseIf InStr(LCase(objargs(i)),"authid:") Then
    	strArrAuthID = split(objargs(i),":")
    	strAuthID = strArrAuthID(1)
    ElseIf InStr(LCase(objargs(i)),"authpassword:") Then
    	strArrAuthPassword = split(objargs(i),":")
    	strAuthPassword = strArrAuthPassword(1)  
	ElseIf Instr(LCase(objargs(i)),"updateid:") Then
		strArrKBID = split(objargs(i),":")
		strKBID = strArrKBID(1)
   End If
  End If
Next
Else
      'strAction = "prompt"
      If blnEmail = 1 and silent = 0 and promptemail = 1 Then strMailto = InputBox("Input the email address you would like the " _
       & "Windows Update agent log sent to:","Email WU Agent logfile to recipient",strMailto)
      If strMailto = "" Then wscript.quit
End If

Set l = ws.OpenTextFile (logfile, ForWriting, True)
l.writeline "------------------------------------------------------------------"
l.writeline "WU force update VBScript" & vbcrlf & Now & vbcrlf & "Computer: " & strComputer
l.writeline "Script version: " & strScriptVer
l.writeline strOU 

l.writeline "Executed by: " & strDomain & "\" & strUser
l.writeline "Command arguments: " & strArguments
l.writeline "------------------------------------------------------------------"
	'Call WMI query to collect parameters for reboot action
	Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//" & strComputer & "/root/cimv2").ExecQuery("select caption,OSArchitecture,ServicePackMajorVersion from Win32_OperatingSystem"_
	 & " where Primary=true") 
		For each item in OpSysSet 
			strOS = item.caption
			strOSArchitecture = item.OSArchitecture
			strSP = item.ServicePackMajorVersion
		Next 
		
		writelog("Operating System: " & strOS)
		writelog("Service Pack: " & strSP)
		writelog("OS Architecture: " & strOSArchitecture)

If blnEmail = 1 then 
    writelog("SMTP Authentication type specified: " & strAuthType)
    If lcase(strAuthType) <> "cdoanonymous" Then
      If strAuthType = "" Then
        strAuthType = "cdoanonymous"
      Else
        writelog("SMTP Auth User ID: " & sAuthID)
    
        If SMTPUserID = "" then 
          writelog("No SMTP user ID was specified, even though SMTP Authentication was configured for " & strAuthType & ".  Attempting to switch to anonymous authentication...")
          strAuthType = "cdoanonymous"
          If strAuthPassword <> "" then writelog("You have specified a SMTP password, but no user ID has been configured for authentication.  Check the INI file (" & sINI & ") again and re-run the script.")
        Else
          if strAuthPassword = "" then writelog("You have specified a SMTP user ID, but have not specified a password.  Switching to anonymous authentication.")
          strAuthType = "cdoanonymous"
        End if
        If strAuthPassword <> "" then writelog("SMTP password configured, but hidden...")
    
      End If
    End If
End If

'Call checkupdateagent

Select Case silent
  Case 0
    silenttext = "Verbose"
  Case 1
    silenttext = "Silent"
  Case Else
End Select

If strForceaction = 1 Then 
	strForceText = " (enforce action)"
Else
	strForceText = " (only if action is pending)"
End If

Select Case strRestart
  Case 0
    restarttext = "Do nothing"
  Case 1 
    restarttext = "Restart"
  Case 2 
    restarttext = "Shut down"
  Case Else
End Select

restarttext = restarttext & strForceText

writelog("Script action is set to: " & strAction)
writelog("Verbose/Silent mode is set to: " & silenttext)
writelog("Restart action is set to: " & restarttext)

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
 strComputer & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
strValueName = "WUServer"
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,regWSUSServer
writelog("Checking local WU settings...")

Call GetAUSchedule()

If regWSUSServer then 
	
Else
	regWSUSServer = "http://windowsupdate.microsoft.com/"
End If

writelog("Update Server: " & regWSUSServer)
writelog("HTTP Status: " & URLGet(regWSUSServer))

strValueName = "TargetGroup"

oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,regTargetGroup
if regTargetGroup <> "" then 
  writelog("Target Group: " & regTargetGroup)
Else
  writelog("Target Group: Not specified")
End If

Set autoUpdateClient = CreateObject("Microsoft.Update.AutoUpdate")
Set updateInfo = autoUpdateClient.Settings

Select Case updateInfo.notificationlevel
	Case 0 
	  writelog("WUA mode: WU agent is not configured.")
	Case 1 
	  writelog("WUA mode: WU agent is disabled.")
	Case 2
	  writelog("WUA mode: Users are prompted to approve updates prior to installing")
	Case 3 
	  writelog("WUA mode Updates are downloaded automatically, and users are prompted to install.")
	Case 4 
	  writelog("WUA mode: Updates are downloaded and installed automatically at a predetermined time.")
	Case Else
End Select


fstyle = "calibri,tahoma,arial,verdana"
bgcolor1 = "aliceblue"
fformat = "<font face='" & fstyle & "'>"

'set some IE status indicator variables...
blnDebugMode = True
blnProcessEvents = True
blnSearchWildcard = False
blnProgressMode = True

On Error Resume Next

Set updateSession = CreateObject("Microsoft.Update.Session")
Set updateSearcher = updateSession.CreateupdateSearcher()

writelog("Instantiating Searcher")
If strKBID <> "" Then
	writelog("Searching for specific UpdateID: " & strKBID)
	strQuery = "UpdateID='" & strKBID & "'"
Else 
	writelog("Checking for all approved updates according to WU agent")
	strQuery = "IsAssigned=1 and IsHidden=0 and IsInstalled=0 and Type='Software'"
End If

writelog("Searcher query: " & strQuery)
Set searchResult = updateSearcher.Search(strQuery)


'Handle some common errors here
If cstr(err.number) <> 0 Then
  If cstr(err.number) = "-2147012744" Then
    strMsg = "ERROR_HTTP_INVALID_SERVER_RESPONSE - The server response could not be parsed." & vbcrlf & vbcrlf & "Actual error was: " _
      & " - Error [" & cstr(err.number) & "] - '" & err.description & "'"
    blnFatal = true
  ElseIf Cstr(err.number) = "-2145107952" Then
    strMsg = "WU_E_PT_EXCEEDED_MAX_SERVER_TRIPS The number of round trips to the server exceeded the maximum limit. " _
     & "Stop/Restart service or reboot the machine if you see this error frequently. " _
     & vbcrlf & vbcrlf & "Actual error was [" & err.number & "] - " & chr(34) _
      & err.description & chr(34)
    blnFatal = true	
  ElseIf CStr(err.number) = "-2145107924" Then
    strMsg = "WU_E_PT_WINHTTP_NAME_NOT_RESOLVED - Winhttp SendRequest/ReceiveResponse failed with 0x2ee7 error. Either the proxy " _
     & "server or target server name can not be resolved. Corresponding to ERROR_WINHTTP_NAME_NOT_RESOLVED. " _
     & "Stop/Restart service or reboot the machine if you see this error frequently. " _
     & vbcrlf & vbcrlf & "Actual error was [" & err.number & "] - " & chr(34) _
      & err.description & chr(34)
    blnFatal = false
  ElseIf cstr(err.number) <> 0 and cstr(err.number) = "-2147012867" Then 
    strMsg = "ERROR_INTERNET_CANNOT_CONNECT - The attempt to connect to the server failed." & vbcrlf _
      & vbcrlf & "Actual error was [" & err.number & "] - " & chr(34) _
      & err.description & chr(34)
    blnFatal = true
  ElseIf CStr(err.number) = "-2145107941" Then 
    strMsg = "SUS_E_PT_HTTP_STATUS_PROXY_AUTH_REQ - Http status 407 - proxy authentication required" & vbcrlf & vbcrlf & "Actual " _
     & "error was [" & err.number & "]" & chr(34) & err.description & chr(34)
  ElseIf CStr(err.number) = "-2145124309" Then 
    strMsg = "WU_E_LEGACYSERVER - The Sus server we are talking to is a Legacy Sus Server (Sus Server 1.0)" _
     & vbcrlf & vbcrlf & "Actual error was [" & err.number & "] - " & chr(34) & err.description & chr(34)
    blnFatal = true
  ElseIf CStr(err.number) = "7" Then 
    strMsg = "Out of memory - In most cases, this error will be resolved by rebooting the client." _ 
     & VbCrLf & VbCrLf & "Actual error was [" & err.number & "] - " & chr(34) & err.description & chr(34) 
    blnFatal = True 
  Else
    If err.description = "" Then 
    	errdescription = "No error description given"
    Else 
        errdescription = err.description
    End If
    If blnIgnoreError = false Then 
    	blnFatal = true 
 	    strScriptAbort = vbcrlf & vbcrlf & "Script will now abort. - if you want to force the script to continue, change the 'blnIgnoreError' variable " _
     	 & "to the value 'true'"
    Else
    	strScriptabort = vbcrlf & vbcrlf & "Script will attempt to continue."
    End If
    
    strMsg = "Error - [" & err.number & "] - " & chr(34) & errdescription & chr(34) & "." & vbcrlf & vbcrlf _
     & "This error is undefined in the script, but you can refer to " & strAddr & " to look up the error number." _
     & strScriptAbort
     strMsgHTML = replace(strMsg,strAddr,"<a href='" & strAddr & "'>" & strAddr & "</a>")
    If silent = 0 Then objdiv.innerhtml = replace(strMsgHTML,"vbcrlf","<br>")
   End If
  
  Call ErrorHandler("UpdateSearcher",strMsg,blnFatal)
End If
'ssManagedServer

If silent = 0 then 
  writelog("Calling IE Status Window")
	on error goto 0
	Call IEStatus
End If

Call CheckPendingStatus("beginning")

strMsg = "Refreshing WUA client information..."
if silent = 0 then objdiv.innerhtml = strMsg

'cause WU agent to detect
on error resume next
autoUpdateClient.detectnow()
if err.number <> 0 then call ErrorHandler("WUA refresh",err.number & " - " & err.description,false)
err.clear
on error goto 0 

strMsg = "WUA mode: <font color='navy'>" & strACtion & "</font><br>WU Server: " & regWSUSServer _
 & "<br>Target Group: " & regTargetGroup & "<br><br>List of applicable items on the machine: <br>"
writelog("WUA mode: " & straction)
writelog("WU Server: " & regWSUSServer)

If (regWSUSServer = "http://windowsupdate.microsoft.com/") Then 

Else
	writelog(regWSUSServer & "/iuident.cab Status: " & URLGet(regWSUSServer & "/iuident.cab"))
End If


If silent = 0 then objdiv.innerhtml = strMsg
If strAction <> "detect" Then 
	writelog("Searching for missing or updates not yet applied...")
	writelog("Missing " & searchResult.Updates.Count & " update(s).") 
End If 
on error resume next

For i = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(i)
	'if update.MsrcSeverity = "Important" then wscript.echo "This item (" & update.Title & ") is " & update.MsrcSeverity 
	strSearchResultUpdates = strSearchResultUpdates & update.Title & "<br>"
    Set objCategories = searchResult.Updates.Item(i).Categories
    writelog("Missing: " & searchResult.Updates.Item(i))
Next


if err.number <> 0 then
    writelog("An error has occured while instantiating search results.  Error " & err.number & " - " & err.description _
        & ".  Check the " & wshShell.ExpandEnvironmentStrings("%windir%") & "\windowsupdate.log file for further information.")

    blnFatal = false
End IF

    if silent = 0 then objdiv.innerhtml = strMsg & strsearchResultUpdates
    
If searchResult.Updates.Count = 0 Then
	
  strMsg = fformat & "There are no further updates needed for your PC at this time."
	
	if silent = 0 then objdiv.innerhtml = strMsg & "<br><br><a href='file:///" & logfile & "'>View log file</a>"

  writelog(replace(strMsg,fformat,""))
  writelog("Events saved to '" & logfile & "'")
  
  Call EndOfScript
  wscript.quit
End If


If intdebug = 1 then WScript.Echo vbCRLF & "Creating collection of updates to download:"
If strAction <> "detect" Then writelog("Creating a catalog of needed updates") 

writelog("********** Cataloging updates **********")

Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")

For I = 0 to searchResult.Updates.Count-1 
	Set update = searchResult.Updates.Item(I) 
	if update.MsrcSeverity <> "" then MsrcSeverity = "(" & update.MsrcSeverity & ") "
	strUpdates = strUpdates & MsrcSeverity & "- " & update.Title & "<br>"
 	writelog("Cataloged: " & MsrcSeverity & update.Title) 
	If Not update.EulaAccepted Then update.AcceptEula 
	updatesToDownload.Add(update) 
Next 

If silent = 0 then objdiv.innerhtml = ""
strMsg = fformat & "This PC requires updates from the configured Update Server" _
 & " (" & regWSUSServer & ").  "
If strAction <> "detect" Then strmsg = strmsg & "<br><br> Downloading needed updates.  Please stand by..."

if silent = 0 then objdiv.innerhtml = strMsg
writelog(replace(replace(strMsg,fformat,""),"<br>",""))

If strAction = "detect" Then 
	
Else
	
	Set downloader = updateSession.CreateUpdateDownloader() 
	on error resume next
	downloader.Updates = updatesToDownload
	writelog("********** Downloading updates **********")

	downloader.Download()

	if err.number <> 0 then
		writelog("Error " & err.number & " has occured.  Error description: " & err.description)
	End if

	strUpdates = ""
	strMsg = ""
	if silent = 0 then objdiv.innerhtml = ""
	strMsg = fformat & "List of downloaded updates: <br><br>"
	if silent = 0 then objdiv.innerhtml = strMsg
	
	For I = 0 To searchResult.Updates.Count-1
	    Set update = searchResult.Updates.Item(I)
	    If update.IsDownloaded Then
	       strDownloadedUpdates = strDownloadedUpdates & update.Title & "<br>"
	    End If
		On Error GoTo 0
		'writelog(searchResult.Updates.Item(i))
	    writelog("Downloaded: " & update.Title)
	    if silent = 0 then objdiv.innerhtml = strMsg & strDownloadedUpdates
	Next
	
	Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")
	
	strUpdates = ""
	strMsg = ""
	if silent = 0 then objdiv.innerhtml = ""
	strMsg = fformat & "Creating collection of updates needed to install:<br><br>" 

	If silent = 0 then objdiv.innerhtml = strMsg
	writelog("********** Adding updates to collection **********")

	For I = 0 To searchResult.Updates.Count-1
	    set update = searchResult.Updates.Item(I)
	    If update.IsDownloaded = true Then
	       strUpdates = strUpdates & update.Title & "<br>"
	       updatesToInstall.Add(update)
	    End If
	       writelog("Adding to collection: " & update.Title)
	       if silent = 0 then objdiv.innerhtml = strMsg & strUpdates	
	Next
End If


If lcase(strAction) = "prompt" Then 
  strMsg = "The Windows Update Agent has detected that this computer is missing updates from the " _
   & " configured server (" & regWSUSServer & ")." & vbcrlf & vbcrlf & "Would you like to install updates now?"
  strResult = MsgBox(strMsg,36,"Install now?")
  strUpdates = ""
  writelog(strMsg & " [Response: " & strResult & "]")
  strMsg = ""
  If silent = 0 then objdiv.innerhtml = ""

ElseIf strAction = "detect" Then
  strMsg = fformat & "Windows Update Agent has finished detecting needed updates." 
  writelog(replace(strMsg,fformat,""))
  
  if silent = 0 then objdiv.innerhtml = strMsg & "<br><br>"
  
  Call EndOfScript
  wscript.quit
ElseIf strAction = "install" Then
  strResult = 6
End If 

strUpdates = ""
if silent = 0 then objdiv.innerhtml = ""

If strResult = 7 Then
  strMsg = strMsg & "<br>User cancelled installation.  This window can be closed."
  writelog(replace(strMsg,"<br>",""))
  
  if silent = 0 then objdiv.innerhtml = strMsg

	WScript.Quit
ElseIf strResult = 6 Then
  strMsg = ""
  Set installer = updateSession.CreateUpdateInstaller()
  installer.AllowSourcePrompts = False 
  on error resume next 

  installer.ForceQuiet = True 
  
  strMsg = fformat & "Installing updates... <br><br>"
  writelog(replace(replace(strMsg,fformat,""),"<br>",""))
  
' If silent = 0 Then objdiv.innerhtml = strMsg & "<br>&bull; " & update.title

 If err.number <> 0 Then
	writelog("Error " & err.number & " has occured.  Error description: " & err.description)
 End if
  
	installer.Updates = updatesToInstall
	
	writelog("********** Installing updates **********")
	
	blnInstall = true
	
	on error resume next	
	Set installationResult = installer.Install()
	
	If Silent = 0 then objdiv.innerhtml = strMsg
	
	writelog(replace(replace(strMsg,fformat,""),"<br>",""))
 	
	If err.number <> 0 then 
	    'strMsg = "Error installing updates... Actual error was " & err.number & " - " & err.description & "."
	    'writelog(strmsg)
    	if silent = 0 then objdiv.innerhtml = strMsg
	End If

	'Output results of install
	strMsg = fformat & "Installation Result: " & installationResult.ResultCode & "<br><br>" _
	 & "Reboot Required: " & installationResult.RebootRequired & "<br><br>" _
	 & "Listing of updates and individual installation results: <br>"

	 For i = 0 to updatesToInstall.Count - 1
		If installationResult.GetUpdateResult(i).ResultCode = 2 Then 
			strResult = "Installed"
		ElseIf installationResult.GetUpdateResult(i).ResultCode = 1 Then 
			strResult = "In progress"
		ElseIf installationResult.GetUpdateResult(i).ResultCode = 3 Then 
			strResult = "Error"
		ElseIf installationResult.GetUpdateResult(i).ResultCode = 4 Then 
			strResult = "Failed"
		ElseIf installationResult.GetUpdateResult(i).ResultCode = 5 Then 
			strResult = "Aborted"			
		End If
		writelog(strResult & ": " & updatesToInstall.Item(i).Title)
		strUpdates = strUpdates & strResult & ": " & updatesToInstall.Item(i).Title & "<br>"
	Next
	if silent = 0 then objdiv.innerhtml = strMsg & strUpdates
End If		

Call EndOfScript
wscript.quit

'*******************************************************************************
'Function Writelog 
'*******************************************************************************
Function WriteLog(strMsg) 
l.writeline "[" & time & "] - " & strMsg
' Output to screen if cscript.exe 
If blnCScript Then WScript.Echo "[" & time & "] " & strMsg 
End Function 

'*******************************************************************************
'Function IE Status
'*******************************************************************************
Function IEStatus

'added by Rob - IE status indicator code
If blnProgressMode Then
	If blnDebugMode Then
		dbgTitle = "Windows Update Script " & strScriptVer
	Else
		dbgTitle = "Windows Update Script " & strScriptVer
	End If	
	dbgToolBar = False
	dbgStatusBar = False
	If blnDebugMode Then
		dbgResizable = True
	Else
		dbgResizable = False
	End If
	dbgWidth = 500
	dbgHeight = 320

 on error resume next

    'get video resolution via WMI
    Set vids = GetObject("WinMgmts:").instancesof("Win32_VideoController")
        for each v in vids
                HorScreen = v.CurrentHorizontalResolution
                VerScreen = v.CurrentVerticalResolution
        next
    If err.number <> 0 then
        dbgLeft = 100
        dbgTop = 200
        err.clear
    Else
        HorScreen = 800
        VerScreen = 600
        dbgLeft = (HorScreen * .5) - (dbgWidth/2)
        dbgTop = (VerScreen * .5) - (dbgHeight/2)
    End if


	dbgLeft = (HorScreen * .5) - (dbgWidth/2)
	dbgTop = (VerScreen * .5) - (dbgHeight/2)
	dbgVisible = True
	dlgBarWidth = 380
	dlgBarHeight = 23 
	dlgBarTop = 5
	dlgBarLeft = 82
	dlgProgBarWidth = 0
	dlgProgBarHeight = 18 
	dlgProgBarTop = 82
	dlgProgBarLeft = 50
	dlgBar = "left: " & dlgBarLeft & "; top: " & dlgBarTop & "; width: " & dlgBarWidth _
	 & "; height: " & dlgBarHeight & ";"
	dlgProgBar = "left: " & dlgProgBarLeft & "; top: " & dlgProgBarTop & "; width: " _
	 & dlgProgBarWidth & "; height: " & dlgProgBarHeight & ";"
	wdBar = 1 * dlgBarWidth
End If

If blnProgressMode Then
  ' in case people has used the search bar in IE, turn it off 
  ' Thank you Torgeir!
  Set IEtmp = CreateObject("InternetExplorer.Application")
  IEtmp.ShowBrowserBar "{30D02401-6A81-11D0-8274-00C04FD5AE38}", False 
  IEtmp.Quit 
  
  Set IEtmp = Nothing 
  WScript.Sleep 1000
	
  Set IE = CreateObject("InternetExplorer.Application")
	'strScriptVer = "version would go here"

	strTempFile = WshSysEnv("TEMP") & "\progress.htm"
	ws.CreateTextFile (strTempFile)
        Set f1 = ws.GetFile(strTempFile)
        Set ts = f1.OpenAsTextStream(2, True)
        ts.WriteLine("<!-- saved from url=(0014)about:internet -->")
        ts.WriteLine("<html><head><title>" & dbgTitle & " " & strScriptVer & " </title>")
        ts.WriteLine("<style>.errortext {color:red}")
     	ts.WriteLine(".hightext {color:blue}</style>")
	ts.WriteLine("</head>")
	ts.WriteLine(strHDRCode & " <br><strong><font size='2' color='" & fcolor & "' face='" & fstyle & "'>" _
	 	& "&nbsp Running Windows Update Client...<br>" _
	 	& "&nbsp &nbsp<br>")
	ts.WriteLine("<center><table width='100%' bgcolor='" & bgcolor1 & "'><tr><td>")
	If blnDebugMode Then
		ts.WriteLine("<body bgcolor ='" & stsBGColor & "' scroll='yes' topmargin='0' leftmargin='0'"_
		& " style='font-family: " & fstyle & "; font-size: 0.6em color: #000000;"_
		& " font-weight: bold; text-align: left'><center><font face=" & fstyle & ">"_
		& " <font size='0.8em'> <hr color='blue'>")
	Else
		ts.WriteLine("<body bgcolor = '" & stsBGColor & "' scroll='no' topmargin='0' leftmargin='0' "_
		& " style='font-family: " & fstyle & "; font-size: 0.6em color: #000000;"_
		& " font-weight: bold; text-align: left'><center><font face=" & fstyle & ">"_
		& " <font size='0.8em'> <hr color='blue'>")
	End If
	ts.WriteLine("<div id='ProgObject' align='left'align='left' style='width: 450px;height: 140px;overflow:scroll'></div><hr color='blue'>")			
	If blnDebugMode Then
		ts.WriteLine("<div id='ProgDebug' align='left'></div>")
	End If

	ts.WriteLine("<script LANGUAGE='JavaScript1.2'>")
	ts.WriteLine("<!-- Begin")
	ts.WriteLine("function initArray() {")
	ts.WriteLine("this.length = initArray.arguments.length;")
	ts.WriteLine("for (var i = 0; i < this.length; i++) {")
	ts.WriteLine("this[i] = initArray.arguments[i];")
	ts.WriteLine("   }")
	ts.WriteLine("}")
	ts.WriteLine("var ctext = ' ';")
	ts.WriteLine("var speed = 1000;")
	ts.WriteLine("var x = 0;")
	ts.WriteLine("var color = new initArray(")
	ts.WriteLine("'red',")
	ts.WriteLine("'blue'")
	ts.WriteLine(");")
	ts.WriteLine("document.write('<div id=" & Chr(34) & "ProgFlash" & Chr(34) & ">"_
	 & "<center>'+ctext+'</center></div>');")
	ts.WriteLine("function chcolor(){")
	ts.WriteLine("document.all.ProgFlash.style.color = color[x];")
	ts.WriteLine("(x < color.length-1) ? x++ : x = 0;")
	ts.WriteLine("}")
	ts.WriteLine("setInterval('chcolor()',1000);")
	ts.WriteLine("// End -->")
	ts.WriteLine("</script>")
	ts.WriteLine("<div id='ProgBarId' align='left'></div>")
	ts.WriteLine("</font></center>")
	ts.WriteLine("</tr></td>")
	ts.WriteLine("</table></center>")
	ts.WriteLine("</body></html>")
	ts.Close
	fctSetupIE(strTempFile)
	Set objDIV = IE.Document.All("ProgObject")
	If blnDebugMode Then
		Set objDBG = IE.Document.All("ProgDebug")
	End If
	Set objFlash = IE.Document.All("ProgFlash")
	Set objPBar = IE.Document.All("ProgBarId")
	Set objBar = IE.Document
End If
If silent = 1 Then
'remarked by Rob Set logwindow = ie.document.all.text1
End If
End Function
'*******************************************************************************'*	Name:	fctSetupIE
'*	Function:	Setup an IE windows of 540 x 200 to display 
'* 	progress information.
'*******************************************************************************
Sub fctSetupIE(File2Load)
	IE.Navigate File2Load
	IE.ToolBar = dbgToolBar
	IE.StatusBar = dbgStatusBar
	IE.Resizable = dbgResizable
	Do
	Loop While IE.Busy
	IE.Width = dbgWidth
	IE.Height = dbgHeight
	IE.Left = dbgLeft
	IE.Top = dbgTop
	IE.Visible = dbgVisible
	wshshell.AppActivate("Microsoft Internet Explorer")
End Sub

Sub GetAUSchedule()
Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
Set objSettings = objAutoUpdate.Settings

Select Case objSettings.ScheduledInstallationDay
    Case 0
        strDay = "every day"
    Case 1
        strDay = "sunday"
    Case 2
        strDay = "monday"
    Case 3
        strDay = "tuesday"
    Case 4
        strDay = "wednesday"
    Case 5
        strDay = "thursday"
    Case 6
        strDay = "friday"
    Case 7
        strDay = "saturday"
    Case Else
        strDay = "The scheduled installation day is could not be determined."
End Select

If objSettings.ScheduledInstallationTime = 0 Then
    strScheduledTime = "12:00 AM"
ElseIf objSettings.ScheduledInstallationTime = 12 Then
    strScheduledTime = "12:00 PM"
Else
    If objSettings.ScheduledInstallationTime > 12 Then
        intScheduledTime = objSettings.ScheduledInstallationTime - 12
        strScheduledTime = intScheduledTime & ":00 PM"
    Else
        strScheduledTime = objSettings.ScheduledInstallationTime & ":00 AM"
    End If
    'strTime = "Scheduled installation time: " & strScheduledTime
End If

writelog("Windows update agent is scheduled to run on " & strDay & " at " & strScheduledTime)
End Sub

'*******************************************************************************
'Function URLGet - Check to see if web page is active
' 
'Thanks to http://www.sebsworld.net/information/?page=VBScript-URL
'*******************************************************************************
Function URLGet(URL)
		on error goto 0
		'  	Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")
		'  	Set Http = CreateObject("WinHttp.WinHttpRequest.5")
		'	Set Http = CreateObject("WinHttp.WinHttpRequest")
		'	Set Http = CreateObject("Msxml2.ServerXMLHTTP.6.0")
		'	Set Http = CreateObject("MSXML2.ServerXMLHTTP")
			Set Http = CreateObject("MSXML2.XMLHTTP")
		'	Set Http = CreateObject("Microsoft.XMLHTTP")
		
		'msgbox URL
		Http.Open "HEAD",URL,False
		'const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056 
		'Http.setOption 2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
		Http.Send
		
		'pagestatus = Http.status
		If Http.status <> "403" and Http.status <> "200" Then
			URLGet = "(ERROR):" & Http.status
		Else
			'URLGet = Http.ResponseBody
			'URLGet = Http.responseText
			URLGet = "OK " & Http.status
		End if
		
		
End Function

Function URLPost(URL,FormData,Boundary)
  Set Http = CreateObject("Microsoft.XMLHTTP")
  Http.Open "POST",URL,True
'  Http.setRequestHeader "Content-Type","multipart/form-data; boundary="& Boundary
  Http.send FormData
  for n = 1 to 9
    If Http.readyState = 4 then exit for
    ' Http.waitForResponse 1
    b = shell.popup("Getting page",1,"Message")
  next
  If Http.readyState <> 4 then
    URLPost = "Failed"
  else
    URLPost = Http.responseText
  end if
End Function

'*******************************************************************************
'Function SendMail - email the warning file
'*******************************************************************************
Function SendMail(strFrom,strTo,strSubject,strMessage)
Dim iMsg, iConf, Flds

writelog("Calling sendmail routine")
writelog("To: " & strMailto)
writelog("From: " & strMailFrom)
writelog("Subject: " & strSubject)
writelog("SMTP Server: " & strSMTPServer)

'If silent = 0 Then objdiv.innerhtml = "<font face=" & strFontStyle & " color=" & strFontColor2& ">" _
' & "sending mail to " & strMailTo & "...</font><br>"

'//  Create the CDO connections.
Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields

If lcase(strAuthType) <> "cdoanonymous" Then
  'Type of authentication, NONE, Basic (Base64 encoded), NTLM
  iMsg.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = strAuthType

  'Your UserID on the SMTP server
  iMsg.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = strAuthID

  'Your password on the SMTP server
  iMsg.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strAuthPassword

End if

'// SMTP server configuration.
With Flds
	.Item(cdoSendUsingMethod) = cdoSendUsingPort
	.Item(cdoSMTPServer) = strSMTPServer
	.Item(cdoSMTPServerPort) = iSMTPServerPort
	.Item(cdoSMTPconnectiontimeout) = 60
	.Update
End With
'l.close

Dim r
Set r = ws.OpenTextFile (logfile, ForReading, False, TristateUseDefault)
strMessage = "<font face='" & strFontStyle & "' size='2'>" & r.readall & "</font>"

'//  Set the message properties.
With iMsg
    Set .Configuration = iConf
        .To       = strMailTo
        .From     = strMailFrom
        .Subject  = strSubject
        '.TextBody = strMessage
End With

'iMsg.AddAttachment wsuslog
iMsg.HTMLBody = replace(strMessage,vbnewline,"<br>")
'//  Send the message.
on error resume next

iMsg.Send ' send the message.
Set iMsg = nothing

If CStr(err.number) <> 0 Then
	strMsg = "Problem sending mail to " & strSMTPServer & "." _
   & "Error [" & err.number & "]: " & err.description & "<br>"
  
  Call ErrorHandler("Sendmail function",replace(strMsg,"<br>",""),"false")
  'writelog(strMsg)
  strStatus = strMsg
  If silent = 0 Then objdiv.innerhtml = strStatus
Else
  strStatus = "Connected successfully to email server " & strSMTPServer
  writelog(strStatus)
	strStatus = strStatus & "<br><br><font face=" & strFontStyle & " color=" & strFontColor2& ">" _
 & "sent email to " & strMailTo & "...</font><br><BR>" _
	 & "Script complete.<br><br><a href='file:///" & logfile & "'>View log file</a>"
	If silent = 0 Then objdiv.innerhtml = strStatus
End If

'cause WU agent to detect
autoUpdateClient.detectnow()
blnEmail = 0

End Function
'*******************************************************************************'Function RestartAction
'Sub to perform a restart action against the computer
'*******************************************************************************
Function RestartAction
  If silent = 0 Then objdiv.innerhtml = strStatus & "<br> Now performing post-execute action (" & restarttext & ")."
  wscript.sleep 4000
  writelog("Processing PostExecuteAction")
	'On Error GoTo 0
	Dim OpSysSet, OpSys
	'writelog("Computer: " & strComputer & vbcrlf & "Post-execution action: " & strRestart)

  	'On Error Resume Next
  	
	'Call WMI query to collect parameters for reboot action
	Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//" & strComputer & "/root/cimv2").ExecQuery("select * from Win32_OperatingSystem"_
	 & " where Primary=true") 
	 
	If CStr(err.number) <> 0 Then 
	  strMsg = "There was an error while attempting to connect to " & strComputer & "." & vbcrlf & vbcrlf _
		 & "The actual error was: " & err.description
		writelog(strMsg)
		blnFatal = true
    	Call ErrorHandler("WMI Connect",strMsg,blnFatal)
	End If

  	Const EWX_LOGOFF = 0 
  	Const EWX_SHUTDOWN = 1 
  	Const EWX_REBOOT = 2 
  	Const EWX_FORCE = 4 
  	Const EWX_POWEROFF = 8 
	
	'set PC to reboot
	If strRestart = 1 Then

		For each OpSys in OpSysSet 
			opSys.win32shutdown EWX_REBOOT + EWX_FORCE
		Next 

	'set PC to shutdown
	ElseIf strRestart = 2 Then
				
		For each OpSys in OpSysSet 
			opSys.win32shutdown EWX_POWEROFF + EWX_FORCE
		Next 
  
  'Do nothing...
  ElseIf strRestart = "0" Then
    				
End If


End Function

'*******************************************************************************
'Sub ErrorHandler
'Sub to help display/log any errors that occur
'*******************************************************************************
Sub ErrorHandler(strSource,strMsg,blnFatal)
    'Set theError = RemoteScript.Error

		If silent = 0 then wscript.echo "Source: " & strSource & " - " & strMsg
		writelog(strMsg)
		If blnFatal = true then wscript.quit
    err.clear
End Sub

'*******************************************************************************'Function EndOfScript
'Function to close out the script
'*******************************************************************************
Function EndOfScript

  If blnInstall = true then Call CheckPendingStatus("end")
  on error goto 0
  writelog("Windows Update VB Script finished")
  l.writeline "---------------------------------------------------------------------------"
  If blnCallRestart = true then writelog("Post-execute action will be called.  " _
   & " Action is set to: " & restarttext & ".")
     

  If blnEmail = 1 Then
     If searchresult.updates.count = 0 and not blnRebootRequired and StrEmailifAllOK = 0 then
        writelog ("No updates required, no pending reboot, therefore not sending email")
     else
        if strFullDNSName = 1 then
           strDomainName = wshShell.ExpandEnvironmentStrings("%USERDNSDOMAIN%")
	  			 strOutputComputerName = strComputer & "." & StrDomainName
        else
           strOutputComputerName = strComputer         
        end if
        if emailifallok = 0 or emailifallok = 1 then
          if instr(strSMTPServer,"x") then
          else
           Call SendMail(strFrom,strTo,strSubject & " " & strOutputComputerName,strMessage)
          end if
        end if
     end if
  Else
      'l.close

  End If

  strMsg = "The script has been configured to " & restarttext _
   		& ".  The update script has detected that this " _
   		& "computer has a reboot pending from a previous update session." & vbcrlf & vbcrlf _
   		& "Would you like to perform this action now?"
  
  If silent = 0 and blnPrompt = true Then 
  	strResult = MsgBox(strMsg,36,"Perform restart/shutdown action?")
  ElseIf blnPrompt = false Then
  	strResult = 6
  End If
     
  If blnCallRestart = true Then 
  	If strResult = 6 Then call RestartAction
  Else
    on error resume next
  	If silent = 0 Then objdiv.innerhtml = strStatus & "<br>This computer has no pending reboots"
  End If
  
  If intSleep > 0 Then
      ' So the user have a chance to see the last output before closing IE
      WScript.Sleep intSleep
      ' Just in case the IE window is already closed by the user
      On Error Resume Next
      ' Close the IE window
      IE.Quit
      On Error Goto 0
  End If
  'l.close
  wscript.quit
  
  Exit Function
   
End Function

'*******************************************************************************
'Function CheckPendingStatus
'Function to restart the computer if there is a reboot pending...
'*******************************************************************************
Function CheckPendingStatus(beforeorafter)
  Set ComputerStatus = CreateObject("Microsoft.Update.SystemInfo")
  Select case beforeorafter
    Case "beginning"
      strCheck = "Pre-check"
    Case "end"
      strCheck = "Post-check"
    Case Else
  End Select
  
  blnRebootRequired = ComputerStatus.RebootRequired

  If ComputerStatus.RebootRequired or strForceAction = 1 Then
     If beforeorafter = "beginning" Then 
        If ComputerStatus.RebootRequired Then strMsg = "This computer has a pending reboot (" & strCheck & ").  Switching to 'detect' mode."
        If strAction = "prompt" Then blnPrompt = true
        strAction = "detect"
        blnCallRestart = true  
     Else
        If ComputerStatus.RebootRequired Then strMsg = "This computer has a pending reboot (" & strCheck & ").  Setting PC to perform post-script " _
          & "execution..."
        blnCallRestart = true        
     End If
  Else
        If not ComputerStatus.RebootRequired Then strMsg = "This computer does not have any pending reboots (" & strCheck & ")."
  End If
  
     If strMsg <> "" Then writelog(strMsg)
     If silent = 0 and strMsg <> "" then objdiv.innerhtml = strMsg
     'wscript.sleep 4000
           
End Function
'*******************************************************************************
