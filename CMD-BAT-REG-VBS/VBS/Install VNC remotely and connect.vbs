strScriptVer = "3.0 2/26/10 Rob.Dunn"

'Set to false when you're ready to run the file in production (or have your 
' folder structures set up!
bSetup = true

If bSetup = true then
  msgbox "Setup mode enabled.  Disable setup mode by changing 'bSetup' to 'true' (without the apostrophes) in the core script." & vbcrlf & vbcrlf & "Once you've verified your folders have been created, disable setup mode.",48,"Setup mode notice"
End If

'Get the root path that the script resides in
strScriptPath = replace(wscript.scriptfullname,wscript.scriptname,"")

Const ForReading = 1
Dim strRemoteSystemDrive, sVNCType, sLocalDriveToMap, sVNCExe, sRegRoot
Dim sRemoveSwitch, sInstallSwitch, sVNCServiceName, strFilestoCopy
Dim fso, wshshell, wshsysenv, ws, e, objWmiService, objLocator, oReg
Dim blnFatal, strRemoteOSVersion, blnUpdate, strFoldertoCopyFrom, strUpdatePassword
Dim strPassword, strComputer, sRemoteFolder, strUserCredentials, strPasswordCredentials
Dim bExempt
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

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objArgs = WScript.Arguments
Set WshSysEnv = WshShell.Environment("PROCESS")
Set ws = wscript.CreateObject("Scripting.FileSystemObject")
Set WshNetwork = WScript.CreateObject("WScript.Network")
Set objLocator = CreateObject("WbemScripting.SWbemLocator")

Const HKEY_LOCAL_MACHINE = &H80000002
Const ForWriting = 2

'above variables for progress indicators

'******************************************************************************
'Set user variables here
'******************************************************************************
'force remote computer to prompt with 'accept/reject' for remote connection
'  0 = do not prompt
'  1 = prompt
iQueryConnect=0

'Default port to connect to via vncviewer
iPortNumber = "5900"

'Default port to connect to via HTTP
iHTTPPortNumber = "5800"

'Allow VNC connection to the remote computer if no one is logged in without
' prompting.
' 0 = allow connection without prompt with no user logged in remotely
' 1 = do not allow connection without prompt with no user logged in remotely
iQueryOnlyIfLoggedOn=0
'
'name of workstation registry file where password is contained.
sWorkstationRegistry = "vnchooks_settings.reg"
'
'name of server registry file where password is contained.
sServerRegistry = "vnchooks_settings_server.reg"

'Use SYSTEMDRIVE:\folder path if you wish for the script to determine the 
' same volume which contains Windows as the drive to install to.  Otherwise
' you can specify your own drive letter.
sRemoteFolder = "SYSTEMDRIVE:\program files"

'Where are the files located that we wish to copy to the remote system?  
' Note this does not include specific VNC flavors, but the general registry
' keys, etc.  Each VNC flavor will be stored in their folders underneath 
' this folder.
strFoldertoCopyFrom = strScriptPath & "vnc"

'Close the status window when we are done with the connection?
' true|false
bCloseStatusWindow = true

'Set the startup mode of the service to...?
'recommend 'manual' or 'automatic', anything else will cause problems.
sStartupMode = "automatic" 

'Stop service after VNC Viewer terminates?  Command-line switch takes
' precedence.
' True|False
bStop = false

'Remove service when complete?  Command-line switch takes precedence
'Setting this to 'true' will automatically change bStop to 'true'.
' True|False
blnUnregister = true

'Use the VNC Flavor Viewer?  If set to false, it will use whatever vncviewer
' is located in the root of the script folder.
bUseVNCFlavorViewer = true

'Name of VNC Viewer EXE
strVNCViewerEXE = "vncviewer.exe"

'###variables for the local status window

'Background image for status window
strBackground = "background.jpg"
bgcolor1 = "black"	'first alternating table color
bgcolor2 = "aliceblue"	'second alternating table color
bgcolor3 = "beige"	'warning table color
fstyle = "arial"	'font face for table
HDRfcolor = "slategray"	'header font color
fcolor = "#DDDFF8"		'default font color for table
stsBGcolor = "black" ' set to transparent once you get a background you like

'******************************************************************************
sVar = split(sRemoteFolder,":\")
sRemoteFolder1 = sVar(1) & "\"

If bSetup = true then 
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  MakeSureDirectoryTreeExists(strScriptPath & "vnc\realvnc")
  Set objFile = objFSO.OpenTextFile(strScriptPath & "vnc\realvnc\neededfiles.txt", ForWriting, TRUE)
  objFile.writeline("Copy the following files here: vncviewer.exe;winvnc4.exe;wm_hooks.dll ")
  objFile.close
  Set objFile = objFSO.OpenTextFile(strScriptPath & "vnc\realvnc\realvnc.url", ForWriting, TRUE)
  objFile.writeline("[InternetShortcut]" & vbcrlf & "URL=http://www.realvnc.com")
  objFile.close

  MakeSureDirectoryTreeExists(strScriptPath & "vnc\tightvnc")
  Set objFile = objFSO.OpenTextFile(strScriptPath & "vnc\tightvnc\neededfiles.txt", ForWriting, True)
  objFile.writeline("Copy the following files here: vncviewer.exe;winvnc.exe;vnchooks.dll ")
  objFile.close
  Set objFile = objFSO.OpenTextFile(strScriptPath & "vnc\tightvnc\tightvnc.url", ForWriting, TRUE)
  objFile.writeline("[InternetShortcut]" & vbcrlf & "URL=http://www.tightvnc.com")
  objFile.close

  MakeSureDirectoryTreeExists(strScriptPath & "vnc\ultravncreg")
  Set objFile = objFSO.OpenTextFile(strScriptPath & "vnc\ultravncreg\neededfiles.txt", ForWriting, True)
  objFile.writeline("Copy the following files here: vncviewer.exe;winvnc.exe;vnchooks.dll;ultravnc.ini")
  objFile.close
  Set objFile = objFSO.OpenTextFile(strScriptPath & "vnc\ultravncreg\ultravnc.url", ForWriting, TRUE)
  objFile.writeline("[InternetShortcut]" & vbcrlf & "URL=http://www.uvnc.com")
  objFile.close

  'MakeSureDirectoryTreeExists(strScriptPath & "vnc\ultravncini")
  
  Set objFile = objFSO.OpenTextFile(strScriptPath & "vnc\vnchooks_settings.reg", ForWriting, True)
  objFile.writeline(chr(34) & "Password" & chr(34) & "=hex:bf,27,fb,81,f5,1e,30,21")
  objFile.close
  Set objFile = objFSO.OpenTextFile(strScriptPath & "vnc\vnchooks_settings_server.reg", ForWriting, True)
  objFile.writeline(chr(34) & "Password" & chr(34) & "=hex:bf,27,fb,81,f5,1e,30,21")
  objFile.close
  Set objFile = objFSO.OpenTextFile(strScriptPath & "vnc\query_exempt.txt", ForWriting, True)
  objFile.writeline("exclude all workstations")
  objFile.close
  
  Set objFile = objFSO.OpenTextFile(strScriptPath & "vnc\readme.txt", ForWriting, True)
  With objFile
    .writeline("Download the different flavors of VNC: RealVNC, TightVNC and UltraVNC" & vbcrlf)
    .writeline("Look in each subfolder to find out which files you need to copy there." & vbcrlf)
    .writeline("Copy your preferred vncviewer.exe here (" & strScriptPath & ")" & vbcrlf)
    .writeline("Find a nice 640x400 dark image, and copy it here and name it 'background.jpg'." & vbcrlf & "This will make your status window look snazzy.")
    .writeline(vbcrlf & "To disable setup mode, edit the vbscript and change the 'bSetup' variable to 'false'")
    .writeline(vbcrlf & "Default connect password is 'remote'.  To export your own password, install")
    .writeline(vbcrlf & "VNC on your local computer, set a connection password, open the registry and ")
    .writeline(vbcrlf & "copy and paste the registry password value into the reg keys in the 'vnc' folder")
    .writeline(vbcrlf & "I recommend using RealVNC for this.  Browse to HKLM\SOFTWARE\RealVNC\WinVNC4\Password and")
    .writeline(vbcrlf & "export the registry key.  Copy only the password value to the reg key in the vnc folder.")
    .writeline(vbcrlf & "Create a different password for your server connections and repeat the process,")
    .writeline(vbcrlf & "only this time paste your value into the server regkey.")
    .writeline(vbcrlf & "Need help?  Post a question to the Spiceworks forums...- Rob.Dunn")
    .close
  End With
  
  Set objFile = objFSO.OpenTextFile(strScriptPath & "vnc\SpiceworksSupport.url", ForWriting, True)
  objFile.writeline("[InternetShortcut]" & vbcrlf & "URL=http://community.spiceworks.com/referral/a8f53140297597ad39f1332c9341c694")  
  objFile.close
  wscript.quit
End If


Function DefineVNC(sVNCType)
  if blnUnregister = true then bStop = true
  
  Select case lcase(sVNCType)
    Case "realvnc"
      sRegRoot = "SOFTWARE\RealVNC\WinVNC4\"
      sRemoveSwitch = "-unregister"
      sInstallSwitch = "-register"
      sVNCServiceName = "winvnc4"
      sVNCExe = "winvnc4.exe"
      strFilestoCopy = sVNCExe & ";wm_hooks.dll"
    Case "ultravncreg"
      sRegRoot = "SOFTWARE\ORL\WinVNC3\Default\"
      sRemoveSwitch = "-uninstall"
      sInstallSwitch = "-install"
      sVNCServiceName = "uvnc_service"
      sVNCExe = "winvnc.exe"
      strFilestoCopy = sVNCExe & ";vnchooks.dll;ultravnc.ini"
    Case "tightvnc"
      sRegRoot = "SOFTWARE\ORL\WinVNC3\Default\"
      sRemoveSwitch = "-remove"
      sInstallSwitch = "-install"
      sVNCServiceName = "winvnc"
      sVNCExe = "winvnc.exe"
      strFilestoCopy = sVNCExe & ";vnchooks.dll"
    Case "ultravncini"
      'NOT WORKING YET
      sRegRoot = ""
      sRemoveSwitch = "-uninstall"
      sInstallSwitch = "-install"
      sVNCServiceName = "winvnc"
      sVNCExe = "winvnc.exe"
      strFilestoCopy = sVNCExe & ";vnchooks.dll;ultravnc.ini"
    End Select
End Function

'Get command-line arguments
If objargs.count < 1 Then
	Call fctInput
Else
 For I = 0 to objArgs.Count - 1
	'msgbox objargs(i)
  
   If InStr(1,LCase(objargs(I)),"computer:") Then
   	arrComputer = split(lcase(objargs(I)),"computer:")
   	strComputer = arrComputer(1)
	 ElseIf InStr(1,LCase(objargs(I)),"update:") Then
   	arrUpdate = split(lcase(objargs(I)),"update:")
   	blnUpdate = arrUpdate(1)
   	If lcase(blnUpdate) <> "true" and lcase(blnUpdate) <> "false" then 
   		blnUpdate = false
   	End If
	 ElseIf InStr(1,LCase(objargs(I)),"user:") Then
   	arrUser = split(lcase(objargs(I)),"user:")
   	strUserCredentials = arrUser(1)
	 ElseIf InStr(1,LCase(objargs(I)),"password:") Then
   	arrPassword = split(objargs(I),"password:")
   	strPasswordCredentials = arrPassword(1)
   ElseIf InStr(1,LCase(objargs(I)),"servicestop:") Then
   	arrStop = split(objargs(I),"servicestop:")
   	bStop = arrStop(1)
   ElseIf Instr(1,LCase(objargs(I)),"tempdrive:") Then
    arrDrive = split(lcase(objargs(I)),"tempdrive:")
    sLocalDriveToMap = arrDrive(1) & ":"
   ElseIf Instr(1,LCase(objargs(I)),"type:") Then
    arrType = split(lcase(objargs(I)),"type:")
    sVNCType = arrType(1)
   ElseIf Instr(1,Lcase(objargs(I)),"setup:") Then
    arrSetup = split(lcase(objargs(I)),"setup:")
    bSetup = arrSetup(1)
	 ElseIf InStr(1,LCase(objargs(I)),"unregister:") Then
   	arrUnregister = split(lcase(objargs(I)),"unregister:")
    on error resume next
   	blnUnregister = cBool(arrUnregister(1))
    If cstr(err.number) <> 0 then
      msgbox "You have specified the 'unregister' command-line switch, but entered an incorrect value.  Correct values are: true/false/0/1." _
       & vbcrlf & "VNC script aborting."
      wscript.quit
    End if
   End If
 Next 
End If

strPrepareFont = "<font color='" & fcolor & "' size='0'>"

on error resume next
strClientIP = "<font size='0' face='verdana' color='" & fcolor & "'>Your IP address(es):<br>"
'Get the connecting client's IP
Set objLocalWMIService = GetObject ("winmgmts:\\" & "." & "\root\cimv2")
Set colAdapters = objLocalWMIService.ExecQuery ("Select IPAddress from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
For Each objAdapter in colAdapters
  If Not IsNull(objAdapter.IPAddress) Then
	   For i = LBound(objAdapter.IPAddress) To UBound(objAdapter.IPAddress)
	      If objAdapter.IPAddress(i) <> "0.0.0.0" then
          strClientIP = strClientIP & strBR & objAdapter.IPAddress(i)
	        strBR = "<br>"
        End If
    Next
  End If
Next

Sub ConnectToComputer(strComputer,strConnType)
  on error goto 0
     If strConnType = "wmi" then
      'msgbox strUserCredentials & " " & strPasswordCredentials
      Set objWMIService = objLocator.ConnectServer(strComputer,"root\cimv2",strUserCredentials,strPasswordCredentials)
     ElseIf strConnType = "reg" then
      
      Set objWMIService = objLocator.ConnectServer(strComputer,"root\default", strUserCredentials, strPasswordCredentials) 
      Set oReg = objWMIService.Get("StdRegProv") 
    End If
   
End Sub

'set some IE status indicator variables...
blnDebugMode = False
blnProcessEvents = True
blnSearchWildcard = False
blnProgressMode = True

Set fso = CreateObject("Scripting.FileSystemObject")

If bUseVNCFlavorViewer <> true then 
  'Set the VNC Viewer path in relation to the script path
  strVNCViewer = chr(34) & strFoldertoCopyFrom & strVNCViewerEXE & " " & chr(34)
Else
  
  strVNCViewer = chr(34) & strFoldertoCopyFrom & "\" & sVNCType & "\" & strVNCViewerEXE & " " & chr(34)
  'msgbox strVNCViewer
End If

blnUpdate = "false"

'Bring up inputbox, this is called only when no command-line arguments for computer name is given
Sub fctInput()
  on error resume next
  TempComputer = WshShell.RegRead("HKCU\Software\RDScripts\InstallVNC\LastComputer")
  TempVNCType = WshShell.RegRead("HKCU\Software\RDScripts\InstallVNC\VNCType")

    Select Case lcase(TempVNCType)
      Case "realvnc"
        TempVNCType = "1"
      Case "tightvnc"
        TempVNCType = "2"
      Case "ultravncreg"
        TempVNCType = "3"
      Case "ultravncini"
        TempVNCType = "4"
    End Select
  
	If strComputer = "" Then strComputer = InputBox("Type the name of the computer you wish to install VNC to and remote control.","Input computer name",TempComputer)
	If strComputer = "" Then wscript.quit
	If sVNCType = "" then sVNCType = Inputbox("Enter which VNC flavor you would like to run:" & vbcrlf & vbcrlf & "1 - RealVNC " & vbcrlf & "2 - TightVNC" & vbcrlf & "3 - UltraVNC (Registry)","Input VNC Type",TempVNCType)
	if sVNCType = "" or sVNCType > 3 then 
    wscript.quit
  Else

    Select Case sVNCType 
      Case "1"
        sVNCType = "realvnc"
      Case "2"
        sVNCType = "tightvnc"
      Case "3"
        sVNCType = "ultravncreg"
      Case "4"
        sVNCType = "ultravncini"
      Case Else
        wscript.quit
    End Select
  End If	
End Sub

Call DefineVNC(sVNCType)

'Open up the status window
Call IEStatus

If strUserCredentials <> "" then 
    strStatus = strPrepareFont & Time & " - Using alternate credentials: [<strong>" & strUserCredentials & "</strong>]<br>" & strStatus
    objdiv.innerhtml = strStatus
End If


Call GetOSVersion(StrComputer)

'msgbox sRemoteFolder & " " & strRemoteSystemDrive
sRemoteFolder = replace(sRemoteFolder,"SYSTEMDRIVE:",strRemoteSystemDrive) & "\" & sVNCType
sRemoteDrive = left(sRemoteFolder,3)
'msgbox sRemoteDrive
  
strStatus = strPrepareFont & Time & " - Starting VNC viewer installation and remote control utility...<br>" & strStatus
objdiv.innerhtml = strStatus

strFoldertoCopyTo = "\\" & strComputer & "\" & replace(sRemoteFolder,":\","$\")


'set the default reg key to be workstation in case it errors out below
strRegSettings = sWorkstationRegistry

On Error Resume Next
If instr(strRemoteOSVersion,"2003") or instr(lcase(strRemoteOSVersion),"server") or instr(lcase(strRemoteOSVersion),"powered") then
  strRegSettings = sServerRegistry
End If
err.clear

'If err.number <> 0 Then
'	strRegSettings = sWorkstationRegistry
'End If

	strStatus = strPrepareFont & Time & " - " & "Remote OS detected: " & strRemoteOSVersion & "<br>" & strstatus
	objdiv.innerhtml = strStatus

If CheckVNCService(strComputer) = "installed" Then 
  objdbg.innerhtml = strClientIP  
	On Error Resume Next
	
	Call SetRemoteOptions(strComputer)
	
	Call startVNCViewer(strComputer)
	
	strStatus = strPrepareFont & Time & " - " & "VNCViewer has terminated. <br>" & strstatus
	objdiv.innerhtml = strStatus
	
  If bStop = true then	Call ChangeVNCServiceState(strComputer,"stop")
	
	strStatus = strPrepareFont & Time & " - " & "Finished with the VNCViewer. <br>" & strstatus
	objdiv.innerhtml = strStatus
	
	If blnUnregister = true then

  	strStatus = strPrepareFont & Time & " - " & "Unregistering service from " & strComputer & ".<br>" & strstatus
    Call RunProcess(chr(34) & replace(replace(strFoldertoCopyTo,"$\",":\"),"\\" & strComputer & "\","") & "\" & sVNCExe & chr(34) & " " & sRemoveSwitch,strComputer)
	  objdiv.innerhtml = strStatus
	End If

	strStatus = strPrepareFont & Time & " - " & "You can close this window when ready.<br>" & strstatus
	objdiv.innerhtml = strStatus

	wscript.sleep 7000
	
  If bCloseStatusWindow = true then IE.quit
Else
	'On Error Resume Next
  objDBG.innerhtml = strClientIP


  strVar = strFoldertoCopyTo
  on error resume next

  Dim objDictionary
  Set objDictionary = CreateObject("Scripting.Dictionary")

  Call ConnectToComputer(strComputer,"wmi")
  
  Set colItems = objWMIService.ExecQuery("SELECT Caption, Description, ProviderName FROM Win32_LogicalDisk")

    For Each objItem In colItems
      'msgbox replace(lcase(objItem.Caption),":","")
      
      If instr(objItem.Description,"Network Connection") Then
        objDictionary.Add replace(lcase(objItem.Caption),":",""), "   [" & ucase(objItem.ProviderName) & "]"
      Else
        objDictionary.Add replace(lcase(objItem.Caption),":",""), "   [" & objItem.Description & "]"
      End If
    Next

  AlphaArray = array("a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z")
  iTempCount = 0

  For each letter in AlphaArray

    If not objDictionary.Exists(letter) Then
      sLocalDriveToMap = ucase(letter) & ":"   
      exit for
    End If

  Next
  
  'msgbox "remote drive: " & sLocalDriveToMap
  
  on error resume next
	strStatus = strPrepareFont & Time & " - " & "Making sure directory " & sLocalDriveToMap & "\" & sRemoteFolder1 & sVNCType & " exists...<br>" & strStatus
	objdiv.innerhtml = strStatus
  
  Call RemoveNetworkDrive(sLocalDriveToMap)

  Call MapDrive(strComputer,"\\" & strComputer & "\" & replace(sRemoteDrive,":\","$"),sLocalDriveToMap)

	wscript.sleep 2000
	
  Call MakeSureDirectoryTreeExists(sLocalDriveToMap & "\" & sRemoteFolder1 & sVNCType)

	strStatus = strPrepareFont & Time & " - " & "Preparing to copy " & strFilestoCopy & "<br>" & strStatus
	
	objdiv.innerhtml = StrStatus
	
	strFiles = split(strFilestoCopy,";")
  
  err.clear
	
  For i = 0 to UBound(strFiles) 
	  On Error goto 0

		Call fctCopyFile(strFoldertoCopyFrom & "\" & sVNCType & "\" & strFiles(i),sLocaldrivetoMap & "\" & sRemoteFolder1 & sVNCType,strFiles(i))
	Next
	
  Call RemoveNetworkDrive(sLocalDriveToMap)

  blnUpdate = "true"
  
  on error resume next
  
	Call SetRemoteOptions(strComputer)
	
	'Run the winvnc4 exe to register the service
	Call RunProcess(Chr(34) & replace(replace(strFoldertoCopyTo,"$\",":\"),"\\" & strComputer & "\","") & "\" & sVNCExe & Chr(34) & " " & sInstallSwitch,strComputer)
	
	'Give the system enough time to register the service prior to starting it.
	wscript.sleep 3000
	
	'Start the VNC service
	Call ChangeVNCServiceState(strComputer,"start")
	

	'Run the viewer on the local computer against the remote computer'
	Call startVNCViewer(strComputer)
  

  
	strStatus = strPrepareFont & Time & " - " & "VNCViewer has terminated. <br>" & strstatus
	objdiv.innerhtml = strStatus

  	
	If bStop = true then Call ChangeVNCServiceState(strComputer,"stop")

	strStatus = strPrepareFont & Time & " - " & "Finished with the VNCviewer.  You can close this window when ready.<br>" & strstatus
	objdiv.innerhtml = strStatus
	
	If blnUnregister = true then
    Call RunProcess(chr(34) & replace(replace(strFoldertoCopyTo,"$\",":\"),"\\" & strComputer & "\","") & "\" & sVNCExe & chr(34) & " " & sRemoveSwitch,strComputer)
  	strStatus = strPrepareFont & Time & " - " & "Unregistering service from " & strComputer & ".<br>" & strstatus
	  objdiv.innerhtml = strStatus
	End If

  wscript.sleep 7000
  
  on error resume next
  
	If bCloseStatusWindow = true then IE.quit
End If

Function StartVNCViewer(strComputer)
	
  If bExempt = false then 
    strStatus = "<font color='yellow' size='0'>" & Time & " - Remote user on [" & strComputer & "] will be prompted to accept or reject the anonymous VNC connection.<br>" & strstatus
		objdiv.innerhtml = strStatus
	End if
  wshshell.run strVNCViewer & " " & strComputer & ":0",1,true 
	'WshShell.AppActivate "VNC Viewer"
End Function

Sub SetRemoteOptions(strComputer)
	WshShell.RegWrite "HKCU\Software\RDScripts\InstallVNC\LastComputer", strComputer, "REG_SZ"
	WshShell.RegWrite "HKCU\Software\RDScripts\InstallVNC\VNCType", sVNCType, "REG_SZ"
  
  Call ConnectToComputer(strComputer,"reg")
  
  'Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}//" & strComputer & "/root/default:StdRegProv")

  on error resume next
  
  hTree		= HKEY_LOCAL_MACHINE
  sKey		= sRegRoot

  
  'Create the VNC reg key on the remote station
  oreg.CreateKey hTree,sKey

  sValueName  = "Password"
  oreg.GetBinaryValue hTree,sKey,sValueName,sValue

  For i = lBound(sValue) to uBound(sValue)
    bPassword = bPassword & delim & Hex(sValue(i))
    delim = ","
  Next

  'msgbox "Password key is set to: " & lcase(bPassword)
  
  if cstr(err.number) <> "0" then
    blnUpdate = "true"
  End if
  

  Set f = ws.OpenTextFile(strFoldertoCopyFrom & "\" & strRegSettings, ForReading)
  Do While f.AtEndOfStream <> True
   	  strReadLine = f.readline
      If InStr(LCase(strReadLine),"password") Then
	      RLArray = split(strReadLine,":")
      End If
  Loop

  strUpdatePassword = RLArray(1)

  'MsgBox strUpdatePassword
  strPassword = split(strUpdatePassword,",")

  For n = 0 to UBound(strPassword)
     'cint("&H" & hexString)
	   strPassword(n) = CInt("&H" & strPassword(n))
	   'MsgBox strPassword(n)
	Next

 'MsgBox join(strPassword,",")

  f.close

  If lcase(bPassword) = lcase(strUpdatePassword) then
    'msgbox "Passwords match."
  Else
    strResult = msgbox("The configured password on the remote system and the password " _
     & "specified in your registry key [" & strRegSettings & "] do not match.  Would you like to " _
     & "force the remote system's password to match the specified registry file entry?",36,"" _
     & "Passwords do not match")
     
    If strResult = 6 then
      blnUpdate = "true"
    Else
      blnUpdate = "false"
    End if
  End if
  

  On Error goto 0
  
  If blnUpdate = "true" then

	      dim strRegKey

 	      sValue		= "Password"

	      'InputBox "value","test",join(strPassword,",")

        'set the remote password
        'msgbox sValue
        oreg.SetBinaryValue HKEY_LOCAL_MACHINE,sKey,sValue,strpassword
        
      	If Err.Number = 0 Then
		      	strStatus = strPrepareFont & Time & " - " & "Added VNC password to remote registry successfully. <br>" & strstatus
		      	objdiv.innerhtml = strStatus
	      Else
    	      ' An error occurred
			      strStatus = strPrepareFont & Time & " - " & err.description & " - Error adding VNC password to remote registry. <br>" & strstatus
			      objdiv.innerhtml = strStatus
  	        Wscript.Echo err.number & ": " & err.description
	      End If
  End If

    sValue		= "PortNumber"
  	dwValue   = iPortNumber
		oreg.SetDWordValue HKEY_LOCAL_MACHINE,sKey,sValue,dwValue

    sValue		= "HTTPPortNumber"
  	dwValue   = iHTTPPortNumber
		oreg.SetDWordValue HKEY_LOCAL_MACHINE,sKey,sValue,dwValue

  On Error goto 0
  
  Call CheckComputerName(strComputer)

  If instr(strRemoteOSVersion,"2003") or instr(lcase(strRemoteOSVersion),"server") or instr(lcase(strRemoteOSVersion),"powered") then
  ' This is a server OS - don't do anything here
  
  Else 'only apply these query connect reg settings to workstations
    sValue		= "QueryConnect"
    If iQueryConnect=1 and bExempt <> true then
  	   dwValue   = 1
    ElseIf iQueryConnect=0 or bExempt = true then
       dwValue = 0
    End If

    'Set QueryConnect value
    oreg.SetDWordValue HKEY_LOCAL_MACHINE,sKey,sValue,dwValue

    sValue		= "QueryOnlyIfLoggedOn"
    If iQueryOnlyIfLoggedOn=1 and bExempt <> true then
  	    dwValue   = 1
    ElseIf iQueryOnlyIfLoggedOn=0 or bExempt = true then
  	    dwValue   = 0
    End If

    'Set QueryOnlyIfLoggedOn value
    oreg.SetDWordValue HKEY_LOCAL_MACHINE,sKey,sValue,dwValue

    sValue = "AutoPortSelect"
    dwValue = 0
    
    oreg.SetDWordValue HKEY_LOCAL_MACHINE,sKey,sValue,dwValue

    sValue = "PortNumber"
    dwValue = iPortNumber
    
    oreg.SetDWordValue HKEY_LOCAL_MACHINE,sKey,sValue,dwValue

    sValue = "HTTPPortNumber"
    dwValue = iHTTPPortNumber
    
    oreg.SetDWordValue HKEY_LOCAL_MACHINE,sKey,sValue,dwValue
    

    If Err.Number <> 0 Then

		  	strStatus = strPrepareFont & Time & " - " & "Error modifiying query connect and login settings.<br>" & strstatus
			  objdiv.innerhtml = strStatus
        Wscript.Echo err.number & ": " & err.description
    	   ' An error occurred
    End If
 	End If

End Sub

Sub CheckComputerName(strComputer)
  bExempt = false
  If strRegSettings = sServerRegistry then bExempt = true
  on error goto 0
  Set e = ws.OpenTextFile(strFoldertoCopyFrom & "\query_exempt.txt", ForReading)
  Do While e.AtEndOfStream <> True
   	strReadLine = e.readline
  	If Instr(Lcase(strReadLine),lcase("exclude all workstations")) Then
  	   bExempt = true
  	   exit sub
  	End If
  	If InStr(LCase(strReadLine),lcase(strComputer)) Then
	    bExempt = true
		  strStatus = "<font color='yellow' size='0'>" & Time & " - " & ucase(strComputer) & " is exempt from QueryConnect/remote prompt settings.<br>" & strstatus
			objdiv.innerhtml = strStatus
    End If
 	Loop
End Sub


Function GetOSVersion(strComputer)
	On Error resume next
	Call ConnectToComputer(strComputer,"wmi")
  'Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select Caption, SystemDrive from Win32_OperatingSystem",,48)
	For Each objItem in colItems
	
	  strRemoteSystemDrive = objItem.SystemDrive
		strRemoteOSVersion = objItem.Caption
		'msgbox strRemoteOSVersion
  Next
  If err.number <> 0 then
  		strStatus = strPrepareFont & Time & " - " & "Cannot detect remote OS.  You may not have sufficient permissions "_
  		 & "on the queried computer.  Script will attempt to connect using workstation credentials.<br>" & strstatus
	
		objdiv.innerhtml = strStatus

   End If
End Function

If cstr(err.number) <> 0 then
  Call ErrHandler(err.number,err.description) 
End If

Function ChangeVNCServiceState(strComputer,strState)
If strState = "start" Then 
	On Error Resume Next
	'MsgBox "Trying to start!"
	strStatus = strPrepareFont & Time & " - " & "Attempting to start " & sVNCServiceName & " service on " & strComputer & ".<br>" & strStatus
	objdiv.innerhtml = strstatus
	Dim ServiceSet, Service, svcState
	'strComputer = "."
	
  Call ConnectToComputer(strComputer,"wmi")  
	Set ServiceSet = objWMIService.ExecQuery("select description, name, state from Win32_Service where name='" & sVNCServiceName & "'",,48)
	
	'Set ServiceSet = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & strComputer).ExecQuery("select Description,Name,State from Win32_Service where name='" & sVNCServiceName & "'")
	
	If cstr(err.number) <> 0 then
	  Call ErrHandler(err.number,err.description) 
	  Exit Function
	End If
	wscript.sleep 1000

 	   strStatus = strPrepareFont & Time & " - Setting " & service.name & " to '" & sStartupMode & "'...<br>" & strStatus
     	   objdiv.innerhtml = strStatus
	
	For each Service in ServiceSet
	   On Error Resume Next
     
     tryCount = 2
	   intcount = 0

	   Service.ChangeStartMode(sStartupMode)
	
	   'If InStr(LCase(service.name),"vnc") Then wscript.echo service.name
	   'If LCase(Service.Name) = "" & sVNCServiceName & "" and LCase(service.state) <> "started" Then 'Wscript.Echo Service.State
	      
	      Do While intcount < trycount
	    		Service.StartService()

        	'wscript.sleep 5000
		     	intCount = intcount + 1

	        If cstr(err.number) <> 0 then

            Call ErrHandler(err.number,err.description) 
	        End If
	
   	  	wscript.sleep 1000

	      Loop

	  ChangeVNCServiceState = "running"
    Next

ElseIf lcase(strState) = "stop" Then
	On Error goto 0
	strStatus = strPrepareFont & Time & " - " & "Attempting to stop VNC service on " & strComputer & ".<br>" & strStatus
	objdiv.innerhtml = strstatus
    
  'objWMIService.Security_.ImpersonationLevel = 3
  Call ConnectToComputer(strComputer,"wmi")  
	Set ServiceSet = objWMIService.ExecQuery("select description, name, state from Win32_Service where name='" & sVNCServiceName & "'",,48)
	'Set ServiceSet = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & strComputer).ExecQuery("select Description,Name,State from Win32_Service where name='" & sVNCServiceName & "'")

	If cstr(err.number) <> 0 then
	  Call ErrHandler(err.number,err.description) 
	  
	  Exit Function
	End If

	For each Service in ServiceSet
			
    strStatus = strPrepareFont & Time & " - Verifying " & service.name & " is set to '" & sStartupMode & "'...<br>" & strStatus

     objdiv.innerhtml = strStatus
	   tryCount = 2
	   Dim intcount
	   intcount = 0
	   
     on error resume next
	   Service.ChangeStartMode(sStartupMode)

	   'If InStr(LCase(service.name),"vnc") Then wscript.echo service.name
	   'If LCase(Service.Name) = "" & sVNCServiceName & "" and LCase(service.state) <> "stopped" Then
	        
	        Do While intcount < trycount
	    		Service.StopService()
		    	
		    	'wscript.sleep 5000
		     	
		     	intCount = intCount + 1
	        	If cstr(err.number) <> 0 then
	          	Call ErrHandler(err.number,err.description) 
	        	err.clear
	        	End If
   		  	wscript.sleep 1500
		Loop   	
	
	   'End If
	Next
	ChangeVNCServiceState = "stopped"
End If

End Function

Function CheckVNCService(strComputer)
strStatus = strPrepareFont & Time & " - " & "Checking for existence of VNC service on " & strcomputer & ".<br>" & strStatus
objdiv.innerhtml = strstatus

On Error Resume Next
  'msgbox blnUpdate
  Set ServiceSet = objWMIService.ExecQuery _
    ("select Description,Name,State from Win32_Service where name='" & sVNCServiceName & "'",,48)
  'Set ServiceSet = GetObject("winmgmts:{impersonationLevel=impersonate}!//" & strComputer).ExecQuery("select Description,Name,State from Win32_Service where name='" & sVNCServiceName & "'")

If err.number <> 0 Then
  'MsgBox err.number
  If err.number = 462 Then
  	blnFatal = true

  	Call ErrHandler(err.number,"This computer is not responding.  It may have a firewall enabled, or permissions locked down." _
  	 & "  Installation cannot continue.")
  	Exit Function
  Else 
  	Call ErrHandler(err.number,err.description) 
  	blnFatal = true
  End If

End If

For each Service in ServiceSet
   strStatus = strPrepareFont & Time & " - " & service.name & " is " & service.state & ".<br>" & strStatus
   objdiv.innerhtml = strStatus
   strStatus = strPrepareFont & Time & " - Setting " & service.name & " to '" & sStartupMode & "'...<br>" & strStatus
   objdiv.innerhtml = strStatus
   
   Service.ChangeStartMode(sStartupMode)

  If cstr(err.number) <> 0 then
    Call ErrHandler(err.number,err.description) 
  End If
   
   If LCase(service.state) <> "running" Then 
   	Call ChangeVNCServiceState(strComputer,"start")
   	CheckVNCService = "installed"	
   ElseIf LCase(service.state) <> "stopped" Then
   	CheckVNCService = "installed"	
   End If
   
   intCount = 0 
Next
End Function

Function RunProcess(strCommand,strComputer)
  objdiv.innerhtml = strstatus 
  
  On Error goto 0
  
  strStatus = strPrepareFont & Time & " - " & "Running command " & strCommand & " on " & strComputer & "<br>" & strStatus
  strServer = strComputer
 
  strArgs = " "
  StrExeName = strCommand
  strCurrentDir = strRemoteSystemDrive
  
  If strUserCredentials = "" then 
    Set objService = objLocator.ConnectServer(strComputer,"root/cimv2")
  Else  
    Set objService = objLocator.ConnectServer(strComputer,"root/cimv2",  strUserCredentials,strPasswordCredentials)
  End If
  
  Set objProcess = objService.Get("WIN32_Process")
  Set objProcessStartup = objService.Get("Win32_ProcessStartup")
  objProcessStartup.PriorityClass = 128
  objProcessStartup.ShowWindow = 1
  Set objMethod = objProcess.Methods_("Create")
  Set objInParameters = objMethod.inParameters.SpawnInstance_()
  objInParameters.CommandLine = strExeName & strArgs
  objInParameters.CurrentDirectory = strCurrentDir
  Set objInParameters.ProcessStartupInformation = objProcessStartup

  Set objOutParameters = objProcess.ExecMethod_("Create", objInParameters)

  'msgbox "Method returned result = " & objOutParameters.returnValue
  If objOutParameters.returnValue = 0 Then
    'msgbox "Id of new process is " & objOutParameters.ProcessID
    strPID = objOutParameters.ProcessID
  Else
    'msgbox "failed"
    'Log "Process creation failed."
  End If

  dim errDescription

If objOutParameters.returnValue = 0 Then errdescription = "Successfully created process on " & strComputer & " with PID: " & strPID
If objOutParameters.returnValue = 2 Then errdescription = "<font color='red'>Access denied</font>"
If objOutParameters.returnValue = 3 Then errdescription = "<font color='red'>Insufficient privileges to create a process on " & strComputer & "</font>"
If objOutParameters.returnValue = 9 Then errdescription = "Path not found for " & strCommand & " on " & strComputer
    
    strStatus = strPrepareFont & Time & " - " & errdescription & "<br>" & strStatus
    'msgbox errdescription
    objdiv.innerhtml = strStatus
  'msgbox intreturn
End Function

Sub ErrHandler(strErrNumber,stErrDescription)
  strStatus = strPrepareFont & Time & " - Error " & err.number & ": " & err.description & "<br>" & strstatus
  objDiv.innerhtml = strStatus

  If blnFatal = "true" Then
  	strStatus "<font color'" & fcolor & "'>" & Time & " - VNC installation cannot continue.  Now exiting installation process.<br>" & strstatus
  	wscript.quit
  End If
End Sub

' The MakeSureDirectoryTreeExists Function

' Although the FSO model doesn't have a direct method to create nested
' folders, you can use the following function. This VBScript function uses
' VBScript's Split function to break the folder path it receives into
' components. From those components, the MakeSureDirectoryTreeExists
' creates subfolders, one at a time. Because the function checks for the
' folder's existence before proceeding, you can pass it any tree, as long as
' you make sure that, after it returns, the entire tree exists as you specified.
' With the MakeSureDirectoryTreeExists function, a call such as

'	MakeSureDirectoryTreeExists "C:\one\two\three"

' is legitimate and won't result in an error message.

Function MakeSureDirectoryTreeExists(dirName)
Dim aFolders, newFolder
	On Error Resume Next
	dim delim
	' Creates the FSO object.
	Set fso = CreateObject("Scripting.FileSystemObject")

	' Checks the folder's existence.
	If Not fso.FolderExists(dirName) Then

		' Splits the various components of the folder name.
		If instr(dirname,"\\") then
		    delim = "-_-_-_-"
			dirname = replace(dirname,"\\",delim)
			'wscript.echo dirname
		End if
		aFolders = split(dirName, "\")
		if instr(dirname,delim) Then
			dirname = replace(aFolders(0),delim,"\\")
			'wscript.echo "aFolders = " & dirname
		End if
		' Obtains the drive's root folder.
		
		newFolder = fso.BuildPath(dirname, "\")
	
		' Scans each component in the array, and create the appropriate folder.
		For i=1 to UBound(aFolders)
			newFolder = fso.BuildPath(newFolder, aFolders(i))

			If Not fso.FolderExists(newFolder) Then

				fso.CreateFolder newFolder
				If CStr(err.number) = 70 or CStr(err.number) = 76 Then 
					strStatus = strPrepareFont & Time & " - You do not appear to have appropriate permissions on " & strComputer & " to " _
					 & "install the remote control client. Installation aborted.<br>" & strStatus
					'MsgBox strStatus
					objdiv.innerhtml = strStatus
					wscript.quit
				End If
			End If
		Next
	End If
End Function

Sub MapDrive(strComputer,strRemoteName,strLocalName)
  on error resume next
  Set WshNetwork = WScript.CreateObject("WScript.Network")
  WshNetwork.MapNetworkDrive strLocalName, strRemoteName , false, strUserCredentials, strPasswordCredentials
  'msgbox strLocalName & " --- " & strRemoteName
  if err.number <> 0 then 
    msgbox "You may not have permissions to the resource specified (" & strRemoteName & ").  Try using alternate credentials, or retry the connection using the computer's IP address instead." & vbcrlf & vbcrlf & "Actual error was: " & err.description,48,"Problem with permissions or access"
  End If
End Sub

Sub RemoveNetworkDrive(strLocalName)
  Set WshNetwork = WScript.CreateObject("WScript.Network")
  WshNetwork.RemoveNetworkDrive strLocalName, true 
End Sub

Function fctCopyFile(strSource,strFoldertoCopyTo,strFileName)
  On Error Resume Next

 	
	Set fso = CreateObject("Scripting.FileSystemObject")
	strStatus = strPrepareFont & Time & " - " & "Copying " & strSource & " to " & strFoldertoCopyTo & "\" & strFileName & "...<br>" & strStatus
	objdiv.innerhtml = strStatus	
	fso.CopyFile strSource, strFoldertoCopyTo & "\" & strFileName
	
	If CStr(err.number) <> 0 Then 
		
		If err.number = 76 Then
			blnFatal = true
			'Call Errhandler(err.number,"Cannot copy files to the remote system.  Aborting install.")
		End If
		Call ErrHandler(err.number,err.description)
	End If
End Function

Function IEStatus
If blnProgressMode Then
	If blnDebugMode Then
		dbgTitle = "VNC installation tool"
	Else
		dbgTitle = "VNC installation tool"
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
	dbgLeft = 200
	dbgTop = 200
	dbgVisible = True
	dlgBarWidth = 380
	dlgBarHeight = 23 
	dlgBarTop = 80
	dlgBarLeft = 50
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
	
	Set IE = CreateObject("InternetExplorer.Application")
	'strScriptVer = "version would go here"

	strTempFile = WshSysEnv("TEMP") & "\progress.htm"
	ws.CreateTextFile (strTempFile)
  Set f1 = ws.GetFile(strTempFile)
  Set ts = f1.OpenAsTextStream(2, True)
  ts.WriteLine("<!-- saved from url=(0014)about:internet -->")
  ts.WriteLine("<html><head><title>" & dbgTitle & " " & strScriptVer & " </title>")
  ts.WriteLine("<style>.errortext {color:red}")
  ts.WriteLine(".hightext {color:blue}")
  ts.WriteLine("Body {scrollbar-base-color:black;background: url(" & strScriptPath & "vnc\" & strBackground & ");}</style>")
	ts.WriteLine("</head>")
	ts.WriteLine(strHDRCode & " <br><strong><font size='2' color='" & hdrfcolor & "' face='verdana'>" _
	 	& "&nbsp VNC Installation status on <font color'=" & hdrfcolor & ">" & strComputer & "...</font></strong><br>" _
	 	& "&nbsp &nbsp<br>")
	ts.WriteLine("<center><table width='100%' style='" & stsbgcolor & "'><tr><td>")
	If blnDebugMode Then
		ts.WriteLine("<body scroll='yes' topmargin='0' leftmargin='0'"_
		& " style='font-family: " & fstyle & "; font-size: 0.6em color: #000000;"_
		& " font-weight: bold; text-align: left'><center><font face=" & fstyle & ">"_
		& " <font size='0.8em'> <hr color='#FBC114'>")
	Else
		ts.WriteLine("<body scroll='no' topmargin='0' leftmargin='0' "_
		& " style='font-family: " & fstyle & "; font-size: 0.6em color: #000000;"_
		& " font-weight: bold; text-align: left'><center><font face=" & fstyle & ">"_
		& " <font size='0.8em'> <hr color='#FBC114'>")
	End If
	ts.WriteLine("<div style='background-color:" & stsBGColor & "' id='ProgObject' align='left'align='left' style='width: 450px;height: 140px;overflow:auto'></div><hr color='#FBC114'>")
	'If blnDebugMode Then
		ts.WriteLine("<div id='ProgDebug' align='left'></div>")
	'End If

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
	'If blnDebugMode Then
		Set objDBG = IE.Document.All("ProgDebug")
	'End If
	Set objFlash = IE.Document.All("ProgFlash")
	Set objPBar = IE.Document.All("ProgBarId")
	Set objBar = IE.Document
End If

End Function

'*******************************************************************
'*	Name:	fctSetupIE
'*	Function:	Setup an IE windows of 540 x 200 to display 
'* 	progress information.
'*******************************************************************
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
	On Error GoTo 0
	wshshell.AppActivate("Microsoft Internet Explorer")
End Sub
'------------------------------------------------------------------------------------------

Sub WriteINIString(Section, KeyName, Value, FileName)
  Dim INIContents, PosSection, PosEndSection
  'Get contents of the INI file As a string
  INIContents = GetFile(FileName)

  'Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)
  If PosSection>0 Then
    'Section exists. Find end of section
    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    '?Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1

    'Separate section contents
    Dim OldsContents, NewsContents, Line
    Dim sKeyName, Found
    OldsContents = Mid(INIContents, PosSection, PosEndSection - PosSection)
    OldsContents = split(OldsContents, vbCrLf)

    'Temp variable To find a Key
    sKeyName = LCase(KeyName & "=")

    'Enumerate section lines
    For Each Line In OldsContents
      If LCase(Left(Line, Len(sKeyName))) = sKeyName Then
        Line = KeyName & "=" & Value
        Found = True
      End If
      NewsContents = NewsContents & Line & vbCrLf
    Next

    If isempty(Found) Then
      'key Not found - add it at the end of section
      NewsContents = NewsContents & KeyName & "=" & Value
    Else
      'remove last vbCrLf - the vbCrLf is at PosEndSection
      NewsContents = Left(NewsContents, Len(NewsContents) - 2)
    End If

    'Combine pre-section, new section And post-section data.
    INIContents = Left(INIContents, PosSection-1) & _
      NewsContents & Mid(INIContents, PosEndSection)
  else'if PosSection>0 Then
    'Section Not found. Add section data at the end of file contents.
    If Right(INIContents, 2) <> vbCrLf And Len(INIContents)>0 Then
      INIContents = INIContents & vbCrLf
    End If
    INIContents = INIContents & "[" & Section & "]" & vbCrLf & _
      KeyName & "=" & Value
  end if'if PosSection>0 Then
  WriteFile FileName, INIContents
End Sub

Function WriteFile(ByVal FileName, ByVal Contents)
Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
On Error Resume Next

'Go To windows folder If full path Not specified.
If InStr(FileName, ":\") = 0 And Left (FileName,2)<>"\\" Then 
FileName = FS.GetSpecialFolder(0) & "\" & FileName
End If

Dim OutStream: Set OutStream = FS.OpenTextFile(FileName, 2, True)
OutStream.Write Contents
End Function

Function GetINIString(Section, KeyName, Default, FileName)
  Dim INIContents, PosSection, PosEndSection, sContents, Value, Found

  'Get contents of the INI file As a string
  INIContents = GetFile(FileName)
  'Find section
  PosSection = InStr(1, INIContents, "[" & Section & "]", vbTextCompare)

  If PosSection > 0 Then
    'Section exists. Find end of section

    PosEndSection = InStr(PosSection, INIContents, vbCrLf & "[")
    '?Is this last section?
    If PosEndSection = 0 Then PosEndSection = Len(INIContents)+1

    'Separate section contents
    sContents = Mid(INIContents, PosSection, PosEndSection - PosSection)

    If InStr(1, sContents, vbCrLf & KeyName & "=", vbTextCompare)>0 Then
      Found = True
      'Separate value of a key.
      Value = SeparateField(sContents, vbCrLf & KeyName & "=", vbCrLf)
    End If
  End If
  If isempty(Found) Then Value = Default
  GetINIString = Value
End Function

'Separates one field between sStart And sEnd
Function SeparateField(ByVal sFrom, ByVal sStart, ByVal sEnd)
  Dim PosB: PosB = InStr(1, sFrom, sStart, 1)
  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    Dim PosE: PosE = InStr(PosB, sFrom, sEnd, 1)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf, 1)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(sFrom, PosB, PosE - PosB)
  End If
End Function

'File functions
Function GetFile(ByVal FileName)

  Set FS = CreateObject("Scripting.FileSystemObject")

  'Go To windows folder If full path Not specified.
  If InStr(FileName, ":\") = 0 And Left(FileName,2)<> "\\" And Left(FileName,2) <> ".\" Then
    FileName = FS.GetSpecialFolder(0) & "\" & FileName
  End If
  On Error Resume Next

  GetFile = FS.OpenTextFile(FileName).ReadAll
  wscript.echo getfile
End Function

Sub WriteINIStringVirtual(Section, KeyName, Value, FileName)
  WriteINIString Section, KeyName, Value, _
    Server.MapPath(FileName)
End Sub

Function GetINIStringVirtual(Section, KeyName, Default, FileName)
  GetINIStringVirtual = GetINIString(Section, KeyName, Default, _
    Server.MapPath(FileName))
End Function
