Option Explicit

Dim strDocumentAuthor, strComputer

' Who Authored the document
strDocumentAuthor = ""


' Script version
Dim strScriptVersion
strScriptVersion = "2.0"

' Fonts to use in document
Dim strFontBodyText, strFontHeading1, strFontHeading2, strFontHeading3, strFontHeading4, strFontTitle, strFontTOC1, strFontTOC2, strFontTOC3
Dim strFontHeader, strFontFooter
strFontBodyText = "Arial"
strFontHeading1 = "Trebuchet MS"
strFontHeading2 = "Trebuchet MS"
strFontHeading3 = "Trebuchet MS"
strFontHeading4 = "Trebuchet MS"
strFontTitle = "Trebuchet MS"
strFontTOC1 = "Trebuchet MS"
strFontTOC2 = "Trebuchet MS"
strFontTOC3 = "Trebuchet MS"
strFontHeader = "Arial"
strFontFooter = "Arial"
Dim nBaseFontSize


'==========================================================
'==========================================================
' Variable Declarations and Constants


' A few counters are always handy
Dim i, j

' Variables for the Win32_BIOS class
Dim strBIOS_SMBIOSBIOSVersion, strBIOS_SMBIOSMajorVersion, strBIOS_SMBIOSMinorVersion, strBIOS_Version
Dim arrBIOS_BiosCharacteristics
Dim strBiosCharacteristics
ReDim arrBIOS_BiosCharacteristics(0)

' Variables for the Win32_Win32_CDROMDrive class
Dim objDbrCDROMDrive

' Variables for the Win32_ComputerSystem class
Dim strComputerSystem_Name, strComputerSystem_TotalPhysicalMemory, strComputerSystem_Domain, nComputerSystem_DomainRole
Dim strTotalPhysicalMemoryMB, strDomainType, strComputerRole

' Variables for the Win32_ComputerSystemProduct class
Dim strComputerSystemProduct_Manufacturer, strComputerSystemProduct_Name, strComputerSystemProduct_IdentifyingNumber

' Variables For the Win32_DiskDrive, Win32_DiskPartition and Win32_LogicalDisk classes
Dim objDbrDrives, objDbrDisks, objDiskPartitions, objDiskPartition, objLogicalDisks, objLogicalDisk

' Variables for the Win32_Group class
Dim objDbrLocalGroups

' Variables for the Win32_GroupUser class
Dim objDbrGroupUser

' Variables for the Win32_IP4RouteTable class
Dim objDbrIP4RouteTable

' Variables for the Win32_NetworkAdapterConfiguration Class
Dim nMaxNetworkAdapters
nMaxNetworkAdapters = 50
Dim arrNetadapter_Description, arrNetadapter_MACAddress, arrNetadapter_DNSHostName, arrNetadapter_DHCPEnabled
Dim arrNetadapter_DHCPServer, arrNetadapter_WINSPrimaryServer, arrNetadapter_WINSSecondaryServer, arrNetadapter_IPAddress
Dim arrNetadapter_IPSubnet, arrNetadapter_DefaultIPGateway, arrNetadapter_DNSServerSearchOrder, arrNetadapter_DNS
ReDim arrNetadapter_Description(0), arrNetadapter_MACAddress(0), arrNetadapter_DNSHostName(0), arrNetadapter_DHCPEnabled(0)
ReDim arrNetadapter_DHCPServer(0), arrNetadapter_DNS(0), arrNetadapter_WINSPrimaryServer(0), arrNetadapter_WINSSecondaryServer(0)
ReDim arrNetadapter_IPAddress(nMaxNetworkAdapters,0), arrNetadapter_IPSubnet(nMaxNetworkAdapters,0)
ReDim arrNetadapter_DefaultIPGateway(nMaxNetworkAdapters,0), arrNetadapter_DNSServerSearchOrder(nMaxNetworkAdapters,0)
Dim strDNSServers

' Variables for the Win32_NTEventLogFile Class
Dim objDbrEventLogFile

' Variables for the Win32_OperatingSystems Class
Dim strOperatingSystem_InstallDate, strOperatingSystem_Caption, strOperatingSystem_ServicePack, strOperatingSystem_WindowsDirectory
Dim strOperatingSystem_LanguageCode, arrOperatingSystem_Name

' Variables for the Win32_PageFile Class
Dim objDbrPagefile

' Variables for the Win32_PhysicalMemory Class
Dim objDbrPhysicalMemory

' Variables for the Win32_Printer class
Dim objDbrPrinters

' Variables for the Win32_Process class
Dim objDbrProcess

' Variables for the Win32_Processor class
Dim strProcessor_MaxClockSpeed, strProcessor_L2CacheSize, strProcessor_Name, strProcessor_Description, strProcessor_ExtClock
Dim objDbrProcessorSockets, bProcessorHTSystem, intProcessors

' Variables for the Win32_Product class
Dim objDbrProducts

' Variables for the Win32_Registry class
Dim nCurrentSize, nMaximumSize

' Variables for the Win32_Services class
Dim objDbrServices

' Variables for the Win32_Share class
Dim objDbrShares

' Variables for the Win32_SoundDevice class
Dim objDbrSoundDevice

' Variables for the Win32_StartupCommand class
Dim objDbrStartupCommand

' Variables for the Win32_SystemEnclosure class
Dim nChassisType, strChassisType

' Variables for the Win32_TapeDrive class
Dim objDbrTapeDrive, bHasTapeDrive

' Variables for the Win32_TimeZone class
Dim strTimeZone

' Variables for the Win32_QuickFixEngineering Class
Dim objDbrPatches

' Variables for the Win32_UserAccount class
Dim objDbrLocalAccounts

' Variables for the Win32_VideoController class
Dim objDbrVideoController

' Variables from registry
Dim objDbrWindowsComponents, strWindowsComponentsClass
Dim objDbrRegPrograms
Dim strPrimaryDomain, strPrintSpoolLocation
Dim strLastUser, strLastUserDomain

' Variables for IIS Server
Dim objDbrIISWebServerSetting, objDbrIISWebServerBindings
Dim objDbrIISVirtualDirSetting

' Variables for System Roles
Dim bRoleDC, bRoleDHCP, bRoleDNS, bRoleFile, bRoleFTP, bRoleIAS
Dim bRoleMediaServer, bRoleNews, bRolePKI, bRolePrint, bRoleRAS
Dim bRoleRIS, bRoleSMTP, bRoleSQL, bRoleTS, bRoleWINS, bRoleWWW
Dim objDbrSystemRoles, nTerminalServerMode

' Variables to handle different versions of Windows
Dim nOperatingSystemLevel

' Variables for other WMI Providers
Dim bHasMicrosoftIISv2
bHasMicrosoftIISv2 = False

' Objects for WMI and Word
Dim objWMIService, colItems, objItem, oReg
Dim oWord, oListTemplate

' Variables for routines
Dim errGatherWMIInformation, errGatherRegInformation, errWin32_Product
Dim bAllowErrors

errGatherWMIInformation = False
errGatherRegInformation = False


' Variables for script options
' WMI
Dim bWMIBios, bWMIRegistry, bWMIApplications,bWMIPatches,bWMIFileShares, bWMIServices, bWMIPrinters
Dim bWMIEventLogFile, bWMILocalAccounts, bWMILocalGroups, bWMIIP4Routes, bWMIRunningProcesses
Dim bWMIHardware, bWMIStartupCommands
' Registry
Dim bRegDomainSuffix, bRegWindowsComponents, bRegPrintSpoolLocation, bRegLastUser
Dim bRegPrograms, bDoRegistryCheck
' Username and Password
Dim strUserName, strPassword
' Other
Dim bInvalidArgument, bDisplayHelp, bAlternateCredentials, bCheckVersion
' Word
Dim bShowWord, bWordExtras, bUseDOTFile, bSaveFile, bUseSpecificTable
Dim strDOTFile, strWordTable
' Export Options
Dim strExportFormat, strSaveFile
' XML Options
Dim strStylesheet, strXSLFreeText

' Constants
Const adVarChar = 200
Const MaxCharacters = 255

'==========================================================
'==========================================================
' Main Body

If LCase (Right (WScript.FullName, 11)) <> "cscript.exe" Then
    MsgBox "This script should be run from a command line (eg ""cscript.exe sydi-server.vbs"")", vbCritical, "Error"
    WScript.Quit
End If


' Get Options from user
GetOptions



If (bInvalidArgument) Then
	WScript.Echo "Invalid Arguments" & VbCrLf
	bDisplayHelp = True
End If

If (bDisplayHelp) Then
	DisplayHelp
Else
	If (bCheckVersion) Then
		CheckVersion
	End If
	If (strComputer = "") Then
		strComputer = InputBox("What Computer do you want to document (default=localhost)","Select Target",".")
	End If
	If (strComputer <> "") Then
		' Run the GatherWMIInformation() function and return the status
		' to errGatherInformation, if the function fails then the
		' rest is skipped. The same applies to GatherRegInformation
		' if it is successful we place the information in a
		' new word document
		errGatherWMIInformation = GatherWMIInformation()
		If (errGatherWMIInformation) Then
			If (bDoRegistryCheck) Then
				errGatherRegInformation = GatherRegInformation
			End If
			GetWMIProviderList
		End If

		If (bHasMicrosoftIISv2) Then ' Does the system have the WMI IIS Provider
			GatherIISInformation
		End If

		SystemRolesSet
		
		If (errGatherWMIInformation) Then
			Select Case strExportFormat
				Case "word"
					PopulateWordfile
				Case "xml"
					PopulateXMLFile
			End Select
		End If
	End If
End If

'==========================================================
'==========================================================
' Procedures

Sub CheckVersion
	Dim strURLSydiVersioncheck
	Dim objHTTP
	strURLSydiVersioncheck="http://sydi.sourceforge.net/versions.php?package=sydi-server"
	Set objHTTP = CreateObject("MSXML2.XMLHTTP")
	Call objHTTP.Open("GET", strURLSydiVersioncheck, FALSE)
	objHTTP.Send
	If (strScriptVersion = objHTTP.ResponseText) Then
		WScript.echo "You have the latest version (v." & strScriptVersion & ")"
	Else
		wscript.echo "A new version of SYDI-Server has been released!"
		wscript.echo "Your version: v." & strScriptVersion
		wscript.echo "Latest version: v." & objHTTP.ResponseText
		wscript.echo "Download it from http://sydi.sourceforge.net/"
	End If
	Wscript.Quit
End Sub ' CheckVersion

Function ConvertWMIDate(dUTCDate)
	' This is a standard function to convert WMI time to "normal" time.
	ConvertWMIDate = CDate(Mid(dUTCDate, 5, 2) & "/" &  Mid(dUTCDate, 7, 2) & "/" & Left(dUTCDate, 4) & " " & _
                          Mid (dUTCDate, 9, 2) & ":" &  Mid(dUTCDate, 11, 2) & ":" & Mid(dUTCDate, 13, 2))
End Function

Sub CalculateProcessorSockets(strSocketDesignation)
	strSocketDesignation = Replace(strSocketDesignation,"'","")
	objDbrProcessorSockets.Filter = " SocketDesignation='" & strSocketDesignation & "'"
	If (objDbrProcessorSockets.Bof) Then
		objDbrProcessorSockets.AddNew
		objDbrProcessorSockets("SocketDesignation") = strSocketDesignation
		objDbrProcessorSockets("Count") = 1
		objDbrProcessorSockets.Update
	Else
		objDbrProcessorSockets("Count") = objDbrProcessorSockets("Count") + 1
		objDbrProcessorSockets.Update
	End If
End Sub ' CalculateProcessorSockets


Sub DisplayHelp
	WScript.Echo "SYDI-Server v." & strScriptVersion
	WScript.Echo "Usage: cscript.exe sydi-server.vbs [options]"
	WScript.Echo "Examples: cscript.exe sydi-server.vbs -wabes -rc -f10 -tSERVER1"
	WScript.Echo "          cscript.exe sydi-server-vbs -ex -sh -o""H:\Server docs\DC1.xml -tDC1"""
	WScript.Echo "Gathering Options"
	WScript.Echo " -w	- WMI Options (Default: -wabefghipPqrsSu)"
 	WScript.Echo "   a	- Windows Installer Applications"
	WScript.Echo "   b	- BIOS Information"
 	WScript.Echo "   e	- Event Log files"
 	WScript.Echo "   f	- File Shares"
 	WScript.Echo "   g	- Local Groups (on non DC machines)"
 	WScript.Echo "   h	- Additional Hardware (ie. Video Controller)"
 	WScript.Echo "   i	- IP Routes (XP and 2003 only)"
 	WScript.Echo "   p	- Printers"
	WScript.Echo "   P	- Processes (running)"
 	WScript.Echo "   q	- Installed Patches"
 	WScript.Echo "   r	- Registry Size"
 	WScript.Echo "   s	- Services"
 	WScript.Echo "   S	- Startup Commands"
 	WScript.Echo "   u	- Local User accounts (on non DC machines)"
 	WScript.Echo " -r	- Registry Options (Default: -racdlp)"
 	WScript.Echo "   a	- Non Windows Installer Applications"
 	WScript.Echo "   c	- Windows Components"
 	WScript.Echo "   d	- FQDN Domain Name"
 	WScript.Echo "   l	- Last Logged on user"
 	WScript.Echo "   p	- Print Spooler Location"
 	WScript.Echo " -t	- Target Machine (Default: ask user)"
	WScript.Echo " -u	- Username (To run with different credentials)"
	WScript.Echo " -p	- Password (To run with different credentials, must be used with -u)"
	WScript.Echo "Output Options"
	WScript.Echo " -e	- Export format"
	WScript.Echo "   w	- Microsoft Word (Default)"
	WScript.Echo "   x	- XML (has to be used with -o)"
 	WScript.Echo " -o	- Save to file (-oc:\corpfiles\server1.doc, use in combination with -d"
 	WScript.Echo "   	  if you don't want to display word at all, use a Path or the file will"
 	WScript.Echo "  	  be placed in your default location usually 'My documents')"
 	WScript.Echo "  	  -oC:\corpfiles\server1.xml"
 	WScript.Echo "  	  WARNING USING -o WILL OVERWRITE TARGET FILE WITHOUT ASKING"
 	WScript.Echo "Word Options"
 	WScript.Echo " -b	- Use specific Word Table (-b""Table Contemporary"""
 	WScript.Echo "   	  or -b""Table List 4"")"
 	WScript.Echo " -f	- Base font size (Default: -f12)"
 	WScript.Echo " -d	- Don't display Word while writing (runs faster)"
 	WScript.Echo " -n	- No extras (minimize the text inside brackets)"
 	WScript.Echo " -T	- Use .dot file as template (-Tc:\corptemplates\server.dot, ignores -f)"
 	WScript.Echo "XML Options"
 	WScript.Echo " -s	- XML Stylesheet"
 	WScript.Echo "  h	- HTML"
 	WScript.Echo "  t	- Free text (-stE:\Files\mytransform.xsl or -stCORP.xsl)"
 	WScript.Echo "Other Options"
 	WScript.Echo " -v	- Check for latest version (requires Internet access)"
 	WScript.Echo " -D	- Debug mode, useful for reporting bugs"
 	WScript.Echo VbCrLf
 	WScript.Echo " -h	- Display help"
 	WScript.Echo VbCrLf
End Sub ' DisplayHelp

Function GatherIISInformation()
	If (bAllowErrors) Then
		On Error Resume Next
	End If
	Dim objItem2
	Const WbemAuthenticationLevelPktPrivacy = 6
	ReportProgress "Start subroutine: GatherIISInformation(" & strComputer & ")"
	Dim objSWbemLocator
	Dim arrGroupUser
	If (bAlternateCredentials) Then
		Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
		objSWbemLocator.Security_.AuthenticationLevel = WbemAuthenticationLevelPktPrivacy
		Set objWMIService = objSWbemLocator.ConnectServer(strComputer,"root\MicrosoftIISv2",strUserName,strPassword)
	Else
		'Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\MicrosoftIISv2")
		Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
		objSWbemLocator.Security_.AuthenticationLevel = WbemAuthenticationLevelPktPrivacy
		Set objWMIService = objSWbemLocator.ConnectServer(strComputer,"root\MicrosoftIISv2")
	End If
	If (Err <> 0) Then
	    ReportProgress Err.Number & " -- " &  Err.Description & " (" & strComputer & ")"
	    Err.Clear
	    GatherIISInformation = False
	    Exit Function
	End If
	
	
	ReportProgress " Gathering Web Server Settings"
	Set colItems = objWMIService.ExecQuery("Select Name, ServerBindings, ServerComment from IISWebServerSetting","WQL",48)
	Set objDbrIISWebServerSetting = CreateObject("ADOR.RecordSet")
	objDbrIISWebServerSetting.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrIISWebServerSetting.Fields.Append "ServerComment", adVarChar, MaxCharacters
	objDbrIISWebServerSetting.Open
	Set objDbrIISWebServerBindings = CreateObject("ADOR.RecordSet")
	objDbrIISWebServerBindings.Fields.Append "ServerName", adVarChar, MaxCharacters
	objDbrIISWebServerBindings.Fields.Append "Hostname", adVarChar, MaxCharacters
	objDbrIISWebServerBindings.Fields.Append "IP", adVarChar, MaxCharacters
	objDbrIISWebServerBindings.Fields.Append "Port", adVarChar, MaxCharacters
	objDbrIISWebServerBindings.Open
	
	For Each objItem In colItems
		objDbrIISWebServerSetting.AddNew
		objDbrIISWebServerSetting("Name") = objItem.Name
		objDbrIISWebServerSetting("ServerComment") = objItem.ServerComment
		For Each objItem2 In objItem.ServerBindings
			objDbrIISWebServerBindings.AddNew
			objDbrIISWebServerBindings("ServerName") = objItem.Name
			objDbrIISWebServerBindings("Hostname") = objItem2.Hostname
			objDbrIISWebServerBindings("Ip") = objItem2.Ip
			objDbrIISWebServerBindings("Port") = objItem2.Port
			objDbrIISWebServerBindings.Update
		Next
	    objDbrIISWebServerSetting.Update
	Next
	objDbrIISWebServerSetting.Sort = "ServerComment"
		
	' Virtual Directories
	Set colItems = objWMIService.ExecQuery("Select Name, Path from IISWebVirtualDirSetting","WQL",48)
	Set objDbrIISVirtualDirSetting = CreateObject("ADOR.RecordSet")
	objDbrIISVirtualDirSetting.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrIISVirtualDirSetting.Fields.Append "Path", adVarChar, MaxCharacters
	objDbrIISVirtualDirSetting.Open
	For Each objItem In colItems
		objDbrIISVirtualDirSetting.AddNew
		objDbrIISVirtualDirSetting("Name") = objItem.Name
		objDbrIISVirtualDirSetting("Path") = objItem.Path
		objDbrIISVirtualDirSetting.Update
	Next
	
	ReportProgress "End subroutine: GatherIISInformation(" & strComputer & ")"
End Function ' GatherIISInformation()

Function GatherRegInformation()
	Dim arrRegValueNames, arrRegValueTypes
	Dim dwValue
	Dim objRegLocator, objRegService
	Dim arrRegPrograms, strRegProgram, strRegProgramsDisplayName, strRegProgramsDisplayVersion, strRegProgramsTmp
	Dim bRegProgramsSkip
	Dim objRegExp
	
	Const HKEY_LOCAL_MACHINE = &H80000002
	If (bAllowErrors) Then
		On Error Resume Next
	End If
	ReportProgress VbCrLf & "Start subroutine: GatherRegInformation(" & strComputer & ")"
	
	If (bAlternateCredentials) Then
		Set objRegLocator = CreateObject("WbemScripting.SWbemLocator")
		Set objRegService = objRegLocator.ConnectServer(strComputer,"root\default",strUserName,strPassword)
		Set oReg = objRegService.Get("StdRegProv")
	Else
		Set oReg=GetObject("winmgmts:\\" &  strComputer & "\root\default:StdRegProv")
	End If
	
	If (Err <> 0) Then
	    Wscript.Echo Err.Number & " -- " &  Err.Description
	    Err.Clear
	    GatherRegInformation = False
	    Exit Function
	End If

	If (bRegDomainSuffix) Then
		ReportProgress " Reading domain information"
		oReg.GetStringValue HKEY_LOCAL_MACHINE,"SYSTEM\CurrentControlSet\Services\Tcpip\Parameters","Domain", strPrimaryDomain
	End If

	If (bRegPrintSpoolLocation) Then
		ReportProgress " Reading print spool location"
		oReg.GetStringValue HKEY_LOCAL_MACHINE,"SYSTEM\CurrentControlSet\Control\Print\Printers","DefaultSpoolDirectory", strPrintSpoolLocation
	End If

	' Checking Terminal Server Settings
	oReg.GetDwordValue HKEY_LOCAL_MACHINE,"SYSTEM\CurrentControlSet\Control\Terminal Server","TSAppCompat", nTerminalServerMode
	
	If (bRegPrograms) Then
		ReportProgress " Reading Programs from Registry"
		Set objDbrRegPrograms = CreateObject("ADOR.Recordset")
		objDbrRegPrograms.Fields.Append "DisplayName", adVarChar, MaxCharacters
		objDbrRegPrograms.Fields.Append "DisplayVersion", adVarChar, MaxCharacters
		objDbrRegPrograms.Open
		Set objRegExp = New RegExp
		oReg.EnumKey HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",arrRegPrograms
		For Each strRegProgram In arrRegPrograms
			bRegProgramsSkip = False
			oReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strRegProgram, "DisplayName", strRegProgramsDisplayName
			
			' Remove programs without Display Name
			If (IsNull(strRegProgramsDisplayName)) Then : bRegProgramsSkip = True : End If 
	
			' Remove MSI applications
			If (Len(strRegProgram) = 38) Then 
				strRegProgramsTmp = Left(strRegProgram,1) & Right(strRegProgram,1)
				If (strRegProgramsTmp = "{}") Then : bRegProgramsSkip = True : End If
			End If
			
			' Remove Patches
			objRegExp.IgnoreCase = True
			objRegExp.Pattern = "KB\d{6}"
			strRegProgramsTmp = objRegExp.Test(strRegProgram)
			If (strRegProgramsTmp) Then : bRegProgramsSkip = True : End If
			objRegExp.Pattern = "Q\d{6}"
			strRegProgramsTmp = objRegExp.Test(strRegProgram)
			If (strRegProgramsTmp) Then : bRegProgramsSkip = True : End If

			
			If Not (bRegProgramsSkip) Then
				objDbrRegPrograms.AddNew
				objDbrRegPrograms("DisplayName") = strRegProgramsDisplayName
				
				oReg.GetStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & strRegProgram, "DisplayVersion", strRegProgramsDisplayVersion
				objDbrRegPrograms("DisplayVersion") = Scrub(strRegProgramsDisplayVersion)
				
				
				objDbrRegPrograms.Update
			End If
		Next
		objDbrRegPrograms.Sort = "DisplayName"
	End If
	
	If (bRegLastUser) Then
		ReportProgress " Reading last user"
		oReg.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon","DefaultDomainName", strLastUserDomain
		oReg.GetStringValue HKEY_LOCAL_MACHINE,"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon","DefaultUserName", strLastUser
		If (Len(strLastUserDomain) > 0) Then
			strLastUser = strLastUserDomain & "\" & strLastUser
		End If
	End If

	
	
	ReportProgress "End subroutine: GatherRegInformation()"
	GatherRegInformation = True
End Function ' GatherRegInformation

Function GatherWMIInformation()
	If (bAllowErrors) Then
		On Error Resume Next
	End If
	ReportProgress "Start subroutine: GatherWMIInformation(" & strComputer & ")"
	Dim objSWbemLocator
	Dim arrGroupUser
	If (bAlternateCredentials) Then
		Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
		Set objWMIService = objSWbemLocator.ConnectServer(strComputer,"root\cimv2",strUserName,strPassword)
	Else
		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	End If
	If (Err <> 0) Then
	    ReportProgress Err.Number & " -- " &  Err.Description & " (" & strComputer & ")"
	    Err.Clear
	    GatherWMIInformation = False
	    Exit Function
	End If


	ReportProgress " Gathering OS information"
	Set colItems = objWMIService.ExecQuery("Select Name, CSDVersion, InstallDate, OSLanguage, Version, WindowsDirectory from Win32_OperatingSystem",,48)
	For Each objItem in colItems
		strOperatingSystem_InstallDate = objItem.InstallDate
		arrOperatingSystem_Name = Split(objItem.Name,"|")
		strOperatingSystem_Caption = arrOperatingSystem_Name(0)
		strOperatingSystem_ServicePack = objItem.CSDVersion
		strOperatingSystem_LanguageCode = Clng(objItem.OSLanguage)
		strOperatingSystem_LanguageCode = Hex(strOperatingSystem_LanguageCode)
		nOperatingSystemLevel = objItem.Version
		strOperatingSystem_WindowsDirectory = objItem.WindowsDirectory
	Next
	nOperatingSystemLevel = Mid(nOperatingSystemLevel,1,1) & Mid(nOperatingSystemLevel,3,1) ' 50 for Win2k 51 for XP
	
	
	If (bWMIBios) Then
		ReportProgress " Gathering BIOS information"
		Set colItems = objWMIService.ExecQuery("Select BiosCharacteristics, SMBIOSBIOSVersion, SMBIOSMajorVersion, SMBIOSMinorVersion, Version from Win32_BIOS",,48)
		For Each objItem in colItems
			strBIOS_SMBIOSBIOSVersion = objItem.SMBIOSBIOSVersion
			strBIOS_SMBIOSMajorVersion = objItem.SMBIOSMajorVersion
			strBIOS_SMBIOSMinorVersion = objItem.SMBIOSMinorVersion
			strBIOS_Version = objItem.Version
			arrBIOS_BiosCharacteristics(0) = 3
			If (IsArray(objItem.BiosCharacteristics)) Then
				For i = 0 To Ubound(objItem.BiosCharacteristics)
					ReDim Preserve arrBIOS_BiosCharacteristics(i)
					arrBIOS_BiosCharacteristics(i) = objItem.BiosCharacteristics(i)
				Next
			End If
		Next
	End If
	
	ReportProgress " Gathering computer system information"
	Set colItems = objWMIService.ExecQuery("Select Domain, DomainRole, Name, TotalPhysicalMemory from Win32_ComputerSystem",,48)
	For Each objItem in colItems
		strComputerSystem_Domain = objItem.Domain
		nComputerSystem_DomainRole = objItem.DomainRole
		strComputerSystem_Name = objItem.Name
		strComputerSystem_TotalPhysicalMemory = objItem.TotalPhysicalMemory
		strTotalPhysicalMemoryMB = Round(strComputerSystem_TotalPhysicalMemory / 1024 / 1024)
		Select Case nComputerSystem_DomainRole
			Case 0 
	            strComputerRole = "Standalone Workstation" : strDomainType = "workgroup"
	        Case 1        
	            strComputerRole = "Member Workstation" : strDomainType = "domain"
	        Case 2
	            strComputerRole = "Standalone Server" : strDomainType = "workgroup"
	        Case 3
	            strComputerRole = "Member Server" : strDomainType = "domain"
	        Case 4
	        	bWMILocalAccounts = False
	        	bWMILocalGroups = False
	            strComputerRole = "Domain Controller" : strDomainType = "domain"
	            bRoleDC = True
	        Case 5
	        	bWMILocalAccounts = False
	        	bWMILocalGroups = False
	            strComputerRole = "Domain Controller (PDC Emulator)" : strDomainType = "domain"
	            bRoleDC = True
		End Select
	Next

	ReportProgress " Gathering CD-ROM Information"
	Set colItems = objWMIService.ExecQuery("Select Drive, Manufacturer, Name from Win32_CDROMDrive",,48)
	Set objDbrCDROMDrive = CreateObject("ADOR.RecordSet")
	objDbrCDROMDrive.Fields.Append "Drive", adVarChar, MaxCharacters
	objDbrCDROMDrive.Fields.Append "Manufacturer", adVarChar, MaxCharacters
	objDbrCDROMDrive.Fields.Append "Name", adVarChar, MaxCharacters
	objDbrCDROMDrive.Open
	For Each objItem in colItems
		objDbrCDROMDrive.AddNew
		objDbrCDROMDrive("Drive") = objItem.Drive
	    objDbrCDROMDrive("Manufacturer") = objItem.Manufacturer
	    objDbrCDROMDrive("Name") = objItem.Name
	    objDbrCDROMDrive.Update
	Next
	objDbrCDROMDrive.Sort = "Drive"
	
	
	
	ReportProgress " Gathering computer system product information"
	Set colItems = objWMIService.ExecQuery("Select Vendor, Name, IdentifyingNumber from Win32_ComputerSystemProduct",,48)
	For Each objItem in colItems
	    strComputerSystemProduct_Manufacturer = objItem.Vendor
	    strComputerSystemProduct_Name = objItem.Name
	    strComputerSystemProduct_IdentifyingNumber = objItem.IdentifyingNumber
	Next

	ReportProgress " Gathering disk information"
	Set colItems = objWMIService.ExecQuery ("SELECT Caption, DeviceID, Interfacetype, Size FROM Win32_DiskDrive")
	Set objDbrDrives = CreateObject("ADOR.Recordset")
	objDbrDrives.Fields.Append "Caption", adVarChar, MaxCharacters
	objDbrDrives.Fields.Append "DeviceID", adVarChar, MaxCharacters
	objDbrDrives.Fields.Append "InterfaceType", adVarChar, MaxCharacters
	objDbrDrives.Fields.Append "Size", adVarChar, MaxCharacters
	objDbrDrives.Open
	Set objDbrDisks = CreateObject("ADOR.Recordset")
	objDbrDisks.Fields.Append "Caption", adVarChar, MaxCharacters
	objDbrDisks.Fields.Append "Size", adVarChar, MaxCharacters 
	objDbrDisks.Fields.Append "FileSystem", adVarChar, MaxCharacters 
	objDbrDisks.Fields.Append "FreeSpace", adVarChar, MaxCharacters 
	objDbrDisks.Fields.Append "VolumeName", adVarChar, MaxCharacters 
	objDbrDisks.Fields.Append "ParentDriveID", adVarChar, MaxCharacters 
	objDbrDisks.Open
	For Each objItem In colItems
		objDbrDrives.AddNew
		objDbrDrives("Caption") = objItem.Caption
		objDbrDrives("DeviceID") = objItem.DeviceID
		objDbrDrives("InterfaceType") = Scrub(objItem.InterfaceType)
		If (IsNull(objItem.Size)) Then
			objDbrDrives("Size") = ""
		Else
			objDbrDrives("Size") = objItem.Size
		End If
	
		Set objDiskPartitions = objItem.Associators_("Win32_DiskDriveToDiskPartition", "Win32_DiskPartition") 
		For Each objDiskPartition In objDiskPartitions
			Set objLogicalDisks = objDiskPartition.Associators_("Win32_LogicalDiskToPartition", "Win32_LogicalDisk")
			For Each objLogicalDisk In objLogicalDisks
				objDbrDisks.AddNew
				objDbrDisks("Caption") = objLogicalDisk.Caption
				objDbrDisks("Size") = objLogicalDisk.Size
				objDbrDisks("FileSystem") = objLogicalDisk.FileSystem
				objDbrDisks("FreeSpace") = objLogicalDisk.FreeSpace
				objDbrDisks("VolumeName") = objLogicalDisk.VolumeName
				objDbrDisks("ParentDriveID") = objItem.DeviceID
				objDbrDisks.Update
			Next
		Next
	
		objDbrDrives.Update
	Next

	If (bWMILocalGroups) Then
		ReportProgress " Gathering local groups"
		Set colItems = objWMIService.ExecQuery("Select Name from Win32_Group Where Domain='" & strComputerSystem_Name & "'",,48)
		Set objDbrLocalGroups = CreateObject("ADOR.Recordset")
		objDbrLocalGroups.Fields.Append "Name", adVarChar, MaxCharacters
		objDbrLocalGroups.Open
		For Each objItem in colItems
			objDbrLocalGroups.AddNew
			objDbrLocalGroups("Name") = objItem.Name
			objDbrLocalGroups.Update
		Next
		objDbrLocalGroups.Sort = "Name"

		If Not (objDbrLocalGroups.Bof) Then
			objDbrLocalGroups.Movefirst
			ReportProgress " Gathering local group members"
		End If

		Set objDbrGroupUser = CreateObject("ADOR.Recordset")
		objDbrGroupUser.Fields.Append "Groupname", adVarChar, MaxCharacters
		objDbrGroupUser.Fields.Append "Member", adVarChar, MaxCharacters
		objDbrGroupUser.Open
		Do Until objDbrLocalGroups.Eof
			Set colItems = objWMIService.ExecQuery("Select * from Win32_GroupUser where GroupComponent=""Win32_Group.Domain='" & strComputerSystem_Name & "',Name='" & objDbrLocalGroups.Fields.Item("Name") & "'""",,48)
			For Each objItem in colItems
				objDbrGroupUser.AddNew
				objDbrGroupUser("Groupname") = objDbrLocalGroups.Fields.Item("Name")
				arrGroupUser = Split(objItem.PartComponent, """")
				objDbrGroupUser("Member") = arrGroupUser(1) & "\" & arrGroupUser(3)
				objDbrGroupUser.Update
			Next
			objDbrLocalGroups.Movenext
		Loop
	End If

	If (bWMIIP4Routes And nOperatingSystemLevel > 50) Then
		ReportProgress " Gathering IP Route information"
		Set colItems = objWMIService.ExecQuery("Select Destination, Mask, NextHop from Win32_IP4RouteTable",,48)
		Set objDbrIP4RouteTable = CreateObject("ADOR.Recordset")
		objDbrIP4RouteTable.Fields.Append "Destination", adVarChar, MaxCharacters
		objDbrIP4RouteTable.Fields.Append "Mask", adVarChar, MaxCharacters
		objDbrIP4RouteTable.Fields.Append "NextHop", adVarChar, MaxCharacters
		objDbrIP4RouteTable.Open
		For Each objItem in colItems
			objDbrIP4RouteTable.AddNew
			objDbrIP4RouteTable("Destination") = objItem.Destination
			objDbrIP4RouteTable("Mask") = objItem.Destination
			objDbrIP4RouteTable("NextHop") = objItem.NextHop
			objDbrIP4RouteTable.Update
		Next
	Else
		bWMIIP4Routes = False
	End If

	ReportProgress " Gathering network adapter configuration"
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True",,48)
	i = -1
	For Each objItem in colItems
	    i = i + 1
		ReDim Preserve arrNetadapter_Description(i), arrNetadapter_MACAddress(i)
		ReDim Preserve arrNetadapter_DNSHostName(i)
		ReDim Preserve arrNetadapter_DHCPEnabled(i)
		ReDim Preserve arrNetadapter_DHCPServer(i)
		ReDim Preserve arrNetadapter_DNS(i)
		ReDim Preserve arrNetadapter_WINSPrimaryServer(i)
		ReDim Preserve arrNetadapter_WINSSecondaryServer(i)
	    arrNetadapter_Description(i) = objItem.Description
	    arrNetadapter_MACAddress(i) = objItem.MACAddress
	    arrNetadapter_DNSHostName(i) = objItem.DNSHostName
		arrNetadapter_DHCPEnabled(i) = objItem.DHCPEnabled
	    arrNetadapter_DHCPServer(i) = objItem.DHCPServer
	    arrNetadapter_DNS(i) = objItem.DNSDomain
	    arrNetadapter_WINSPrimaryServer(i) = objItem.WINSPrimaryServer
	    arrNetadapter_WINSSecondaryServer(i) = objItem.WINSSecondaryServer
		If Not IsNull(objItem.IPAddress) Then
			For j = 0 To UBound(objItem.IPAddress)
				If (j > UBound(arrNetadapter_IPAddress,2)) Then
					ReDim Preserve arrNetadapter_IPAddress(nMaxNetworkAdapters,j)
				End If
				arrNetadapter_IPAddress(i,j) = objItem.IPAddress(j)
			Next
		End If
		If Not IsNull(objItem.IPSubnet) Then
			For j = 0 To UBound(objItem.IPSubnet)
				If (j > UBound(arrNetadapter_IPSubnet,2)) Then
					ReDim Preserve arrNetadapter_IPSubnet(nMaxNetworkAdapters,j)
				End If
				arrNetadapter_IPSubnet(i,j) = objItem.IPSubnet(j)
			Next
		End If
		If Not IsNull(objItem.DefaultIPGateway) Then
			For j = 0 To UBound(objItem.DefaultIPGateway)
				If (j > UBound(arrNetadapter_DefaultIPGateway,2)) Then
					ReDim Preserve arrNetadapter_DefaultIPGateway(nMaxNetworkAdapters,j)
				End If
				arrNetadapter_DefaultIPGateway(i,j) = objItem.DefaultIPGateway(j)
			Next
		End If
		If Not IsNull(objItem.DNSServerSearchOrder) Then
			For j = 0 To UBound(objItem.DNSServerSearchOrder)
				If (j > UBound(arrNetadapter_DNSServerSearchOrder,2)) Then
					ReDim Preserve arrNetadapter_DNSServerSearchOrder(nMaxNetworkAdapters,j)
				End If
				arrNetadapter_DNSServerSearchOrder(i,j) = objItem.DNSServerSearchOrder(j)
			Next
		End If
	Next

	If (bWMIEventLogFile) Then
		ReportProgress " Gathering event log settings"
		Set colItems = objWMIService.ExecQuery("Select LogFileName, MaxFileSize, Name, OverwritePolicy from Win32_NTEventLogFile",,48)
		Set objDbrEventLogFile = CreateObject("ADOR.RecordSet")
		objDbrEventLogFile.Fields.Append "LogFileName", adVarChar, MaxCharacters
		objDbrEventLogFile.Fields.Append "MaxFileSize", adVarChar, MaxCharacters
		objDbrEventLogFile.Fields.Append "Name", adVarChar, MaxCharacters
		objDbrEventLogFile.Fields.Append "OverwritePolicy", adVarChar, MaxCharacters
		objDbrEventLogFile.Open
		For Each objItem in colItems
			objDbrEventLogFile.AddNew
			objDbrEventLogFile("LogFileName") = objItem.LogFileName
		    objDbrEventLogFile("MaxFileSize") = objItem.MaxFileSize
		    objDbrEventLogFile("Name") = objItem.Name
		    objDbrEventLogFile("OverwritePolicy") = objItem.OverwritePolicy
		    objDbrEventLogFile.Update
		Next
		objDbrEventLogFile.Sort = "LogFileName"
	End If
	

	ReportProgress " Gathering pagefile information"
	Set colItems = objWMIService.ExecQuery("Select Drive, InitialSize, MaximumSize from Win32_PageFile",,48)
	Set objDbrPagefile = CreateObject("ADOR.RecordSet")
	objDbrPagefile.Fields.Append "Drive", adVarChar, MaxCharacters
	objDbrPagefile.Fields.Append "InitialSize", adVarChar, MaxCharacters
	objDbrPagefile.Fields.Append "MaximumSize", adVarChar, MaxCharacters
	objDbrPagefile.Open
	For Each objItem in colItems
		objDbrPagefile.AddNew
		objDbrPagefile("Drive") = objItem.Drive
		objDbrPagefile("InitialSize") = objItem.InitialSize
		objDbrPagefile("MaximumSize") = objItem.MaximumSize
		objDbrPagefile.Update
	Next

	ReportProgress " Gathering information about physical memory"
	Set colItems = objWMIService.ExecQuery("Select BankLabel, Capacity, FormFactor, MemoryType from Win32_PhysicalMemory",,48)
	Set objDbrPhysicalMemory = CreateObject("ADOR.RecordSet")
	objDbrPhysicalMemory.Fields.Append "BankLabel", adVarChar, MaxCharacters
	objDbrPhysicalMemory.Fields.Append "Capacity", adVarChar, MaxCharacters
	objDbrPhysicalMemory.Fields.Append "FormFactor", adVarChar, MaxCharacters
	objDbrPhysicalMemory.Fields.Append "MemoryType", adVarChar, MaxCharacters
	objDbrPhysicalMemory.Open
	For Each objItem In colItems
		objDbrPhysicalMemory.AddNew
		objDbrPhysicalMemory("BankLabel") = objItem.BankLabel
		objDbrPhysicalMemory("Capacity") = objItem.Capacity
		objDbrPhysicalMemory("FormFactor") = objItem.FormFactor
		objDbrPhysicalMemory("MemoryType") = objItem.MemoryType
		objDbrPhysicalMemory.Update
	Next
	objDbrPhysicalMemory.Sort = "BankLabel"
	
	If (bWMIPrinters) Then
		ReportProgress " Gathering printer information"
		Set colItems = objWMIService.ExecQuery("Select DriverName, Name, PortName from Win32_Printer Where ServerName = Null",,48)
		Set objDbrPrinters = CreateObject("ADOR.Recordset")
		objDbrPrinters.Fields.Append "DriverName", adVarChar, MaxCharacters
		objDbrPrinters.Fields.Append "Name", adVarChar, MaxCharacters
		objDbrPrinters.Fields.Append "PortName", adVarChar, MaxCharacters
		objDbrPrinters.Open
		For Each objItem In colItems
			objDbrPrinters.AddNew
			objDbrPrinters("Drivername") = objItem.DriverName
			objDbrPrinters("Name") = objItem.Name
			objDbrPrinters("Portname") = objItem.PortName
			objDbrPrinters.Update
		Next
		objDbrPrinters.Sort = "Name"
	End If

	If (bWMIRunningProcesses) Then
		ReportProgress " Gathering process information"
		Set colItems = objWMIService.ExecQuery("Select Caption, ExecutablePath from Win32_Process",,48)
		Set objDbrProcess = CreateObject("ADOR.Recordset")
		objDbrProcess.Fields.Append "Caption", adVarChar, MaxCharacters
		objDbrProcess.Fields.Append "ExecutablePath", adVarChar, MaxCharacters
		objDbrProcess.Open
		For Each objItem In colItems
			objDbrProcess.AddNew
			objDbrProcess("Caption") = Scrub(objItem.Caption)
			objDbrProcess("ExecutablePath") = Scrub(objItem.ExecutablePath)
			objDbrProcess.Update
		Next
		objDbrProcess.Sort = "Caption"
	
	End If

	ReportProgress " Gathering processor information"
	Set colItems = objWMIService.ExecQuery("Select Description, ExtClock, L2CacheSize, Name, MaxClockSpeed, SocketDesignation from Win32_Processor",,48)
	Set objDbrProcessorSockets = CreateObject("ADOR.Recordset")
	objDbrProcessorSockets.Fields.Append "SocketDesignation", adVarChar, MaxCharacters
	objDbrProcessorSockets.Fields.Append "Count", adVarChar, MaxCharacters
	objDbrProcessorSockets.Open

	i = 0
	For Each objItem in colItems
		i = i + 1
		strProcessor_L2CacheSize = objItem.L2CacheSize
		strProcessor_ExtClock = objItem.ExtClock
		strProcessor_Name = objItem.Name
		strProcessor_Description = objItem.Description
		strProcessor_MaxClockSpeed = objItem.MaxClockSpeed
		CalculateProcessorSockets Scrub(objItem.SocketDesignation)
	Next
	' Remove filter
	objDbrProcessorSockets.Filter = " Count > 0 "
	intProcessors = objDbrProcessorSockets.Recordcount
	If (i > objDbrProcessorSockets.Recordcount) Then
		bProcessorHTSystem = True
	Else
		bProcessorHTSystem = False
	End If
	
	Err.Clear
	If (bWMIApplications) Then
		ReportProgress " Gathering application information"
		Set colItems = objWMIService.ExecQuery("Select Name, Vendor, Version, InstallDate from Win32_Product WHERE Name <> Null",,48)
		Set objDbrProducts = CreateObject("ADOR.Recordset")
		objDbrProducts.Fields.Append "ProductName", adVarChar, MaxCharacters
		objDbrProducts.Fields.Append "Vendor", adVarChar, MaxCharacters
		objDbrProducts.Fields.Append "Version", adVarChar, MaxCharacters
		objDbrProducts.Fields.Append "InstallDate", adVarChar, MaxCharacters
		objDbrProducts.Open
		For Each objItem In colItems
			If (Err <> 0) Then
			    ReportProgress " " & Err.Number & " -- " &  Err.Description & " (" & strComputer & ")"
			    ReportProgress " Win32_Product class is not installed (Windows Installer Applications will not appear)"
			    ReportProgress " You can add it with Add/Remove Windows Components -> Management and Monitoring -> WMI Windows Installer Provider"
			    Err.Clear
			    errWin32_Product = False
			    Exit For
			Else
				errWin32_Product = True
			End If
	
			objDbrProducts.AddNew
			objDbrProducts("ProductName") = objItem.Name
			If (IsNull(objItem.Vendor)) Then
				objDbrProducts("Vendor") = ""
			Else
				objDbrProducts("Vendor") = objItem.Vendor
			End If
			objDbrProducts("Version") = Scrub(objItem.Version)
			If (IsNull(objItem.InstallDate)) Then
				objDbrProducts("InstallDate") = "N/A"
			Else
				objDbrProducts("InstallDate") = objItem.InstallDate
			End If
			objDbrProducts.Update
		Next
		objDbrProducts.Sort = "ProductName"
	End If
	
	If (bWMIRegistry) Then
		ReportProgress " Gathering registry size information"
		Set colItems = objWMIService.ExecQuery("Select CurrentSize, MaximumSize from Win32_Registry",,48)
		For Each objItem In colItems
			nCurrentSize = objItem.CurrentSize
			nMaximumSize = objItem.MaximumSize
		Next
	End If

	If (bWMIServices) Then 
		ReportProgress " Gathering information about services"
		Set colItems = objWMIService.ExecQuery("Select Caption, Started, StartMode, StartName from Win32_Service Where ServiceType ='Share Process' Or ServiceType ='Own Process'",,48)
		Set objDbrServices = CreateObject("ADOR.Recordset")
		objDbrServices.Fields.Append "Caption", adVarChar, MaxCharacters
		objDbrServices.Fields.Append "Started", adVarChar, MaxCharacters
		objDbrServices.Fields.Append "StartMode", adVarChar, MaxCharacters
		objDbrServices.Fields.Append "StartName", adVarChar, MaxCharacters
		objDbrServices.Open
		For Each objItem In colItems
			objDbrServices.AddNew
			objDbrServices("Caption") = objItem.Caption
			objDbrServices("Started") = objItem.Started
			objDbrServices("StartMode") = objItem.StartMode 
			objDbrServices("StartName") = objItem.StartName
			If (LCase(objItem.Caption) = "mssqlserver") Then
				bRoleSQL = True
			End If
			objDbrServices.Update
		Next
		objDbrServices.Sort = "Caption"
	End If
	
	If (bWMIFileShares) Then
		ReportProgress " Gathering information about shares"
		Set colItems = objWMIService.ExecQuery("Select Name, Description, Path, Type from Win32_Share",,48)
		Set objDbrShares = CreateObject("ADOR.Recordset")
		objDbrShares.Fields.Append "Name", adVarChar, MaxCharacters
		objDbrShares.Fields.Append "Description", adVarChar, MaxCharacters
		objDbrShares.Fields.Append "Path", adVarChar, MaxCharacters
		objDbrShares.Open
		For Each objItem in colItems
			objDbrShares.AddNew
			objDbrShares("Name") = objItem.Name
			objDbrShares("Description") = objItem.Description
			objDbrShares("Path") = objItem.Path
		    objDbrShares.Update
		    If (objItem.Type = 0) Then
		    	bRoleFile = True
		    End If
		    If (objItem.Type = 1) Then
		    	bRolePrint = True
		    End If
		Next
		objDbrShares.Sort = "Name"
	End If

	If (bWMIHardware) Then
		ReportProgress " Gathering Sound Device information"
		Set colItems = objWMIService.ExecQuery("Select Name, Manufacturer from Win32_SoundDevice",,48)
		Set objDbrSoundDevice = CreateObject("ADOR.Recordset")
		objDbrSoundDevice.Fields.Append "Name", adVarChar, MaxCharacters
		objDbrSoundDevice.Fields.Append "Manufacturer", adVarChar, MaxCharacters
		objDbrSoundDevice.Open
		For Each objItem in colItems
			objDbrSoundDevice.AddNew
			objDbrSoundDevice("Name") = objItem.Name
			objDbrSoundDevice("Manufacturer") = objItem.Manufacturer
			objDbrSoundDevice.Update
		Next
	End If
	
	If (bWMIStartupCommands) Then
		ReportProgress " Gathering Startup Commands information"
		Set colItems = objWMIService.ExecQuery("Select Command, Name, User from Win32_StartupCommand",,48)
		Set objDbrStartupCommand = CreateObject("ADOR.Recordset")
		objDbrStartupCommand.Fields.Append "Name", adVarChar, MaxCharacters
		objDbrStartupCommand.Fields.Append "Command", adVarChar, MaxCharacters
		objDbrStartupCommand.Fields.Append "User", adVarChar, MaxCharacters
		objDbrStartupCommand.Open
		For Each objItem in colItems
			objDbrStartupCommand.AddNew
			objDbrStartupCommand("Name") = objItem.Name
			objDbrStartupCommand("Command") = objItem.Command
			objDbrStartupCommand("User") = objItem.User
			objDbrStartupCommand.Update
		Next
		objDbrStartupCommand.Sort = "User, Name"
	End If
	
	
	
	ReportProgress " Gathering system enclosure information"
	Set colItems = objWMIService.ExecQuery("Select ChassisTypes from Win32_SystemEnclosure",,48)
	For Each objItem in colItems
	    For i = Lbound(objItem.ChassisTypes) to Ubound(objItem.ChassisTypes)
	        nChassisType = objItem.ChassisTypes(i)
	    Next
		Select Case nChassisType
	        Case 1        
	            strChassisType = "Other" 
	        Case 2
	            strChassisType = "Unknown" 
	        Case 3
	            strChassisType = "Desktop"
	        Case 4
	            strChassisType = "Low-profile desktop"
	        Case 5
	            strChassisType = "Pizza box"
	        Case 6
	            strChassisType = "Mini tower"
	        Case 7
	            strChassisType = "Tower"
	        Case 8
	            strChassisType = "Portable"
	        Case 9
	            strChassisType = "Laptop"
	        Case 10
	            strChassisType = "Notebook"
	        Case 11
	            strChassisType = "Hand-held"
	        Case 12
	            strChassisType = "Docking station"
	        Case 13
	            strChassisType = "All-in-one"
	        Case 14
	            strChassisType = "Subnotebook"
	        Case 15
	            strChassisType = "Space-saving"
	        Case 16
	            strChassisType = "Lunch box"
	        Case 17
	            strChassisType = "Main system chassis"
	        Case 18
	            strChassisType = "Expansion chassis"
	        Case 19
	            strChassisType = "Subchassis"
	        Case 20
	            strChassisType = "Bus-expansion chassis"
	        Case 21
	            strChassisType = "Peripheral chassis"
	        Case 22
	            strChassisType = "Storage chassis"
	        Case 23
	            strChassisType = "Rack Mount chassis"
	        Case 24
	            strChassisType = "Sealed-case computer"
		End Select
	Next

	If (bWMIHardware) Then
		ReportProgress " Gathering TapeDrive information"
		Set colItems = objWMIService.ExecQuery("Select Name, Description, Manufacturer from Win32_TapeDrive",,48)
		Set objDbrTapeDrive = CreateObject("ADOR.Recordset")
		objDbrTapeDrive.Fields.Append "Name", adVarChar, MaxCharacters
		objDbrTapeDrive.Fields.Append "Description", adVarChar, MaxCharacters
		objDbrTapeDrive.Fields.Append "Manufacturer", adVarChar, MaxCharacters
		objDbrTapeDrive.Open
		For Each objItem in colItems
			objDbrTapeDrive.AddNew
			bHasTapeDrive = True
			objDbrTapeDrive("Name") = objItem.Name
			objDbrTapeDrive("Description") = objItem.Description
			objDbrTapeDrive("Manufacturer") = objItem.Manufacturer
			objDbrTapeDrive.Update
		Next
	End If	
	
	ReportProgress " Gathering time zone"
	Set colItems = objWMIService.ExecQuery("Select Description from Win32_TimeZone",,48)
	For Each objItem in colItems
		strTimeZone = objItem.Description
	Next
	
	If (bWMIPatches) Then
		ReportProgress " Gathering information about patches"
		Set colItems = objWMIService.ExecQuery("Select Description, HotFixID, InstalledOn from Win32_QuickFixEngineering Where HotfixID <> 'File 1' And HotfixID <> 'Q147222'",,48)
		Set objDbrPatches = CreateObject("ADOR.Recordset")
		objDbrPatches.Fields.Append "Description", adVarChar, MaxCharacters
		objDbrPatches.Fields.Append "HotfixID", adVarChar, MaxCharacters
		objDbrPatches.Fields.Append "InstallDate", adVarChar, MaxCharacters
		objDbrPatches.Open
		For Each objItem in colItems
			objDbrPatches.AddNew
			objDbrPatches("Description") = objItem.Description
			objDbrPatches("HotfixID") = objItem.HotfixID
			If (IsNull(objItem.InstalledOn) Or objItem.InstalledOn = "") Then
				objDbrPatches("InstallDate") = "N/A"
			Else
				objDbrPatches("InstallDate") = objItem.InstalledOn
			End If
			objDbrPatches.Update
		Next
	End If

	If (bWMILocalAccounts) Then
		ReportProgress " Gathering local users"
		Set colItems = objWMIService.ExecQuery("Select Name from Win32_UserAccount Where Domain='" & strComputerSystem_Name & "'",,48)
		Set objDbrLocalAccounts = CreateObject("ADOR.Recordset")
		objDbrLocalAccounts.Fields.Append "UserName", adVarChar, MaxCharacters
		objDbrLocalAccounts.Open
		For Each objItem in colItems
			objDbrLocalAccounts.AddNew
			objDbrLocalAccounts("UserName") = objItem.Name
			objDbrLocalAccounts.Update
		Next
		objDbrLocalAccounts.Sort = "UserName"
	End If

	If (bWMIHardware) Then
		ReportProgress " Gathering Video Controller information"
		Set colItems = objWMIService.ExecQuery("Select AdapterCompatibility, AdapterRAM, Name from Win32_VideoController",,48)
		Set objDbrVideoController = CreateObject("ADOR.Recordset")
		objDbrVideoController.Fields.Append "AdapterCompatibility", adVarChar, MaxCharacters
		objDbrVideoController.Fields.Append "AdapterRAM", adVarChar, MaxCharacters
		objDbrVideoController.Fields.Append "Name", adVarChar, MaxCharacters
		objDbrVideoController.Open
		For Each objItem in colItems
			objDbrVideoController.AddNew
			objDbrVideoController("AdapterCompatibility") = objItem.AdapterCompatibility
			objDbrVideoController("AdapterRAM") = objItem.AdapterRAM
			objDbrVideoController("Name") = objItem.Name
			objDbrVideoController.Update
		Next
	End If
	
	Set objWMIService = Nothing
	ReportProgress "End subroutine: GatherWMIInformation()"
	GatherWMIInformation = True
End Function ' GatherWMIInformation


Sub GetOptions()
	Dim objArgs, nArgs
	' Default settings
	bWMIBios = True
	bWMIRegistry = True
	bWMIApplications = True
	bWMIPatches = True
	bWMIEventLogFile = True
	bWMIFileShares = True
	bWMIIP4Routes = True
	bWMILocalAccounts = True
	bWMILocalGroups = True
	bWMIServices = True
	bWMIStartupCommands = True
	bWMIPrinters = True
	bWMIRunningProcesses = True
	bWMIHardware = True
	bHasTapeDrive = False
	bRegDomainSuffix = True
	bRegPrintSpoolLocation = True
	bRegPrograms = True
	bRegWindowsComponents = True
	bRegLastUser = True
	bDoRegistryCheck = True
	strComputer = ""
	bAlternateCredentials = False
	bInvalidArgument = False
	bDisplayHelp = False
	bShowWord = True
	bWordExtras = True 
	nBaseFontSize = 12
	bUseSpecificTable = False
	bUseDOTFile = False
	bSaveFile = True
	bCheckVersion = False
	strExportFormat = "word"
	strStylesheet = ""
	bAllowErrors = True
	Set objArgs = WScript.Arguments
	If (objArgs.Count > 0) Then
		For nArgs = 0 To objArgs.Count - 1
			SetOptions objArgs(nArgs)
		Next
	Else
		WScript.Echo "For help type: cscript.exe sydi-server.vbs -h"
	End If
	SystemRolesDefine
	If (bSaveFile = False And strExportFormat = "xml") Then
		bInvalidArgument = True
	End If
End Sub ' GetOptions

Sub GetWMIProviderList
	If (bAllowErrors) Then
		On Error Resume Next
	End If
	Dim colNameSpaces
	ReportProgress vbCrlf & "Checking for Other WMI Providers"
	Dim objSWbemLocator
	If (bAlternateCredentials) Then
		Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
		Set objWMIService = objSWbemLocator.ConnectServer(strComputer,"root",strUserName,strPassword)
	Else
		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root")
	End If
	If (Err <> 0) Then
	    ReportProgress Err.Number & " -- " &  Err.Description & " (" & strComputer & ")"
	    Err.Clear
	    Exit Sub
	End If
	Set colNameSpaces = objWMIService.InstancesOf("__NAMESPACE")
	For Each objItem In colNameSpaces
		Select Case objItem.Name
			Case "MicrosoftIISv2"
				bHasMicrosoftIISv2 = True
				ReportProgress " Found MicrosoftIISv2 (Internet Information Services)"
		End Select
	Next	
End Sub ' GetWMIProviderList

Sub PopulateWordfile()
	If (bAllowErrors) Then
		On Error Resume Next
	End If
	' WdListNumberStyle
	Const wdListNumberStyleArabic = 0
	Const wdListNumberStyleUppercaseRoman = 1
	Const wdListNumberStyleLowercaseRoman = 2
	Const wdListNumberStyleUppercaseLetter = 3
	Const wdListNumberStyleLowercaseLetter = 4
	Const wdListNumberStyleOrdinal = 5
	Const wdListNumberStyleCardinalText = 6
	Const wdListNumberStyleOrdinalText = 7
	Const wdListNumberStyleArabicLZ = 22
	Const wdListNumberStyleBullet = 23
	Const wdListNumberStyleLegal = 253
	Const wdListNumberStyleLegalLZ = 254
	Const wdListNumberStyleNone = 255
	
	' WdListGalleryType
	Const wdBulletGallery = 1
	Const wdNumberGallery = 2
	Const wdOutlineNumberGallery = 3
	' WdBreakType
	Const wdPageBreak = 7
	' WdBuiltInProperty
	Const wdPropertyAuthor = 3
	Const wdPropertyComments = 5
	' WdBuiltInStyle
	Const wdStyleBodyText = -67
	Const wdStyleFooter = -33
	Const wdStyleHeader = -32
	Const wdStyleHeading1 = -2
	Const wdStyleHeading2 = -3
	Const wdStyleHeading3 = -4
	Const wdStyleHeading4 = -5
	Const wdStyleTitle = -63
	Const wdStyleTOC1 = -20
	Const wdStyleTOC2 = -21
	Const wdStyleTOC3 = -22
	' WdFieldType
	'Const wdFieldEmpty = -1
	Const wdFieldNumPages = 26
	Const wdFieldPage = 33
	' WdParagraphAlignment
	Const wdAlignParagraphRight = 2
	' WdSeekView
	Const wdSeekMainDocument = 0
	Const wdSeekCurrentPageHeader = 9
	Const wdSeekCurrentPageFooter = 10
	' Page Viewing
	Const wdPaneNone = 0
	Const wdPrintView = 3

	
	ReportProgress VbCrLf & "Start subroutine: PopulateWordfile()"
	Set oWord = CreateObject("Word.Application")
	If (Err <> 0) Then
	    ReportProgress Err.Number & " -- " &  Err.Description & " (" & strComputer & ")"
	    ReportProgress " Could not open Microsoft Word, verify that it is correctly installed on the computer you are scanning from."
	    Err.Clear
	    Exit Sub
	End If

	'oWord.Activate
	
	If (bUseDOTFile) Then
		oWord.Documents.Add strDOTFile
		If (Err <> 0) Then
		    ReportProgress Err.Number & " -- " &  Err.Description & " (" & strDOTFile & ")"
		    ReportProgress " Unable to open the template file " & strDOTFile
		    ReportProgress " Did you use the correct path?"
		    Err.Clear
		    Exit Sub
		End If
	Else
		oWord.Documents.Add
	End If
	oWord.Application.Visible = bShowWord
	ReportProgress " Opening Empty document"
	Set oListTemplate = oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1)
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(1).Numberformat = "%1."
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(1).NumberStyle = wdListNumberStyleArabic
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(2).Numberformat = "%1.%2."
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(2).NumberStyle = wdListNumberStyleArabic
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(3).Numberformat = "%1.%2.%3."
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(3).NumberStyle = wdListNumberStyleArabic
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(4).Numberformat = "%1.%2.%3.%4."
	oWord.ListGalleries(wdOutlineNumberGallery).ListTemplates(1).listlevels(4).NumberStyle = wdListNumberStyleArabic

	If Not (bUseDOTFile) Then
		oWord.ActiveDocument.Styles(wdStyleTOC1).Font.Bold = True
		oWord.ActiveDocument.Styles(wdStyleBodyText).Font.Name = strFontBodyText
		oWord.ActiveDocument.Styles(wdStyleBodyText).Font.Size = nBaseFontSize
		oWord.ActiveDocument.Styles(wdStyleHeading1).Font.Name = strFontHeading1
		oWord.ActiveDocument.Styles(wdStyleHeading1).Font.Size = (nBaseFontSize + 4)
		oWord.ActiveDocument.Styles(wdStyleHeading2).Font.Name = strFontHeading2
		oWord.ActiveDocument.Styles(wdStyleHeading2).Font.Size = (nBaseFontSize + 2)
		oWord.ActiveDocument.Styles(wdStyleHeading3).Font.Name = strFontHeading3
		oWord.ActiveDocument.Styles(wdStyleHeading3).Font.Size = (nBaseFontSize + 1)
		oWord.ActiveDocument.Styles(wdStyleHeading4).Font.Name = strFontHeading4
		oWord.ActiveDocument.Styles(wdStyleHeading4).Font.Size = nBaseFontSize
		oWord.ActiveDocument.Styles(wdStyleTitle).Font.Name = strFontTitle
		oWord.ActiveDocument.Styles(wdStyleTitle).Font.Size = (nBaseFontSize + 4)
		oWord.ActiveDocument.Styles(wdStyleTOC1).Font.Name = strFontTOC1
		oWord.ActiveDocument.Styles(wdStyleTOC1).Font.Size = nBaseFontSize
		oWord.ActiveDocument.Styles(wdStyleTOC2).Font.Name = strFontTOC2
		oWord.ActiveDocument.Styles(wdStyleTOC2).Font.Size = nBaseFontSize
		oWord.ActiveDocument.Styles(wdStyleTOC3).Font.Name = strFontTOC3
		oWord.ActiveDocument.Styles(wdStyleTOC3).Font.Size = nBaseFontSize
		oWord.ActiveDocument.Styles(wdStyleHeader).Font.Name = strFontHeader
		oWord.ActiveDocument.Styles(wdStyleHeader).Font.Size = (nBaseFontSize - 1)
		oWord.ActiveDocument.Styles(wdStyleFooter).Font.Name = strFontFooter
		oWord.ActiveDocument.Styles(wdStyleFooter).Font.Size = (nBaseFontSize - 1)
		ReportProgress " Setting styles"	
	End If

	oWord.Selection.Style = wdStyleTitle
	oWord.Selection.TypeText "Basic documentation For " & strComputerSystem_Name & VbCrLf & VbCrLf

	If (strDocumentAuthor = "") Then
		strDocumentAuthor = oWord.ActiveDocument.BuiltInDocumentProperties(wdPropertyAuthor).Value
	End If
	oWord.Selection.Style = wdStyleBodyText
	oWord.Selection.TypeText "Document versions:" & vbCrLf & "Version 1.0" & vbTab & Date & vbTab & strDocumentAuthor & vbTab & vbCrLf & vbCrLf
	
	oWord.Selection.Font.Bold = True
	oWord.Selection.TypeText "SUMMARY" & vbCrLf
	oWord.Selection.Font.Bold = False
	
	oWord.Selection.Style = wdStyleBodyText
	If (bWordExtras) Then 
		oWord.Selection.TypeText ""
	End If
	oWord.Selection.TypeText "The system is running " & strOperatingSystem_Caption & " " & strOperatingSystem_ServicePack & VbCrLf

	If (bWordExtras) Then 
		oWord.Selection.TypeText "System Owner: "
		oWord.Selection.TypeText "[provide name and title]" & VbCrLf
	End If

	If (bRegDomainSuffix) Then
		oWord.Selection.TypeText "FQDN: " &  LCase(strComputerSystem_Name) & "." & strPrimaryDomain & VbCrLf
	End If
	oWord.Selection.TypeText "NetBIOS: " & strComputerSystem_Name & VbCrLf
	oWord.Selection.TypeText "Roles: "

	i = 0
	If Not (objDbrSystemRoles.Bof) Then
		objDbrSystemRoles.MoveFirst
	End If
	Do Until objDbrSystemRoles.EOF
		If (i = 0) Then
			oWord.Selection.TypeText Cstr(objDbrSystemRoles.Fields.Item("Role"))
		Else
			oWord.Selection.TypeText ", " & Cstr(objDbrSystemRoles.Fields.Item("Role"))
		End If
		i = i + 1
		objDbrSystemRoles.MoveNext
	Loop
	If (i = 0) Then
		oWord.Selection.TypeText "[provide the roles of this system]" & vbCrLf
	Else
		oWord.Selection.TypeText vbCrLf
	End If
	If (bWordExtras) Then 
		oWord.Selection.TypeText "Physical location: "
		oWord.Selection.TypeText "[Austin]" & vbCrLf
		oWord.Selection.TypeText "Logical location: " 
		oWord.Selection.TypeText "[provide info: Server VLAN 2]" & VbCrLf
	End If

	oWord.Selection.TypeText "Identifying Number: " & strComputerSystemProduct_IdentifyingNumber & VbCrLf 
	If (bWordExtras) Then 
		oWord.Selection.TypeText "Support contract: "  
		oWord.Selection.TypeText "[provide hardware service level purchased for this server]" & VbCrLf
		oWord.Selection.TypeText "Maintenance and changes to this documentation are recorded in "  
		oWord.Selection.TypeText "[reference to log file/system]." & VbCrLf
		oWord.Selection.TypeText "Continuity and disaster recovery are covered in "  
		oWord.Selection.TypeText "[reference to continuity plan]" & VbCrLf
	End If		

	
	ReportProgress " Writing summary"
	oWord.Selection.Font.Bold = True
	oWord.Selection.TypeText vbCrLf & "TABLE OF CONTENTS" & vbCrLf
	oWord.Selection.Font.Bold = False

	oWord.ActiveDocument.TablesOfContents.Add oWord.Selection.Range, False, 2, 3, , , , ,oWord.ActiveDocument.Styles(wdStyleHeading1)& ";1", True
	
	
	ReportProgress " Inserting Table Of Contents"
	oWord.Selection.TypeText vbCrLf
	oWord.Selection.InsertBreak wdPageBreak

	'--------------------------------------------------------------------------------
	'Chapter 1 - System Information
	'--------------------------------------------------------------------------------
	If (bWordExtras) Then 
		ReportProgress " Writing System Information"
		WriteHeader 1,"System Information"
		oWord.Selection.Style = wdStyleBodyText
		oWord.Selection.TypeText "[Enter information about your server, what the system means to your organization, the purpose of this document etc.]" & vbCrLf
	End If

	'--------------------------------------------------------------------------------
	'Chapter 2 - Hardware Platform
	'--------------------------------------------------------------------------------
	ReportProgress " Writing Hardware Platform"
	WriteHeader 1,"Hardware Platform"
	WriteHeader 2,"General Information"
	
	oWord.Selection.Style = wdStyleBodyText
	oWord.Selection.TypeText "Manufacturer: " & strComputerSystemProduct_Manufacturer & vbCrLf
	oWord.Selection.TypeText "Product name: " & strComputerSystemProduct_Name & vbCrLf
	oWord.Selection.TypeText "Identifying Number: " & strComputerSystemProduct_IdentifyingNumber & vbCrLf 
	oWord.Selection.TypeText "Chassis: " & strChassisType & vbCrLf
	
	oWord.Selection.TypeText VbCrLf

	
	oWord.Selection.Font.Bold = True
	oWord.Selection.TypeText "Processor" & VbCrLf
	oWord.Selection.Font.Bold = False
	oWord.Selection.TypeText "Name: " & strProcessor_Name & VbCrLf
	oWord.Selection.TypeText "Description: " & strProcessor_Description & VbCrLf
	oWord.Selection.TypeText "Speed: " & strProcessor_MaxClockSpeed & " MHz" & vbCrLf
	oWord.Selection.TypeText "L2 Cache Size: " & strProcessor_L2CacheSize & " Kb" & VbCrLf
	oWord.Selection.TypeText "External clock: " & strProcessor_ExtClock & " MHz" & VbCrLf
	If (intProcessors > 0) Then
		oWord.Selection.TypeText "The system has " & intProcessors  & " processors." & VbCrLf
	End If
	If (bProcessorHTSystem) Then
		oWord.Selection.TypeText "The system has Hyper-Threading enabled." & VbCrLf
	End If

	
	oWord.Selection.TypeText VbCrLf
	oWord.Selection.Font.Bold = True
	oWord.Selection.TypeText "Memory" & vbCrLf
	oWord.Selection.Font.Bold = False
	oWord.Selection.TypeText "Total Memory: " & strTotalPhysicalMemoryMB & "Mb" & VbCrLf
	oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrPhysicalMemory.Recordcount + 1, 4
	If Not (bUseSpecificTable) Then
		oWord.Selection.Font.Bold = True
	End If
	oWord.Selection.TypeText "Bank" : oWord.Selection.MoveRight
	If Not (bUseSpecificTable) Then
		oWord.Selection.Font.Bold = True
	End If
	oWord.Selection.TypeText "Capacity" : oWord.Selection.MoveRight
	If Not (bUseSpecificTable) Then
		oWord.Selection.Font.Bold = True
	End If
	oWord.Selection.TypeText "Form" : oWord.Selection.MoveRight
	If Not (bUseSpecificTable) Then
		oWord.Selection.Font.Bold = True
	End If
	oWord.Selection.TypeText "Type" : oWord.Selection.MoveRight
	If Not (objDbrPhysicalMemory.Bof) Then
		objDbrPhysicalMemory.MoveFirst
	End If
	Do Until objDbrPhysicalMemory.EOF
		oWord.Selection.TypeText Cstr(objDbrPhysicalMemory.Fields.Item("BankLabel")) : oWord.Selection.MoveRight
		oWord.Selection.TypeText ReturnBytes2Megabytes(objDbrPhysicalMemory.Fields.Item("Capacity")) & " Mb" : oWord.Selection.MoveRight
		oWord.Selection.TypeText ReturnPhysicalMemoryFormFactor(objDbrPhysicalMemory.Fields.Item("FormFactor")) : oWord.Selection.MoveRight
		oWord.Selection.TypeText ReturnPhysicalMemoryMemoryType(objDbrPhysicalMemory.Fields.Item("MemoryType")) : oWord.Selection.MoveRight
		objDbrPhysicalMemory.MoveNext
	Loop
	oWord.Selection.TypeText VbCrLf

	'CD-ROM Information
	oWord.Selection.Font.Bold = True
	oWord.Selection.TypeText "CD-ROM" & vbCrLf
	oWord.Selection.Font.Bold = False
	oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrCDROMDrive.Recordcount + 1, 3
	If Not (bUseSpecificTable) Then
		oWord.Selection.Font.Bold = True
	End If
	oWord.Selection.TypeText "Name" : oWord.Selection.MoveRight
	If Not (bUseSpecificTable) Then
		oWord.Selection.Font.Bold = True
	End If
	oWord.Selection.TypeText "Drive" : oWord.Selection.MoveRight
	If Not (bUseSpecificTable) Then
		oWord.Selection.Font.Bold = True
	End If
	oWord.Selection.TypeText "Manufacturer" : oWord.Selection.MoveRight
	If Not (objDbrCDROMDrive.Bof) Then
		objDbrCDROMDrive.MoveFirst
	End If
	Do Until objDbrCDROMDrive.EOF
		oWord.Selection.TypeText Cstr(objDbrCDROMDrive.Fields.Item("Name")) : oWord.Selection.MoveRight
		oWord.Selection.TypeText Cstr(objDbrCDROMDrive.Fields.Item("Drive")) : oWord.Selection.MoveRight
		oWord.Selection.TypeText Cstr(objDbrCDROMDrive.Fields.Item("Manufacturer")) : oWord.Selection.MoveRight
		objDbrCDROMDrive.MoveNext
	Loop
	oWord.Selection.TypeText VbCrLf

	If (bHasTapeDrive) Then
		' Tape Drive
		oWord.Selection.Font.Bold = True
		oWord.Selection.TypeText "Tape Drive" & vbCrLf
		oWord.Selection.Font.Bold = False
		oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrTapeDrive.Recordcount + 1, 3
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Name" : oWord.Selection.MoveRight
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Description" : oWord.Selection.MoveRight
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Manufacturer" : oWord.Selection.MoveRight
		If Not (objDbrTapeDrive.Bof) Then
			objDbrTapeDrive.MoveFirst
		End If
		Do Until objDbrTapeDrive.EOF
			oWord.Selection.TypeText Cstr(objDbrTapeDrive.Fields.Item("Name")) : oWord.Selection.MoveRight
			oWord.Selection.TypeText Cstr(objDbrTapeDrive.Fields.Item("Description")) : oWord.Selection.MoveRight
			oWord.Selection.TypeText Cstr(objDbrTapeDrive.Fields.Item("Manufacturer")) : oWord.Selection.MoveRight
			objDbrTapeDrive.MoveNext
		Loop
		oWord.Selection.TypeText VbCrLf
	End If

	
	If (bWMIHardware) Then
		'Sound Card Information
		oWord.Selection.Font.Bold = True
		oWord.Selection.TypeText "Sound Card" & vbCrLf
		oWord.Selection.Font.Bold = False
		oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrSoundDevice.Recordcount + 1, 2
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Name" : oWord.Selection.MoveRight
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Manufacturer" : oWord.Selection.MoveRight
		If Not (objDbrSoundDevice.Bof) Then
			objDbrSoundDevice.MoveFirst
		End If
		Do Until objDbrSoundDevice.EOF
			oWord.Selection.TypeText Cstr(objDbrSoundDevice.Fields.Item("Name")) : oWord.Selection.MoveRight
			oWord.Selection.TypeText Cstr(objDbrSoundDevice.Fields.Item("Manufacturer")) : oWord.Selection.MoveRight
			objDbrSoundDevice.MoveNext
		Loop
		oWord.Selection.TypeText VbCrLf
	End If

			
	If (bWMIHardware) Then
		'oWord.Selection.TypeText VbCrLf
		oWord.Selection.Font.Bold = True
		oWord.Selection.TypeText "Video Controller" & vbCrLf
		oWord.Selection.Font.Bold = False
		oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrVideoController.Recordcount + 1, 3
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Name" : oWord.Selection.MoveRight
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Adapter RAM" : oWord.Selection.MoveRight
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Compatibility" : oWord.Selection.MoveRight
		If Not (objDbrVideoController.Bof) Then
			objDbrVideoController.MoveFirst
		End If
		Do Until objDbrVideoController.EOF
			oWord.Selection.TypeText Cstr(objDbrVideoController.Fields.Item("Name")) : oWord.Selection.MoveRight
			oWord.Selection.TypeText ReturnBytes2Megabytes(objDbrVideoController.Fields.Item("AdapterRAM")) & " Mb" : oWord.Selection.MoveRight
			oWord.Selection.TypeText Cstr(objDbrVideoController.Fields.Item("AdapterCompatibility")) : oWord.Selection.MoveRight
			objDbrVideoController.MoveNext
		Loop
		oWord.Selection.TypeText VbCrLf
	End If
	
	
	
	If (bWMIBios) Then	
		WriteHeader 2,"BIOS Information"
		oWord.Selection.Style = wdStyleBodyText
		oWord.Selection.TypeText "BIOS Version: "  & strBIOS_Version & vbCrLf
		oWord.Selection.TypeText "SMBIOS Version: "  & strBIOS_SMBIOSBIOSVersion & " (Major: " & strBIOS_SMBIOSMajorVersion & ", Minor: " & strBIOS_SMBIOSMinorVersion & ")" & vbCrLf
		oWord.Selection.Font.Bold = True
		oWord.Selection.TypeText "Bios Characteristics: "
		oWord.Selection.Font.Bold = False
		
		For i = 0 To Ubound(arrBIOS_BiosCharacteristics)
			If (strBiosCharacteristics = "") Then
				strBiosCharacteristics = ReturnBiosCharacteristic(arrBIOS_BiosCharacteristics(i))
			Else
				strBiosCharacteristics = strBiosCharacteristics & ", " & ReturnBiosCharacteristic(arrBIOS_BiosCharacteristics(i))
			End If
		Next
		oWord.Selection.TypeText strBiosCharacteristics & vbCrLf
	End If
	
	'--------------------------------------------------------------------------------
	'Chapter 3 - Software Platform
	'--------------------------------------------------------------------------------
	ReportProgress " Writing Software Platform"
	WriteHeader 1,"Software Platform"

	WriteHeader 2,"General Information"
	oWord.Selection.Style = wdStyleBodyText
	oWord.Selection.TypeText "OS Name: " & strOperatingSystem_Caption & VbCrLf	
	oWord.Selection.TypeText "OS Configuration: " & strComputerRole & " in the " & strComputerSystem_Domain & " " & strDomainType & vbCrLf 
	oWord.Selection.TypeText "Windows is installed at " & strOperatingSystem_WindowsDirectory & VbCrLf
	oWord.Selection.TypeText "Install date: " & ConvertWMIDate(strOperatingSystem_InstallDate) & VbCrLf
	oWord.Selection.TypeText "Operating System Language: " & ReturnOperatingSystemLanguage(strOperatingSystem_LanguageCode) & VbCrLf

	If (bRegLastUser) Then
		oWord.Selection.TypeText "Last Logged on User: " & strLastUser & vbCrLf
	End If
	
	

	If (bWMIPatches) Then
		WriteHeader 2,"Installed Patches"
		oWord.Selection.Style = wdStyleBodyText
		If Not (objDbrPatches.Bof) Then
			objDbrPatches.Movefirst
			oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrPatches.Recordcount + 1, 3
			If Not (bUseSpecificTable) Then
				oWord.Selection.Font.Bold = True
			End If
			oWord.Selection.TypeText "Patch ID" : oWord.Selection.MoveRight
			If Not (bUseSpecificTable) Then
				oWord.Selection.Font.Bold = True
			End If
			oWord.Selection.TypeText "Description" : oWord.Selection.MoveRight
			If Not (bUseSpecificTable) Then
				oWord.Selection.Font.Bold = True
			End If
			oWord.Selection.TypeText "Install Date" : oWord.Selection.MoveRight
		Else
			oWord.Selection.TypeText VbCrLf
		End If
		Do Until objDbrPatches.Eof
			oWord.Selection.TypeText CStr(objDbrPatches.Fields.Item("HotfixID")) : oWord.Selection.MoveRight
			oWord.Selection.TypeText  CStr(objDbrPatches.Fields.Item("Description"))  : oWord.Selection.MoveRight
			oWord.Selection.TypeText  CStr(objDbrPatches.Fields.Item("InstallDate"))  : oWord.Selection.MoveRight
			objDbrPatches.MoveNext
			oWord.Selection.TypeText VbCrLf
		Loop
	End If

	If (bWordExtras) Then 	
		WriteHeader 2,"Backup"
		oWord.Selection.Style = wdStyleBodyText
		oWord.Selection.TypeText "[Enter information about the systems backup routines.]" & vbCrLf
	
		WriteHeader 2,"Antivirus"
		oWord.Selection.Style = wdStyleBodyText
		oWord.Selection.TypeText "[Enter information about the systems antivirus protection.]" & vbCrLf
	End If

	If (bWMIApplications) Then
		If (errWin32_Product) Then
			WriteHeader 2,"Currently Installed Programs (Windows Installer)"
			ReportProgress "  Writing Installed Programs (Windows Installer)"
			oWord.Selection.Style = wdStyleBodyText
			oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrProducts.Recordcount + 1, 4
			If Not (bUseSpecificTable) Then
				oWord.Selection.Font.Bold = True
			End If
			oWord.Selection.TypeText "Name" : oWord.Selection.MoveRight
			If Not (bUseSpecificTable) Then
				oWord.Selection.Font.Bold = True
			End If
			oWord.Selection.TypeText "Vendor" : oWord.Selection.MoveRight
			If Not (bUseSpecificTable) Then
				oWord.Selection.Font.Bold = True
			End If
			oWord.Selection.TypeText "Version" : oWord.Selection.MoveRight
			If Not (bUseSpecificTable) Then
				oWord.Selection.Font.Bold = True
			End If
			oWord.Selection.TypeText "Install Date" : oWord.Selection.MoveRight
			If Not (objDbrProducts.Bof) Then
				objDbrProducts.MoveFirst
			End If
			Do Until objDbrProducts.EOF
				oWord.Selection.TypeText CStr(objDbrProducts.Fields.Item("ProductName")) : oWord.Selection.MoveRight
				oWord.Selection.TypeText CStr(objDbrProducts.Fields.Item("Vendor")) : oWord.Selection.MoveRight
				oWord.Selection.TypeText CStr(objDbrProducts.Fields.Item("Version")) : oWord.Selection.MoveRight
				oWord.Selection.TypeText CStr(objDbrProducts.Fields.Item("InstallDate")) : oWord.Selection.MoveRight
				objDbrProducts.MoveNext
			Loop
			oWord.Selection.TypeText VbCrLf
		End If
	End If
	
	If (bRegPrograms) Then
		WriteHeader 2,"Currently Installed Programs (From Registry)"
		ReportProgress "  Writing Installed Programs (Registry)"
		oWord.Selection.Style = wdStyleBodyText
		oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrRegPrograms.Recordcount + 1, 2
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Name" : oWord.Selection.MoveRight
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Version" : oWord.Selection.MoveRight
		If Not (objDbrRegPrograms.Bof) Then
			objDbrRegPrograms.MoveFirst
		End If
		Do Until objDbrRegPrograms.EOF
			oWord.Selection.TypeText CStr(objDbrRegPrograms.Fields.Item("DisplayName")) : oWord.Selection.MoveRight
			oWord.Selection.TypeText CStr(objDbrRegPrograms.Fields.Item("DisplayVersion")) : oWord.Selection.MoveRight
			objDbrRegPrograms.MoveNext
		Loop
		oWord.Selection.TypeText VbCrLf
	End If
	'--------------------------------------------------------------------------------
	'Chapter 4 - Storage
	'--------------------------------------------------------------------------------
	ReportProgress " Writing storage information"
	WriteHeader 1,"Storage"
	WriteHeader 2,"General Information"
	oWord.Selection.Style = wdStyleBodyText
	If Not (objDbrDrives.Bof) Then
		objDbrDrives.Movefirst
	End If
	Do Until objDbrDrives.Eof
		oWord.Selection.Font.Bold = True
		oWord.Selection.TypeText objDbrDrives.Fields.Item("Caption") & " - " & objDbrDrives.Fields.Item("DeviceID") & VbCrLf
		oWord.Selection.Font.Bold = False
		oWord.Selection.TypeText "Interface: " & objDbrDrives.Fields.Item("InterfaceType") & vbcrlf
		oWord.Selection.TypeText "Total Disk Size: " & Round(ReturnBytes2Gigabytes(objDbrDrives.Fields.Item("Size")), 2) & " Gb" & VbCrLf

		If Not (objDbrDisks.Bof) Then
			objDbrDisks.MoveFirst
		End If
		objDbrDisks.Filter = " ParentDriveID='" & objDbrDrives.Fields.Item("DeviceID") & "'"
		Do Until objDbrDisks.Eof
			oWord.Selection.TypeText objDbrDisks.Fields.Item("Caption") & " "
			oWord.Selection.TypeText round(ReturnBytes2Gigabytes(objDbrDisks.Fields.Item("Size")),2) & " Gb "
			oWord.Selection.TypeText "(" & round(ReturnBytes2Gigabytes(objDbrDisks.Fields.Item("FreeSpace")),2) & " Gb Free) "
			oWord.Selection.TypeText objDbrDisks.Fields.Item("FileSystem") & vbcrlf
			objDbrDisks.MoveNext
		Loop
		objDbrDrives.MoveNext
	Loop
	
	'--------------------------------------------------------------------------------
	'Chapter 5 - Network
	'--------------------------------------------------------------------------------
	ReportProgress " Writing Network configuration"
	WriteHeader 1,"Network Configuration"
	WriteHeader 2,"IP Configuration"
	
	For i = 0 To UBound(arrNetadapter_Description)
		oWord.Selection.Style = wdStyleBodyText
		oWord.Selection.Font.Bold = True
		oWord.Selection.TypeText "Network Adapter " & (i + 1) & vbCrLf 
		oWord.Selection.Font.Bold = False
		oWord.Selection.TypeText "Description: " & arrNetadapter_Description(i) & vbCrLf
		oWord.Selection.TypeText "MAC: " & arrNetadapter_MACAddress(i) & vbCrLf
		For j = 0 To UBound(arrNetadapter_IPAddress,2)
			If (arrNetadapter_IPAddress(i,j) <> "") Then
				oWord.Selection.TypeText "IP Address: " & arrNetadapter_IPAddress(i,j) & "/" & arrNetadapter_IPSubnet(i,j) & vbCrLf
			End If
			
		Next
		For j = 0 To UBound(arrNetadapter_DefaultIPGateway,2)
			oWord.Selection.TypeText "Gateway: " & arrNetadapter_DefaultIPGateway(i,j) & vbCrlf
		Next
		strDNSServers  = ""
		For j = 0 To Ubound(arrNetadapter_DNSServerSearchOrder,2)
			If (strDNSServers = "") Then
				strDNSServers = arrNetadapter_DNSServerSearchOrder(i,j)
			Else
				If (arrNetadapter_DNSServerSearchOrder(i,j) <> "") Then
					strDNSServers = strDNSServers & ", " & arrNetadapter_DNSServerSearchOrder(i,j)
				End If
			End If
		Next
		oWord.Selection.TypeText "DNS Servers: " & strDNSServers & vbCrLf
		oWord.Selection.TypeText "DNS Domain: " & arrNetadapter_DNS(i) & vbCrlf
		If (arrNetadapter_WINSPrimaryServer(i) <> "" And arrNetadapter_WINSPrimaryServer(i) <> "127.0.0.0") Then
			oWord.Selection.TypeText "Primary WINS Server: " & arrNetadapter_WINSPrimaryServer(i) & vbCrLf
		End If
		If (arrNetadapter_WINSSecondaryServer(i) <> "" And arrNetadapter_WINSSecondaryServer(i) <> "127.0.0.0") Then
			oWord.Selection.TypeText "Secondary WINS Server: " & arrNetadapter_WINSSecondaryServer(i) & vbCrLf
		End If
		If (arrNetadapter_DHCPEnabled(i)) Then
		 oWord.Selection.TypeText "DHCP Server: " & arrNetadapter_DHCPServer(i) & vbCrlf
		End If
	Next

	If (bWMIIP4Routes) Then
		WriteHeader 2,"IP Routes"
		oWord.Selection.Style = wdStyleBodyText
		If Not (objDbrIP4RouteTable.Bof) Then
			objDbrIP4RouteTable.Movefirst
			oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrIP4RouteTable.Recordcount + 1, 3
			If Not (bUseSpecificTable) Then
				oWord.Selection.Font.Bold = True
			End If
			oWord.Selection.TypeText "Destination" : oWord.Selection.MoveRight
			If Not (bUseSpecificTable) Then
				oWord.Selection.Font.Bold = True
			End If
			oWord.Selection.TypeText "Subnet Mask" : oWord.Selection.MoveRight
			If Not (bUseSpecificTable) Then
				oWord.Selection.Font.Bold = True
			End If
			oWord.Selection.TypeText "Gateway" : oWord.Selection.MoveRight
			Do Until objDbrIP4RouteTable.EOF
				oWord.Selection.TypeText CStr(objDbrIP4RouteTable.Fields.Item("Destination")) : oWord.Selection.MoveRight
				oWord.Selection.TypeText CStr(objDbrIP4RouteTable.Fields.Item("Mask")) : oWord.Selection.MoveRight
				oWord.Selection.TypeText CStr(objDbrIP4RouteTable.Fields.Item("NextHop")) : oWord.Selection.MoveRight
				objDbrIP4RouteTable.MoveNext
			Loop
			oWord.Selection.TypeText VbCrLf
		End If
	End If

	'--------------------------------------------------------------------------------
	'Chapter 6 - General Settings
	'--------------------------------------------------------------------------------
	If (bHasMicrosoftIISv2) Then
		ReportProgress " Writing IIS Information"
		WriteHeader 1,"Internet Information Server"
		ReportProgress " Writing Web Configuration"
		WriteHeader 2,"WWW Server"
		If Not (objDbrIISWebServerSetting.Bof) Then
			objDbrIISWebServerSetting.Movefirst
		End If				
		Do Until objDbrIISWebServerSetting.Eof
			WriteHeader 3,CStr(objDbrIISWebServerSetting.Fields.Item("ServerComment"))
			oWord.Selection.Style = wdStyleBodyText
			'objDbrIISVirtualDirSetting
			If Not (objDbrIISVirtualDirSetting.Bof) Then
				objDbrIISVirtualDirSetting.Movefirst
			End If			
			objDbrIISVirtualDirSetting.Filter = " Name='" & objDbrIISWebServerSetting("Name") & "/root'"
			Do Until objDbrIISVirtualDirSetting.Eof
				oWord.Selection.TypeText "Home Directory: " & objDbrIISVirtualDirSetting("Path") & VbCrLf
				objDbrIISVirtualDirSetting.MoveNext
			Loop
			objDbrIISWebServerBindings.Filter = " ServerName='" & objDbrIISWebServerSetting("Name") & "'"
			If (objDbrIISWebServerBindings.Recordcount > 0) Then
				oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrIISWebServerBindings.Recordcount + 1, 3
				If Not (bUseSpecificTable) Then
					oWord.Selection.Font.Bold = True
				End If
				oWord.Selection.TypeText "Hostname" : oWord.Selection.MoveRight
				If Not (bUseSpecificTable) Then
					oWord.Selection.Font.Bold = True
				End If
				oWord.Selection.TypeText "Ip" : oWord.Selection.MoveRight
				If Not (bUseSpecificTable) Then
					oWord.Selection.Font.Bold = True
				End If
				oWord.Selection.TypeText "Port" : oWord.Selection.MoveRight
				If Not (objDbrIISWebServerBindings.Bof) Then
					objDbrIISWebServerBindings.MoveFirst
				End If
				Do Until objDbrIISWebServerBindings.EOF
					oWord.Selection.TypeText CStr(objDbrIISWebServerBindings("Hostname")) : oWord.Selection.MoveRight
					oWord.Selection.TypeText Cstr(objDbrIISWebServerBindings("Ip")) : oWord.Selection.MoveRight
					oWord.Selection.TypeText CStr(objDbrIISWebServerBindings("Port")) : oWord.Selection.MoveRight
					objDbrIISWebServerBindings.MoveNext
				Loop
				oWord.Selection.TypeText VbCrLf
			End If
			objDbrIISWebServerSetting.Movenext
		Loop
	End If

	
	'--------------------------------------------------------------------------------
	'Chapter 7 - General Settings
	'--------------------------------------------------------------------------------
	ReportProgress " Writing Miscellaneous configuration"
	WriteHeader 1,"Miscellaneous Configuration"
	
	If (bWMIEventLogFile) Then
		ReportProgress " Writing Event Log configuration"
		WriteHeader 2,"Event Log files"
		oWord.Selection.Style = wdStyleBodyText
		If Not (objDbrEventLogFile.Bof) Then
			objDbrEventLogFile.Movefirst
		End If
		Do Until objDbrEventLogFile.Eof
			oWord.Selection.Font.Bold = True
			oWord.Selection.TypeText CStr(objDbrEventLogFile.Fields.Item("LogFileName")) & VbCrLf
			oWord.Selection.Font.Bold = False
			oWord.Selection.TypeText "File: " & objDbrEventLogFile.Fields.Item("Name") & VbCrLf
			oWord.Selection.TypeText "Maximum size: " & ReturnBytes2Megabytes(objDbrEventLogFile.Fields.Item("MaxFileSize")) & " Mb" & VbCrLf
			oWord.Selection.TypeText "Overwrite Policy: " & objDbrEventLogFile.Fields.Item("OverwritePolicy") & VbCrLf
			objDbrEventLogFile.Movenext
		Loop
	End If

	If (bWMILocalGroups) Then
		ReportProgress " Writing Local Groups"
		WriteHeader 2,"Local Groups"
		If Not (objDbrLocalGroups.Bof) Then
			objDbrLocalGroups.Movefirst
		End If
		oWord.Selection.Style = wdStyleBodyText
		Do Until objDbrLocalGroups.Eof
			oWord.Selection.Font.Bold = True
			oWord.Selection.TypeText Cstr(objDbrLocalGroups.Fields.Item("Name")) & VbCrLf
			objDbrGroupUser.Filter = " Groupname='" & objDbrLocalGroups.Fields.Item("Name") & "'"
			oWord.Selection.Font.Bold = False
			Do Until objDbrGroupUser.Eof
				oWord.Selection.TypeText Vbtab & objDbrGroupUser.Fields.Item("Member") & VbCrLf
				objDbrGroupUser.MoveNext
			Loop
			objDbrLocalGroups.Movenext
		Loop
	End If


	If (bWMILocalAccounts) Then
		ReportProgress " Writing Local User Accounts"
		WriteHeader 2,"Local User Accounts"
		If Not (objDbrLocalAccounts.Bof) Then
			objDbrLocalAccounts.Movefirst
		End If
		oWord.Selection.Style = wdStyleBodyText
		Do Until objDbrLocalAccounts.Eof
			oWord.Selection.TypeText Cstr(objDbrLocalAccounts.Fields.Item("UserName")) & VbCrLf
			objDbrLocalAccounts.Movenext
		Loop
	End If

	If (bWMIPrinters = True Or bRegPrintSpoolLocation = True) Then
		ReportProgress " Writing Printer information"
		WriteHeader	2, "Printers"
		oWord.Selection.Style = wdStyleBodyText
		oWord.Selection.TypeText "Print Spooler Location: " & strPrintSpoolLocation & VbCrLf

		If (bWMIPrinters) Then
			If (objDbrPrinters.Recordcount > 0) Then
				oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrPrinters.Recordcount + 1, 3
				If Not (bUseSpecificTable) Then
					oWord.Selection.Font.Bold = True
				End If
				oWord.Selection.TypeText "Name" : oWord.Selection.MoveRight
				If Not (bUseSpecificTable) Then
					oWord.Selection.Font.Bold = True
				End If
				oWord.Selection.TypeText "Driver" : oWord.Selection.MoveRight
				If Not (bUseSpecificTable) Then
					oWord.Selection.Font.Bold = True
				End If
				oWord.Selection.TypeText "Port" : oWord.Selection.MoveRight
				If Not (objDbrPrinters.Bof) Then
					objDbrPrinters.MoveFirst
				End If
				Do Until objDbrPrinters.EOF
					oWord.Selection.TypeText CStr(objDbrPrinters.Fields.Item("Name")) : oWord.Selection.MoveRight
					oWord.Selection.TypeText CStr(objDbrPrinters.Fields.Item("DriverName")) : oWord.Selection.MoveRight
					oWord.Selection.TypeText CStr(objDbrPrinters.Fields.Item("PortName")) : oWord.Selection.MoveRight
					objDbrPrinters.MoveNext
				Loop
				oWord.Selection.TypeText VbCrLf
			End If
		End If

	End If

	WriteHeader 2,"Regional settings"
	ReportProgress " Writing Regional settings"
	oWord.Selection.Style = wdStyleBodyText
	oWord.Selection.TypeText "Time Zone: " & strTimeZone & VbCrLf

	If (bWMIServices) Then 
		ReportProgress " Writing Services"
		WriteHeader 2,"Services"
		oWord.Selection.Style = wdStyleBodyText
		oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrServices.Recordcount + 1, 4
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Name" : oWord.Selection.MoveRight
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Start Mode" : oWord.Selection.MoveRight
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Started" : oWord.Selection.MoveRight
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Start Name" : oWord.Selection.MoveRight
	
		If Not (objDbrServices.Bof) Then
			objDbrServices.MoveFirst
		End If
		Do Until objDbrServices.EOF
			oWord.Selection.TypeText CStr(objDbrServices.Fields.Item("Caption")) : oWord.Selection.MoveRight
			oWord.Selection.TypeText CStr(objDbrServices.Fields.Item("StartMode")) : oWord.Selection.MoveRight
			oWord.Selection.TypeText CStr(objDbrServices.Fields.Item("Started")) : oWord.Selection.MoveRight
			oWord.Selection.TypeText CStr(objDbrServices.Fields.Item("StartName")) : oWord.Selection.MoveRight
			objDbrServices.MoveNext
		Loop
		oWord.Selection.TypeText VbCrLf
	End If

	If (bWMIFileShares) Then
		ReportProgress " Writing File share info"
		WriteHeader 2,"Shares"
		oWord.Selection.Style = wdStyleBodyText
		oWord.ActiveDocument.Tables.Add oWord.Selection.Range, objDbrShares.Recordcount + 1, 3
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Name" : oWord.Selection.MoveRight
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Path" : oWord.Selection.MoveRight
		If Not (bUseSpecificTable) Then
			oWord.Selection.Font.Bold = True
		End If
		oWord.Selection.TypeText "Description" : oWord.Selection.MoveRight
		If Not (objDbrShares.Bof) Then
			objDbrShares.MoveFirst
		End If
		Do Until objDbrShares.EOF
			oWord.Selection.TypeText CStr(objDbrShares.Fields.Item("Name")) : oWord.Selection.MoveRight
			oWord.Selection.TypeText CStr(objDbrShares.Fields.Item("Path")) : oWord.Selection.MoveRight
			oWord.Selection.TypeText CStr(objDbrShares.Fields.Item("Description")) : oWord.Selection.MoveRight
			objDbrShares.MoveNext
		Loop
		oWord.Selection.TypeText VbCrLf
	End If

	ReportProgress " Writing Page File information"
	WriteHeader	2, "Virtual Memory"
	oWord.Selection.Style = wdStyleBodyText
	oWord.Selection.TypeText "Pagefile(s): " & VbCrLf
	If Not (objDbrPagefile.Bof) Then
		objDbrPagefile.Movefirst
	End If
	Do Until objDbrPagefile.Eof
		oWord.Selection.TypeText objDbrPagefile.Fields.Item("Drive") & "\ (" & objDbrPagefile.Fields.Item("InitialSize") & " Mb - " & _
			objDbrPagefile.Fields.Item("MaximumSize") & " Mb)" & VbCrLf
		objDbrPagefile.Movenext
	Loop


	If (bWMIRegistry) Then
		ReportProgress " Writing Registry information"
		WriteHeader 2,"Windows Registry"
		oWord.Selection.Style = wdStyleBodyText
		oWord.Selection.TypeText "Current registry size: " & nCurrentSize &  " Mb" & VbCrLf
		oWord.Selection.TypeText "Maximum registry size: " & nMaximumSize &  " Mb" & VbCrLf
	End If


	
	'--------------------------------------------------------------------------------
	'Chapter 8 - Contact Information
	'--------------------------------------------------------------------------------
	If (bWordExtras) Then 
		ReportProgress " Writing Contact Information"
		WriteHeader 1,"Contact Information"
		oWord.Selection.Style = wdStyleBodyText
		oWord.Selection.TypeText "[System owner [Name, title, telephone, e-mail]" & VbCrLf
		oWord.Selection.TypeText "[Hardware vendors [Name, telephone, website, e-mail etc.] " & VbCrLf
		oWord.Selection.TypeText "[Software vendors [Name, telephone, website, e-mail etc.]" & VbCrLf
	End If
	
	
	'--------------------------------------------------------------------------------
	'Chapter 9 - Passwords
	'--------------------------------------------------------------------------------
	If (bWordExtras) Then 
		ReportProgress " Writing Password section"
		WriteHeader 1,"Passwords"
		oWord.Selection.Style = wdStyleBodyText
		oWord.Selection.TypeText "[Depending on your security policy and where you are planning on keeping this document you might want to delete this section.]" & vbCrLf 
	End If

	If (bUseSpecificTable) Then
		For i = 1 To CInt(oWord.ActiveDocument.Tables.Count)
			oWord.ActiveDocument.Tables(i).Style = strWordTable
		Next
	
	End If
	
	If Not (bUseDOTFile) Then
		' Adding header and footer
		If oWord.ActiveWindow.View.SplitSpecial = wdPaneNone Then
			oWord.ActiveWindow.ActivePane.View.Type = wdPrintView
		Else
			oWord.ActiveWindow.View.Type = wdPrintView
		End If
		oWord.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
		oWord.Selection.TypeText "Basic documentation For " & strComputerSystem_Name
		oWord.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
		oWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
		oWord.Selection.TypeText "Page ("
		oWord.Selection.Fields.Add oWord.Selection.Range, wdFieldPage
		oWord.Selection.TypeText "/"
		oWord.Selection.Fields.Add oWord.Selection.Range, wdFieldNumPages
		oWord.Selection.TypeText ")"
		oWord.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
	End If
	
	' Update table of contents
	ReportProgress " Updating Tables Of Contents"	
	oWord.ActiveDocument.TablesOfContents.Item(1).Update
	
	oWord.ActiveDocument.BuiltInDocumentProperties(wdPropertyComments).Value = "Generated by SYDI-Server " & strScriptVersion & " (http://sydi.sourceforge.net)"
	
	If (bSaveFile) Then 
		ReportProgress " Saving Document"	
		oWord.ActiveDocument.SaveAs strComputerSystem_Name
		If (Err <> 0) Then
		    ReportProgress Err.Number & " -- " &  Err.Description & " (" & strSaveFile & ")"
		    ReportProgress " Would not save to " & strSaveFile
		    ReportProgress " Did you specify a path?"
		    Err.Clear
		    Exit Sub
		End If
	End If
	
	If (bShowWord = False And bSaveFile = True) Then
		ReportProgress " Document Saved"
		oWord.Application.Quit
		Set oWord = Nothing
		ReportProgress "End subroutine: PopulateWordfile()"
	Else
		oWord.Application.Visible = True
		ReportProgress "End subroutine: PopulateWordfile()"
	End If
	
	Set oListTemplate = Nothing
	Set oWord = Nothing
End Sub ' PopulateWordfile


Sub ReportProgress(strMessage)
	WScript.Echo strMessage
End Sub ' ReportProgress

Function ReturnBiosCharacteristic(nBiosCharacteristic)
	Dim strBiosCharacteristic
	Select Case nBiosCharacteristic
		Case 0        
	    	strBiosCharacteristic = "Reserved" 
		Case 1
			strBiosCharacteristic = "Reserved" 
		Case 2
			strBiosCharacteristic = "Unknown"
		Case 3
			strBiosCharacteristic = "BIOS Characteristics Not Supported"
		Case 4
			strBiosCharacteristic = "ISA is supported"
		Case 5
			strBiosCharacteristic = "MCA is supported"
		Case 6
			strBiosCharacteristic = "EISA is supported"
		Case 7
			strBiosCharacteristic = "PCI is supported"
		Case 8
			strBiosCharacteristic = "PC Card (PCMCIA) is supported"
		Case 9
			strBiosCharacteristic = "Plug and Play is supported"
		Case 10
			strBiosCharacteristic = "APM is supported"
		Case 11
			strBiosCharacteristic = "BIOS is Upgradable (Flash)"
		Case 12
			strBiosCharacteristic = "BIOS shadowing is allowed"
		Case 13
			strBiosCharacteristic = "VL-VESA is supported"
		Case 14
			strBiosCharacteristic = "ESCD support is available"
		Case 15
			strBiosCharacteristic = "Boot from CD is supported"
		Case 16
			strBiosCharacteristic = "Selectable Boot is supported"
		Case 17
			strBiosCharacteristic = "BIOS ROM is socketed"
		Case 18
			strBiosCharacteristic = "Boot From PC Card (PCMCIA) is supported"
		Case 19
			strBiosCharacteristic = "EDD (Enhanced Disk Drive) Specification is supported"
		Case 20
			strBiosCharacteristic = "Int 13h - Japanese Floppy for NEC 9800 1.2mb (3.5, 1k Bytes/Sector, 360 RPM) is supported"
		Case 21
			strBiosCharacteristic = "Int 13h - Japanese Floppy for Toshiba 1.2mb (3.5, 360 RPM) is supported"
		Case 22
			strBiosCharacteristic = "Int 13h - 5.25 / 360 KB Floppy Services are supported"
		Case 23
			strBiosCharacteristic = "Int 13h - 5.25 /1.2MB Floppy Services are supported"
		Case 24
			strBiosCharacteristic = "13h - 3.5 / 720 KB Floppy Services are supported"
		Case 25
			strBiosCharacteristic = "Int 13h - 3.5 / 2.88 MB Floppy Services are supported"
		Case 26
			strBiosCharacteristic = "Int 5h, Print Screen Service is supported"
		Case 27
			strBiosCharacteristic = "Int 9h, 8042 Keyboard services are supported"
		Case 28
			strBiosCharacteristic = "Int 14h, Serial Services are supported"
		Case 29
			strBiosCharacteristic = "Int 17h, printer services are supported"
		Case 30
			strBiosCharacteristic = "Int 10h, CGA/Mono Video Services are supported"
		Case 31
			strBiosCharacteristic = "NEC PC-98"
		Case 32
			strBiosCharacteristic = "ACPI supported"
		Case 33
			strBiosCharacteristic = "USB Legacy is supported"
		Case 34
			strBiosCharacteristic = "AGP is supported"
		Case 35
			strBiosCharacteristic = "I2O boot is supported"
		Case 36
			strBiosCharacteristic = "LS-120 boot is supported"
		Case 37
			strBiosCharacteristic = "ATAPI ZIP Drive boot is supported"
		Case 38
			strBiosCharacteristic = "1394 boot is supported"
		Case 39
			strBiosCharacteristic = "Smart Battery supported"
		Case Else
			strBiosCharacteristic = "Unknown (Undocumented)"
	End Select
	ReturnBiosCharacteristic = strBiosCharacteristic
End Function ' ReturnBiosCharacteristic

Function ReturnBytes2Gigabytes(nBytes)
	Dim nGigabytes
	If (IsNumeric(nBytes)) Then
		nGigabytes = nbytes / (1024 * 1024 * 1024)
	Else
		nGigabytes = 0
	End If
	ReturnBytes2Gigabytes = nGigabytes
End Function ' ReturnBytes2Megabytes

Function ReturnBytes2Megabytes(nBytes)
	Dim nMegabytes
	If (IsNumeric(nBytes)) Then
		nMegabytes = nbytes / (1024 * 1024)
	Else
		nMegabytes = 0
	End If
	ReturnBytes2Megabytes = nMegabytes
End Function ' ReturnBytes2Megabytes

Function ReturnOperatingSystemLanguage(strOSLanguageCode)
	Dim strOSLanguageName, strOSTempCode
	strOSTempCode = Cstr(strOSLanguageCode)
	Select Case strOSLanguageCode
		Case "1"
	    	strOSLanguageName = "Arabic"
		Case "4"
	    	strOSLanguageName = "Chinese"
		Case "9"
	    	strOSLanguageName = "English"
		Case "401"
	    	strOSLanguageName = "Arabic - Saudi Arabia"
		Case "402"
	    	strOSLanguageName = "Bulgarian"
		Case "403"
	    	strOSLanguageName = "Catalan"
		Case "404"
	    	strOSLanguageName = "Chinese - Taiwan"
		Case "405"
	    	strOSLanguageName = "Czech"
		Case "406"
	    	strOSLanguageName = "Danish"
		Case "407"
	    	strOSLanguageName = "German"
		Case "408"
	    	strOSLanguageName = "Greek"
		Case "409"
			strOSLanguageName = "English" 
		Case "40A"
	    	strOSLanguageName = "Spanish - Traditional Sort"
		Case "40B"
	    	strOSLanguageName = "Finnish"
		Case "40C"
	    	strOSLanguageName = "French - France"
		Case "40D"
	    	strOSLanguageName = "Hebrew"
		Case "40E"
			strOSLanguageName = "Hungarian"
		Case "40F"
	    	strOSLanguageName = "Icelandic"
		Case "410"
	    	strOSLanguageName = "Italian - Italy"
		Case "411"
	    	strOSLanguageName = "Japanese"
		Case "412"
	    	strOSLanguageName = "Korean"
		Case "413"
	    	strOSLanguageName = "Dutch - Netherlands"
		Case "414"
	    	strOSLanguageName = "Norwegian - Bokmal"
		Case "415"
	    	strOSLanguageName = "Polish"
		Case "416"
	    	strOSLanguageName = "Portuguese - Brazil"
		Case "417"
	    	strOSLanguageName = "Rhaeto-Romanic"
		Case "418"
	    	strOSLanguageName = "Romanian"
		Case "419"
	    	strOSLanguageName = "Russian"
		Case "41A"
	    	strOSLanguageName = "Croatian"
		Case "41B"
	    	strOSLanguageName = "Slovak"
		Case "41C"
	    	strOSLanguageName = "Albanian"
		Case "41D"
	    	strOSLanguageName = "Swedish"
		Case "41E"
	    	strOSLanguageName = "Thai"
		Case "41F"
	    	strOSLanguageName = "Turkish"
		Case "420"
	    	strOSLanguageName = "Urdu"
		Case "421"
	    	strOSLanguageName = "Indonesian"
		Case "422"
	    	strOSLanguageName = "Ukrainian"
		Case "423"
	    	strOSLanguageName = "Belarusian"
		Case "424"
	    	strOSLanguageName = "Slovenian"
		Case "425"
	    	strOSLanguageName = "Estonian"
		Case "426"
	    	strOSLanguageName = "Estonian"
		Case "426"
	    	strOSLanguageName = "Latvian"
		Case "427"
	    	strOSLanguageName = "Lithuanian"
		Case "429"
	    	strOSLanguageName = "Persion"
		Case "42A"
	    	strOSLanguageName = "Vietnamese"
		Case "42D"
	    	strOSLanguageName = "Basque"
		Case "42E"
	    	strOSLanguageName = "Serbian"
		Case "42F"
	    	strOSLanguageName = "Macedonian (FYROM)"
		Case "430"
	    	strOSLanguageName = "Sutu"
		Case "431"
	    	strOSLanguageName = "Tsonga"
		Case "432"
	    	strOSLanguageName = "Tswana"
		Case "434"
	    	strOSLanguageName = "Xhosa"
		Case "435"
	    	strOSLanguageName = "Zulu"
		Case "436"
	    	strOSLanguageName = "Afrikaans"
		Case "438"
	    	strOSLanguageName = "Faeroese"
		Case "43A"
	    	strOSLanguageName = "Maltese"
		Case "43C"
	    	strOSLanguageName = "Gaelic"
		Case "43D"
	    	strOSLanguageName = "Yiddish"
		Case "43E"
	    	strOSLanguageName = "Malay - Malaysia"
		Case "801"
	    	strOSLanguageName = "Arabic - Iraq"
		Case "804"
	    	strOSLanguageName = "Chinese - PRC"
		Case "807"
	    	strOSLanguageName = "German - Switzerland"
		Case "809"
	    	strOSLanguageName = "English - United Kingdom"
		Case "80A"
	    	strOSLanguageName = "Spanish - Mexico"
		Case "80C"
	    	strOSLanguageName = "French - Belgium"
		Case "810"
	    	strOSLanguageName = "Italian - Switzerland"
		Case "813"
	    	strOSLanguageName = "Dutch - Belgium"
		Case "814"
	    	strOSLanguageName = "Norwegian - Nynorsk"
		Case "816"
	    	strOSLanguageName = "Portuguese - Portugal"
		Case "818"
	    	strOSLanguageName = "Romanian - Moldova"
		Case "819"
	    	strOSLanguageName = "Russian - Moldova"
		Case "81A"
	    	strOSLanguageName = "Serbian - Latin"
		Case "81D"
	    	strOSLanguageName = "Swedish - Finland"
		Case "C01"
	    	strOSLanguageName = "Arabic - Egypt"
		Case "C04"
	    	strOSLanguageName = "Chinese - Hong Kong SAR"
		Case "C07"
	    	strOSLanguageName = "German - Austria"
		Case "C09"
	    	strOSLanguageName = "English - Australia"
		Case "C0A"
	    	strOSLanguageName = "Spanish - International Sort"
		Case "C0C"
	    	strOSLanguageName = "French - Canada"
		Case "C1A"
	    	strOSLanguageName = "Serbian - Cyrillic"
		Case "1004"
	    	strOSLanguageName = "Chinese - Singapore"
		Case "1007"
	    	strOSLanguageName = "German - Luxembourg"
		Case "1009"
	    	strOSLanguageName = "English - Canada"
		Case "100A"
	    	strOSLanguageName = "Spanish - Guatemala"
		Case "100C"
	    	strOSLanguageName = "French - Switzerland"
		Case "1401"
	    	strOSLanguageName = "Arabic - Algeria"
		Case "1409"
	    	strOSLanguageName = "English - New Zealand"
		Case "140A"
	    	strOSLanguageName = "Spanish - Costa Rica"
		Case "140C"
	    	strOSLanguageName = "French - Luxembourg"
		Case "1801"
	    	strOSLanguageName = "Arabic - Morocco"
		Case "1809"
	    	strOSLanguageName = "English - Ireland"
		Case "180A"
	    	strOSLanguageName = "Spanish - Panama"
		Case "1C01"
	    	strOSLanguageName = "Arabic - Tunisia"
		Case "1C09"
	    	strOSLanguageName = "English - South Africa"
		Case "1C0A"
	    	strOSLanguageName = "Spanish - Dominican Republic"
		Case "2001"
	    	strOSLanguageName = "Arabic - Oman"
		Case "2009"
	    	strOSLanguageName = "English - Jamaica"
		Case "200A"
	    	strOSLanguageName = "Spanish - Venezuela"
		Case "2401"
	    	strOSLanguageName = "Arabic - Yemen"
		Case "240A"
	    	strOSLanguageName = "Spanish - Colombia"
		Case "2801"
	    	strOSLanguageName = "Arabic - Syria"
		Case "2809"
	    	strOSLanguageName = "English - Belize"
		Case "280A"
	    	strOSLanguageName = "Spanish - Peru"
		Case "2C01"
	    	strOSLanguageName = "Arabic - Jordan"
		Case "2C09"
	    	strOSLanguageName = "English - Trinidad"
		Case "2C0A"
	    	strOSLanguageName = "Spanish - Argentina"
		Case "3001"
	    	strOSLanguageName = "Arabic - Lebanon"
		Case "300A"
	    	strOSLanguageName = "Spanish - Ecuador"
		Case "3401"
	    	strOSLanguageName = "Arabic - Kuwait"
		Case "340A"
	    	strOSLanguageName = "Spanish - Chile"
		Case "3801"
	    	strOSLanguageName = "Arabic - U.A.E."
		Case "380A"
	    	strOSLanguageName = "Spanish - Uruguay"
		Case "3C01"
	    	strOSLanguageName = "Arabic - Bahrain"
		Case "3C0A"
	    	strOSLanguageName = "Spanish - Paraguay"
		Case "4001"
	    	strOSLanguageName = "Arabic - Qatar"
		Case "400A"
	    	strOSLanguageName = "Spanish - Bolivia"
		Case "440A"
	    	strOSLanguageName = "Spanish - El Salvador"
		Case "480A"
	    	strOSLanguageName = "Spanish - Honduras"
		Case "4C0A"
	    	strOSLanguageName = "Spanish - Nicaragua"
		Case "500A"
	    	strOSLanguageName = "Spanish - Puerto Rico"
		Case Else
			strOSLanguageName = "Unknown"
	End Select
	ReturnOperatingSystemLanguage = strOSLanguageName	
End Function ' ReturnOperatingSystemLanguage


Function ReturnPhysicalMemoryFormFactor(nFormFactor)
	Dim strFormFactor
	Select Case nFormFactor
		Case 0        
	    	strFormFactor = "Unknown" 
		Case 1
			strFormFactor = "Other" 
		Case 2
			strFormFactor = "SIP"
		Case 3
			strFormFactor = "DIP"
		Case 4
			strFormFactor = "ZIP"
		Case 5
			strFormFactor = "SOJ"
		Case 6
			strFormFactor = "Proprietary"
		Case 7
			strFormFactor = "SIMM"
		Case 8
			strFormFactor = "DIMM"
		Case 9
			strFormFactor = "TSOP"
		Case 10
			strFormFactor = "PGA"
		Case 11
			strFormFactor = "RIMM"
		Case 12
			strFormFactor = "SODIMM"
		Case 13
			strFormFactor = "SRIMM"
		Case 14
			strFormFactor = "SMD"
		Case 15
			strFormFactor = "SSMP"
		Case 16
			strFormFactor = "QFP"
		Case 17
			strFormFactor = "TQFP"
		Case 18
			strFormFactor = "SOIC"
		Case 19
			strFormFactor = "LCC"
		Case 20
			strFormFactor = "PLCC"
		Case 21
			strFormFactor = "BGA"
		Case 22
			strFormFactor = "FPBGA"
		Case 23
			strFormFactor = "LGA"
		Case Else
			strFormFactor = "Unknown"
	End Select
	ReturnPhysicalMemoryFormFactor = strFormFactor	
End Function ' ReturnPhysicalMemoryFormFactor

Function ReturnPhysicalMemoryMemoryType(nMemoryType)
	Dim strMemoryType
	Select Case nMemoryType
		Case 0        
	    	strMemoryType = "Unknown" 
		Case 1
			strMemoryType = "Other" 
		Case 2
			strMemoryType = "DRAM"
		Case 3
			strMemoryType = "Synchronous DRAM"
		Case 4
			strMemoryType = "Cache DRAM"
		Case 5
			strMemoryType = "EDO"
		Case 6
			strMemoryType = "EDRAM"
		Case 7
			strMemoryType = "VRAM"
		Case 8
			strMemoryType = "SRAM"
		Case 9
			strMemoryType = "RAM"
		Case 10
			strMemoryType = "ROM"
		Case 11
			strMemoryType = "Flash"
		Case 12
			strMemoryType = "EEPROM"
		Case 13
			strMemoryType = "FEPROM"
		Case 14
			strMemoryType = "EPROM"
		Case 15
			strMemoryType = "CDRAM"
		Case 16
			strMemoryType = "3DRAM"
		Case 17
			strMemoryType = "SDRAM"
		Case 18
			strMemoryType = "SGRAM"
		Case 19
			strMemoryType = "RDRAM"
		Case 20
			strMemoryType = "DDR"
		Case Else
			strMemoryType = "Unknown"
	End Select
	ReturnPhysicalMemoryMemoryType = strMemoryType
End Function ' ReturnPhysicalMemoryMemoryType


Function Scrub(strInput)
	If (IsNull(strInput)) Then
		strInput = ""
	End If
	Scrub = strInput
End Function ' Scrub
	
Function Scrub4XML(strInput)
	If (IsNull(strInput)) Then
		strInput = ""
	Else
		strInput = Replace(strInput,"&","&#38;")
		strInput = Replace(strInput,"""","&quot;")
		strInput = Replace(strInput,"<","&lt;")
		strInput = Replace(strInput,">","&gt;")
		strInput = Replace(strInput,"'","&apos;")
		strInput = Replace(strInput,"–","-") '  Can break XML files with SYDI-Transform
		
	End If
	Scrub4XML = strInput
End Function ' Scrub4XML

Sub SetOptions(strOption)
	Dim strFlag, strParameter
	Dim nArguments
	nArguments = Len(strOption)
	If (nArguments < 2) Then
		bInvalidArgument = True
	Else
		strFlag = Left(strOption,2)
		Select Case strFlag
			Case "-b"
				strWordTable = ""
				bUseSpecificTable = True
				If (nArguments > 2) Then
					strWordTable = Right(strOption,(nArguments - 2))
				End If
				If (strWordTable = "") Then
					bInvalidArgument = True
				End If
			Case "-D"
					bAllowErrors = False
			Case "-w"
				bWMIBios = False
				bWMIRegistry = False
				bWMIApplications = False
				bWMIPatches = False
				bWMIFileShares = False
				bWMIServices = False
				bWMIPrinters = False
				bWMIEventLogFile = False
				bWMILocalAccounts = False
				bWMIIP4Routes = False
				bWMILocalGroups = False
				bWMIRunningProcesses = False
				bWMIHardware = False
				bWMIStartupCommands = False
				If (nArguments > 2) Then
					For i = 3 To nArguments
						strParameter = Mid(strOption,i,1)
						Select Case strParameter
							Case "b"
								bWMIBios = True
							Case "r"
								bWMIRegistry = True
							Case "a"
								bWMIApplications = True
							Case "q"
								bWMIPatches = True
							Case "e"
								bWMIEventLogFile = True
							Case "f"
								bWMIFileShares = True
							Case "g"
								bWMILocalGroups = True
							Case "h"
								bWMIHardware = True
							Case "i"
								bWMIIP4Routes = True
							Case "s" 
								bWMIServices = True 
							Case "S" 
								bWMIStartupCommands = True 
							Case "p"
								bWMIPrinters = True
							Case "P"
								bWMIRunningProcesses = True
							Case "u"
								bWMILocalAccounts = True
							Case Else
								bInvalidArgument = True
						End Select
					Next
				End If
			Case "-r"
				bRegWindowsComponents = False
				bRegDomainSuffix = False
				bRegPrintSpoolLocation = False
				bRegPrograms = False
				bRegLastUser = False
				bDoRegistryCheck = False
				If (nArguments > 2) Then

					For i = 3 To nArguments
						strParameter = Mid(strOption,i,1)
						Select Case strParameter
							Case "a"
								bRegPrograms = True
								bDoRegistryCheck = True
							Case "c"
								bRegWindowsComponents = True
								bDoRegistryCheck = True
							Case "d"
								bRegDomainSuffix = True
								bDoRegistryCheck = True
							Case "l"
								bRegLastUser = True
								bDoRegistryCheck = True
							Case "p"
								bRegPrintSpoolLocation = True
								bDoRegistryCheck = True
							Case Else
								bInvalidArgument = True
						End Select
					
					Next
				End If
			Case "-e"
				If (nArguments > 2) Then

					For i = 3 To nArguments
						strParameter = Mid(strOption,i,1)
						Select Case strParameter
							Case "w"
								strExportFormat = "word"
							Case "x"
								strExportFormat = "xml"
							Case Else
								bInvalidArgument = True
						End Select
					
					Next
				End If
			Case "-s"
				If (nArguments > 2) Then
					strParameter = Mid(strOption,3,1)
					Select Case strParameter
						Case "h"
							strStylesheet = "html"
						Case "t"
							strStylesheet = "freetext"
							If (Len(strOption) < 4) Then
								bInvalidArgument = True
							Else
								strXSLFreeText = Mid(strOption,4)
							End If
						Case Else
							bInvalidArgument = True
					End Select
				End If
			Case "-t"
				If (nArguments > 2) Then
					strComputer = Right(strOption,(nArguments - 2))
				End If
			Case "-d"
					bShowWord = False
			Case "-n"
					bWordExtras  = False
			Case "-T"
					bUseDOTFile  = True
				If (nArguments > 2) Then
					strDOTFile = Right(strOption,(nArguments - 2))
				Else
					bInvalidArgument = True
				End If
			Case "-o"
					bSaveFile  = True
				If (nArguments > 2) Then
					strSaveFile = Right(strOption,(nArguments - 2))
				Else
					bInvalidArgument = True
				End If
			Case "-f"
				If (nArguments > 2) Then
					nBaseFontSize = Right(strOption,(nArguments - 2))
					If Not (IsNumeric(nBaseFontSize)) Then
						bInvalidArgument = True
					End If
				End If
			Case "-u"
				strUserName = ""
				bAlternateCredentials = True
				If (nArguments > 2) Then
					strUserName = Right(strOption,(nArguments - 2))
				End If
			Case "-p"
				strPassword = ""
				If (nArguments > 2) Then
					strPassword = Right(strOption,(nArguments - 2))
				End If
			Case "-v"
					bCheckVersion = True
			Case "-h"
				bDisplayHelp = True
			Case Else
				bInvalidArgument = True
		End Select
	
	End If

End Sub ' SetOptions

Sub SystemRolesDefine()
	bRoleDC = False
	bRoleDHCP = False
	bRoleDNS = False
	bRoleFile = False
	bRoleFTP = False
	bRoleIAS = False
	bRoleMediaServer = False
	bRoleNews = False
	bRolePKI = False
	bRolePrint = False
	bRoleRAS = False
	bRoleRIS = False
	bRoleSMTP = False
	bRoleSQL = False
	bRoleTS = False
	bRoleWINS = False
	bRoleWWW = False
End Sub ' SystemRolesDefine

Sub SystemRolesSet()
	Set objDbrSystemRoles = CreateObject("ADOR.Recordset")
	objDbrSystemRoles.Fields.Append "Role", adVarChar, MaxCharacters
	objDbrSystemRoles.Open

	If (bRoleDC) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "DC"
		objDbrSystemRoles.Update
	End If
	If (bRoleDHCP) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "DHCP"
		objDbrSystemRoles.Update
	End If
	If (bRoleDNS) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "DNS"
		objDbrSystemRoles.Update
	End If
	If (bRoleFile) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "File"
		objDbrSystemRoles.Update
	End If
	If (bRoleFTP) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "FTP"
		objDbrSystemRoles.Update
	End If
	If (bRoleIAS) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "IAS"
		objDbrSystemRoles.Update
	End If
	If (bRoleMediaServer) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "Media"
		objDbrSystemRoles.Update
	End If
	If (bRoleNews) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "News"
		objDbrSystemRoles.Update
	End If
	If (bRolePKI) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "PKI"
		objDbrSystemRoles.Update
	End If
	If (bRolePrint) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "Print"
		objDbrSystemRoles.Update
	End If
	If (bRoleRAS) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "RAS"
		objDbrSystemRoles.Update
	End If
	If (bRoleRIS) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "RIS"
		objDbrSystemRoles.Update
	End If
	If (bRoleSMTP) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "SMTP"
		objDbrSystemRoles.Update
	End If
	If (bRoleSQL) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "SQL"
		objDbrSystemRoles.Update
	End If
	If (bRoleTS) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "TS"
		objDbrSystemRoles.Update
	End If
	If (bRoleWINS) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "WINS"
		objDbrSystemRoles.Update
	End If
	If (bRoleWWW) Then
		objDbrSystemRoles.AddNew
		objDbrSystemRoles("Role") = "WWW"
		objDbrSystemRoles.Update
	End If

End Sub ' SystemRolesSet


Sub WriteHeader(nHeaderLevel,strHeaderText)
	Const wdStyleHeading1 = -2
	Const wdStyleHeading2 = -3
	Const wdStyleHeading3 = -4
	Const wdStyleHeading4 = -5

	Select Case nHeaderLevel
		Case 1
			oWord.Selection.Style = wdStyleHeading1
		Case 2
			oWord.Selection.Style = wdStyleHeading2
		Case 3
			oWord.Selection.Style = wdStyleHeading3
		Case 4
			oWord.Selection.Style = wdStyleHeading4
	End Select
	oWord.Selection.Range.ListFormat.ApplyListTemplate oListTemplate, True
	oWord.Selection.TypeText strHeaderText & vbCrLf
End Sub ' WriteHeader



'==========================================================