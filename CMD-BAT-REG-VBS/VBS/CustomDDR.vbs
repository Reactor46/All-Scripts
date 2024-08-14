' Script to generate a DDR destined for the ConfigMgr Site servers inbox\DDM.BOX folder to modify an existing Clients Resource record

' By Robert Marshall - SMSMarshall (2014)

' Version 1.0

' TODO

' Download the ConfigMgr 2012 SDK from here: www.microsoft.com/en-gb/download/details.aspx?id=29559
' If the link is dead this script is very old, it may work with a later version of ConfigMgr, see if the DLL's are in the current SDK and give it a go if they are.

' Extract the two DLL's smsrsgen.dll and smsrsgenctl.dll and put them and this script in the same folder, for ease of reference call the new folder CustomDDR.

' To get help type cscript /nologo CreateDDR.vbs /? 

' NOTE: The script needs to be run using the 64-bit CScript engine, not the 32-bit engine, we then invoke a 32-bit CScript to accomodate the 32-bit SDK DLL usage

' The following errors are produced: ' *** UPDATE THIS LIST ***

' 1 - Required argument 'SiteServer' is missing
' 2 - Required argument 'Share' is missing
' 3 - No override and required argument 'Path' is missing
' 4 - No override and required argument 'PropertyName' is missing
' 5 - Required argument 'AgentName' is missing
' 6 - Cannot find the property

' OPTION EXPLICIT

' ON ERROR RESUME NEXT

' Define variables.

' CIM Constants

	CONST 	CIM_ILLEGAL     = 4095 '   // 0xFFF
	CONST   CIM_EMPTY       = 0 '    // 0x0
	CONST   CIM_SINT8       = 16 '   // 0x10
	CONST   CIM_UINT8       = 17 '   // 0x11
	CONST   CIM_SINT16      = 2 '    // 0x2
	CONST   CIM_UINT16      = 18 '   // 0x12
	CONST   CIM_SINT32      = 3 '    // 0x3
	CONST   CIM_UINT32      = 19 '   // 0x13
	CONST   CIM_SINT64      = 20 '   // 0x14
	CONST   CIM_UINT64      = 21 '   // 0x15
	CONST   CIM_REAL32      = 4 '    // 0x4
	CONST   CIM_REAL64      = 5 '    // 0x5
	CONST   CIM_BOOLEAN     = 11 '   // 0xB
	CONST   CIM_STRING      = 8 '    // 0x8
	CONST   CIM_DATETIME    = 101 '  // 0x65
	CONST   CIM_REFERENCE   = 102 '  // 0x66
	CONST   CIM_CHAR16      = 103 '  // 0x67
	CONST   CIM_OBJECT      = 13 '   // 0xD
	CONST   CIM_FLAG_ARRAY  = 8192 ' // 0x2000

' Registry Constants

	Const HKEY_CLASSES_ROOT   = &H80000000
	Const HKEY_CURRENT_USER   = &H80000001
	Const HKEY_LOCAL_MACHINE  = &H80000002
	Const HKEY_USERS          = &H80000003

	Const REG_SZ        = 1
	Const REG_EXPAND_SZ = 2 ' Not implemented
	Const REG_BINARY    = 3 ' Not implemented
	Const REG_DWORD     = 4
	Const REG_MULTI_SZ  = 7

' File System Constants

	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	Const OverwriteExisting = True
	Const ForWriting = 2
	Const ForAppending = 8

	Dim objWMIService, sSMSGUID, oSMSClient, resourceID, existingResource, newDDR, siteCode, organizationalUnit, allCCMClientObjects, eachCCMClientObject, WshShell, WshNetwork, Return
	Dim DDRDestPath, fixedDDRDestPath, fso, SiteServerShare, fixedSiteServerShare, colNamedArguments, oClass, colClass, TSEnv, tempDDRFileName, executionState

	Dim resultsArray(1)

	Set colNamedArguments = WScript.Arguments.Named

	if colNamedArguments.Exists("?") = TRUE or colNamedArguments.Count = 0 Then

		WScript.Echo 
		WScript.Echo "Creates a custom Discovery Data Record (DDR) and delivers it to a ConfigMgr Site Server"
		WScript.Echo
		WScript.Echo "By Robert Marshall - SMSMarshall Ltd (2014)"
		WScript.Echo
		WScript.Echo "Required properties: SiteServer, Share, Path, PropertyName,"
		WScript.Echo " AgentName"
		WScript.Echo
		WScript.Echo "Optional properties: Namespace, Temp, ForceValue, PropertyType, Simulation, "
		WScript.Echo "  PropertyLength, PropertyTypeAuto"
		WScript.Echo
		WScript.Echo "SiteServer:"
		WScript.Echo " The FQDN or NetBIOS name of the site server"
		WScript.Echo "Share:"
		WScript.Echo " The Share Name for the inbox\DDM.BOX folder"
		WScript.Echo "Namespace:"
		WScript.Echo " The Registry Hive, WMI Namespace, Environment or Task Sequence Variable,"
		WScript.Echo "  if not specified then defaults to WMI and the default root\cimv2 namespace"
		WScript.Echo " Use the following Registry Hives: HKLM, HKCU, HKCR, HKCC, HKU"
		WScript.Echo " Use ENV for Environment and TSENV for Task Sequence Variable"
		WScript.Echo "Path:"
		WScript.Echo " Path to the registry value, the WMI Class\Property combination or"
		WScript.Echo " variable name"
		WScript.Echo " For WMI use Class\Property, include a WHERE clause as"
		WScript.Echo " [WHERE property [< > = LIKE] value]"
		WScript.Echo " Do not include qoutes around the WHERE clauses value as this is done by"
		WScript.Echo " the script"
		WScript.Echo "PropertyName:"
		WScript.Echo " The Name of the New or Existing Discovery Property to be added to the Resource"
		WScript.Echo " Record"
		WScript.Echo "AgentName:"
		WScript.Echo " The Name of the agent that this Discovery Record will be submitted as, can be"
		WScript.Echo " anything but we recommend you do not use the same name as an Agent handling"
		WScript.Echo " existing discovery properties so as to save confusion"
		WScript.Echo "ForceValue:"
		WScript.Echo " When specified, the Namespace and Path arguments are ignored"
		WScript.Echo " and this value is directly added to the Client's Resource Record"
		WScript.Echo "PropertyType:"
		WScript.Echo " When specified the properties type is cast as either String or Integer"
		WScript.Echo " Note: Be careful recasting values as this can lead to unexpected results"
		WScript.Echo "  When not specified, the default is to detect the property type and is"
		WScript.Echo "  the safer method to use"
		WScript.Echo "PropertyLength:"
		WScript.Echo " When PropertyType is String, the width of the string is defined"
		WScript.Echo " Note: When not specified the default width of 64 is used"
		WScript.Echo "PropertyTypeAuto:"
		WScript.Echo " When specified the Environment variable type is detected as either"
		WScript.Echo " String or Integer based on its contents instead of using the default"
		WScript.Echo " type of String"
		WScript.Echo " Note: Be careful recasting values as this can lead to unexpected results"
		WScript.Echo "Simulation:"
		WScript.Echo " When specified the DDR is not copied to the Site server and remains in"
		WScript.Echo " the temporary folder, so that you can simulate and run without impact"
		WScript.Echo
		WScript.Echo "Example commands:"
		WScript.Echo
		WScript.Echo "Use Registry, obtain Path value from HKLM\Software\7Zip:"
		WScript.Echo
		WScript.Echo "CustomDDR.vbs /SiteServer:CONFIGMGR.SMSMARSHALL.COM /Temp:C:\TEMP"
		WScript.Echo " /Share:CustomDDR$ /Namespace:HKLM /Path:" & CHR(34) & "Software\7Zip\Path" & CHR(34)
		WScript.Echo " /PropertyName:7ZipPath /AgentName:CustomDDR"
		WScript.Echo
		WScript.Echo "Use WMI, obtain Model from WIN32_ComputerSystem:"
		WScript.Echo
		WScript.Echo "CustomDDR.vbs /SiteServer:CONFIGMGR.SMSMARSHALL.COM /Temp:C:\TEMP"
		WScript.Echo " /Share:CustomDDR$ /PATH:" & CHR(34) & "WIN32_ComputerSystem\Model" & CHR(34)
		WScript.Echo " /PropertyName:Model /AgentName:CustomDDR"
		WScript.Echo
		WScript.Echo "Use WMI, obtain Model from WIN32_ComputerSystem and include a WHERE clause:"
		WScript.Echo
		WScript.Echo "CustomDDR.vbs /SiteServer:CONFIGMGR.SMSMARSHALL.COM /Temp:C:\TEMP"
		WScript.Echo " /Share:CustomDDR$ /PATH:" & CHR(34) & "WIN32_ComputerSystem\Model WHERE Model = TESTVALUE" & CHR(34)
		WScript.Echo " /PropertyName:Model /AgentName:CustomDDR"
		WScript.Echo
		WScript.Echo "Simulation mode, DDR will be created but not copied to the Site server"
		WScript.Echo 
		WScript.Echo "CustomDDR.vbs /SiteServer:CONFIGMGR.SMSMARSHALL.COM /Share:CustomDDR$"
		WScript.Echo " /TEMP:C:\TEMP /FORCEVALUE:" & CHR(34) & "ForcedValue" & CHR(34)
		WScript.Echo " /PropertyName:ForceValue /AgentName:CustomDDR /Simulation:TRUE"
		WScript.Echo 
		WScript.Echo "Use Environment, obtain a specific environment variable "
		WScript.Echo 
		WScript.Echo "CustomDDR.vbs /SiteServer:CONFIGMGR.SMSMARSHALL.COM /Share:CustomDDR$"
		WScript.Echo " /Namespace:ENV /Path:SystemRoot /PropertyName:SystemRoot"
		WScript.Echo " /AgentName:CustomDDR"
		WScript.Echo 
		WScript.Echo "Use Task Sequence Environment, obtain a specific Task Sequence Environment variable "
		WScript.Echo 
		WScript.Echo "CustomDDR.vbs /SiteServer:CONFIGMGR.SMSMARSHALL.COM /Share:CustomDDR$"
		WScript.Echo " /Namespace:TSENV /Path:TESTVAR /PropertyName:TSVAR"
		WScript.Echo " /AgentName:CustomDDR"
		WScript.Echo 
		WScript.Echo "Using the Override to force a specific value to be used:"
		WScript.Echo 
		WScript.Echo "CustomDDR.vbs /SiteServer:CONFIGMGR.SMSMARSHALL.COM /Share:CustomDDR$"
		WScript.Echo " /FORCEVALUE:" & CHR(34) & "ForcedValue" & CHR(34) & " /PropertyName:ForceValue"
		WScript.Echo " /AgentName:CustomDDR"
		WScript.Echo

		WScript.Quit(0)

	End If

	Set WshShell = WScript.CreateObject("WScript.Shell")

	If colNamedArguments.Exists("Temp") = False Then ' No DDRTemp specified defaulting

		DDRDestPath = WshShell.ExpandEnvironmentStrings("%windir%") & "\temp\"

		LogMessage "################################################################################"
		LogMessage "CustomDDR Started " & Date & " " & Time
		LogMessage "################################################################################"

		LogMessage "Default setting for temporary DDR file location set to " & WshShell.ExpandEnvironmentStrings("%windir%") & "\TEMP"

	Else

		DDRDestPath = WScript.Arguments.Named.Item("Temp")

		LogMessage "################################################################################"
		LogMessage "CustomDDR Started " & Date & " " & Time
		LogMessage "################################################################################"

		LogMessage "Custom temporary DDR file location specified as " & WScript.Arguments.Named.Item("Temp")

	End If

	If colNamedArguments.Exists("SiteServer") = False Then ' Required argument 'SiteServer' is missing

		LogMessage "Required argument 'SiteServer' is missing"

		WScript.Quit(1)

	End If

	If colNamedArguments.Exists("Share") = False Then

		LogMessage "Required argument 'DDRShare' is missing"

		WScript.Quit(2) ' Required argument 'DDRShare' is missing

	Else

		SiteServerShare = "\\" & WScript.Arguments.Named.Item("SiteServer") & "\" & WScript.Arguments.Named.Item("Share")

	End If

	If colNamedArguments.Exists("ForceValue") = False and colNamedArguments.Exists("Path") = False Then

		LogMessage "No override and required argument 'Path' is missing"

		WScript.Quit(3) ' No override and required argument 'Path' is missing

	End If

	If colNamedArguments.Exists("PropertyName") = False Then

		LogMessage "No override and required argument 'PropertyName' is missing"

		WScript.Quit(4) ' No override and required argument 'PropertyName' is missing

	End If

	If colNamedArguments.Exists("AgentName") = False Then

		LogMessage "Required argument 'AgentName' is missing"

		WScript.Quit(5) ' Required argument 'AgentName' is missing

	End If

	If colNamedArguments.Exists("propertyLength") = False Then ' No PropertyTypeLength specified defaulting

		LogMessage "Default setting for Property Length set to 64"

		DDRPropertyLength = 64

	Else

		LogMessage "Custom Property Length specified as " & WScript.Arguments.Named.Item("PropertyLength")

		DDRPropertyLength = WScript.Arguments.Named.Item("PropertyLength")

	End If

' Generate a unique filename for the DDR

	tempDDRFileName = DDRFileName

	LogMessage "Generated DDR temporary file name: " & tempDDRFileName

' Create the DDR

	Call CreateDDR

' Submit DDR

	LogMessage "Preparing to copy DDR to Site Server"

	If colNamedArguments.Exists("Simulation") = False Then

		Set fso = CreateObject("Scripting.FileSystemObject")

		' Prepare the two arguments

		If not MID(SiteServerShare,LEN(SiteServerShare),1) = "\" Then

			fixedSiteServerShare = SiteServerShare & "\"
	
		Else

			fixedSiteServerShare = SiteServerShare

		End If

		If not MID(DDRDestPath,LEN(DDRDestPath),1) = "\" Then

			fixedDDRDestPath = DDRDestPath & "\"

		Else

			fixedDDRDestPath = DDRDestPath

		End If

		Err.Clear

		LogMessage "Remote location: " & fixedSiteServerShare & tempDDRFileName
		LogMessage "Local location: " & fixedDDRDestPath & tempDDRFileName

		WScript.Sleep 100 ' Wait an extra moment for the DDR 32-bit Script to complete

		Err.Clear

		Set fso = CreateObject("Scripting.FileSystemObject")

		On Error Resume Next

		fso.CopyFile fixedDDRDestPath & tempDDRFileName, fixedSiteServerShare & tempDDRFileName, OverwriteExisting	

		If Err.Number <> 0 Then 

			LogMessage "Could not copy the DDR to the Inbox"

			WScript.Quit(8) ' Could not copy the DDR to the Inbox

		End If

		' Remove the local DDR file

		Err.Clear

		fso.DeleteFile(fixedDDRDestPath & tempDDRFileName)

		If Err.Number <> 0 Then 

			LogMessage "Could not delete the temporary DDR file"

			WScript.Quit(9) ' Could not delete the temporary DDR file

		End If

		On Error Goto 0

		Set fso = Nothing

		LogMessage "DDR was reported as being copied to the Site Server, review the Site Server DDM log for processing status"

	Else

		LogMessage "Simulation mode enabled, DDR was not copied to the Site Server or deleted from its temporary location"

	End If

' Clean up

	Set fso = Nothing
	Set WshShell = Nothing

	LogMessage "################################################################################"
	LogMessage "CustomDDR Ended " & Date & " " & Time
	LogMessage "################################################################################"


' Functions and Procedures

Function ComputerName ' Obtain Computer Name

	Set WshNetwork = CreateObject("WScript.Network")

	ComputerName = WshNetwork.ComputerName

	Set WshNetwork = Nothing	

End Function

Function GetSiteCode

	On Error Resume Next

	Err.Clear

	Set oSMSClient = CreateObject ("Microsoft.SMS.Client")

	GetSiteCode = oSMSClient.GetAssignedSite()

	If Err.Number = 0 Then

		LogMessage "Site Code is " & GetSiteCode

	Else

		LogMessage "Failure to obtain a Site code, either Client not installed or not Assigned to a Primary Site server"

		WScript.Quit(12)

	End If

	On Error Goto 0

	Set oSMSClient=nothing

End Function

Function GetClientGUID

	GetClientGUID = GetObject("WinMgmts:root\CCM:CCM_Client=@").ClientID

End Function

Sub CreateDDR

' Retrieve the value and type

	LogMessage "Retrieving Value from specified repository"

	splitArray = Split(RetrieveProperty,"&&&&",-1,1) ' Property value and Property Type will now be retrieved by the RetrieveProperty function

	runString =             "/DDRDestPath:" & CHR(34) & DDRDestPath & CHR(34) & " "
	runString = runString & "/DDRProperty:" & CHR(34) & splitarray(0) & CHR(34) & " "
	runString = runString & "/DDRPropertyType:" & CHR(34) & splitarray(1) & CHR(34) & " "
	runString = runString & "/DDRPropertyName:" & CHR(34) & colNamedArguments.Item("PropertyName") & CHR(34) & " "
	runString = runString & "/DDRClientGUID:" & CHR(34) & GetClientGUID & CHR(34) & " "
	runString = runString & "/DDRSiteCode:" & CHR(34) & GetSiteCode & CHR(34) & " "
	runString = runString & "/DDRPropertyLength:" & CHR(34) & DDRPropertyLength & CHR(34) & " "
	runString = runString & "/DDRAgentName:" & CHR(34) & colNamedArguments.Item("AgentName") & CHR(34) & " "
	runString = runString & "/DDRFileName:" & CHR(34) & tempDDRFileName & CHR(34) & " "
	runString = runString & "/DDRCWD:" & CHR(34) & WshShell.CurrentDirectory & CHR(34)

	LogMessage "Running CScript using the following arguments: " & runString

	Err.Clear

	LogMessage "Launching Script now"

	Set fso = CreateObject("Scripting.FileSystemObject")

	If (fso.FolderExists(WshShell.ExpandEnvironmentStrings("%windir%") & "\SYSWOW64")) Then ' We must be running within a 64-bit Operating System so run the 32-bit CSCRIPT engine

		LogMessage "Detected we are running on a 64-bit OS, so using 32-bit CSCRIPT Engine"

		Set oExec = WshShell.Exec(WshShell.ExpandEnvironmentStrings("%windir%") & "\SYSWOW64\CSCRIPT.EXE /NOLOGO CustomDDR-Write.vbs " & runString)

	Else ' We must be running within a 32-bit Operating System so run the standard CSCRIPT engine

		LogMessage "Detected we are running on a 32-bit OS, so using normal CSCRIPT Engine"

		Set oExec = WshShell.Exec(WshShell.ExpandEnvironmentStrings("%windir%") & "\SYSTEM32\CSCRIPT.EXE /NOLOGO CustomDDR-Write.vbs " & runString)

	End If

	Set fso = Nothing

	Do While oExec.Status = 0

		WScript.Sleep 1000

	Loop

	LogMessage "Script returned with the following status " & Err.Number

End Sub

Function RetrieveProperty

' Here we determine if we're accessing the registry, WMI Environment or Task Sequence Environment and invoke the appropriate function to get the property, or alternatively we use a static value

If colNamedArguments.Exists("ForceValue") Then ' First we check for the override argument /ForceValue for the static value

	LogMessage "ForceValue detected so will return the specified value of " & colNamedArguments.Item("ForceValue")

	RetrieveProperty = colNamedArguments.Item("ForceValue")	

	If colNamedArguments.Exists("PropertyType") Then

		RetrieveProperty = RetrieveProperty & "&&&&" & colNamedArguments.Item("PropertyType")

		LogMessage "1: " & RetrieveProperty

	Else

		RetrieveProperty = RetrieveProperty & "&&&&String"

		LogMessage "2: " & RetrieveProperty

	End If

Else

	If colNamedArguments.Exists("Namespace") = False Then

		LogMessage "NameSpace not specified, defaulting to WMI and the root\CIMV2 namespace"

		RetrieveProperty = PropertyFromWMI("root\CIMV2",colNamedArguments.Item("Path"))

	Else

		Select Case UCase(colNamedArguments.Item("Namespace"))

			Case "HKLM"	

				LogMessage "Using Registry HKLM hive with path " & colNamedArguments.Item("Path")

				RetrieveProperty = PropertyFromRegistry(HKEY_LOCAL_MACHINE,colNamedArguments.Item("Path"))

			Case "HKCU"	

				LogMessage "Using Registry HKCU hive with path " & colNamedArguments.Item("Path")

				RetrieveProperty = PropertyFromRegistry(HKEY_CURRENT_USER,colNamedArguments.Item("Path"))


			Case "HKCR"	

				LogMessage "Using Registry HKCR hive with path " & colNamedArguments.Item("Path")

				RetrieveProperty = PropertyFromRegistry(HKEY_CLASSES_ROOT,colNamedArguments.Item("Path"))

			Case "HKU"	

				LogMessage "Using Registry HKU hive with path " & colNamedArguments.Item("Path")

				RetrieveProperty = PropertyFromRegistry(HKEY_USERS,colNamedArguments.Item("Path"))

			Case "HKCC"	

				LogMessage "Using Registry HKCC hive with path " & colNamedArguments.Item("Path")

				RetrieveProperty = PropertyFromRegistry(HKEY_CURRENT_CONFIG,colNamedArguments.Item("Path"))

			Case "ENV"	

				LogMessage "Using Environment Variable " & colNamedArguments.Item("Path")

				RetrieveProperty = PropertyFromEnvironment(colNamedArguments.Item("Path"))

			Case "TSENV"	

				LogMessage "Using Task Sequence Variable " & colNamedArguments.Item("Path")

				RetrieveProperty = PropertyFromTSEnvironment(colNamedArguments.Item("Path"))

			Case Else	

				LogMessage "Using WMI " & colNamedArguments.Item("Path")

				RetrieveProperty = PropertyFromWMI(colNamedArguments.Item("Namespace"),colNamedArguments.Item("Path"))

		End Select

	End If

	If RetrieveProperty = "" or isNull(RetrieveProperty) Then

		LogMessage "No value was returned, considered failure"

		WScript.Quit(6) ' Cannot obtain the property

	End If

End If

End Function

Function WhatIsIt(fpropertyValue)

	If isNumeric(fpropertyValue) Then 

		WhatIsIt = "Integer"

		LogMessage "PropertyTypeAuto specified, testing value " & fpropertyvalue & " and identified as Integer"

	Else

		WhatIsIt = "String"

		LogMessage "PropertyTypeAuto specified, testing value " & fpropertyvalue & " and identified as String"

	End If

End Function

Function PropertyFromEnvironment(fPath)

	PropertyFromEnvironment = wshShell.ExpandEnvironmentStrings("%" & fPath & "%")

	If PropertyFromEnvironment <> "" Then

		if colNamedArguments.Exists("PropertyTypeAuto") = TRUE then

			If colNamedArguments.Item("PropertyTypeAuto") = "TRUE" Then

				PropertyFromEnvironment = PropertyFromEnvironment & "&&&&" & WhatIsIt(PropertyFromEnvironment)

			End If

		Else

			PropertyFromEnvironment = PropertyFromEnvironment  & "&&&&" & "String"

		End If

		LogMessage "Environment value: " & PropertyFromEnvironment

	End If

End Function

Function PropertyFromTSEnvironment(fPath)

	On Error Resume Next

	Err.Clear

	Set TSEnv = CreateObject("Microsoft.SMS.TSEnvironment")

	If Err.Number <> 0 Then 

		LogMessage "Critical error registering the SMS Task Sequence Environment COM object - " & Err.Number & " - " & Err.Description

	Else 

		PropertyFromTSEnvironment = TSEnv(fPath)

		If PropertyFromTSEnvironment <> "" Then

			PropertyFromTSEnvironment = PropertyFromTSEnvironment & "&&&&String"

		End If
	
	End If

	LogMessage "Task Sequence value: " & PropertyFromTSEnvironment

	Set TSEnv = Nothing

End Function

Function PropertyFromRegistry(fnameSpace,fPath)

	On Error Resume Next

	strKeyProperty = MID(fPath, InStrRev(fPath,"\") + 1, LEN(fPath) - InStrRev(fPath,"\") + 1) ' The Path

	strKeyPath = MID(fPath, 1, InStrRev(fPath,"\")-1) ' The Property

	LogMessage "Key Property: " & strKeyProperty

	LogMessage "Key Path: " & strKeyPath

	LogMessage "Hive: " & fnameSpace

	Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

	oReg.EnumValues fnameSpace, strKeyPath, arrValueNames, arrValueTypes

	On Error Goto 0

	For I=0 To UBound(arrValueNames)

		If arrValueNames(I) = strKeyProperty Then

			Select Case arrValueTypes(I)
        	
			Case REG_SZ

				oReg.GetStringValue fnameSpace, strKeyPath, strKeyProperty, strValue

				LogMessage "Registry value is: " & strValue

				PropertyFromRegistry = strValue

				PropertyFromRegistry = PropertyFromRegistry & "&&&&String"

		        Case REG_EXPAND_SZ

				LogMessage "Data Type: Expanded String - Unsupported"

			Case REG_BINARY

				LogMessage "Data Type: Binary - Unsupported"

			Case REG_DWORD

				oReg.GetDWORDValue fnameSpace, strKeyPath, strKeyProperty, uValue

				LogMessage "Registry value is: " & uValue

				PropertyFromRegistry = uValue

				PropertyFromRegistry = PropertyFromRegistry & "&&&&Integer"

			Case REG_MULTI_SZ

				oReg.GetMultiStringValue fnameSpace, strKeyPath, strKeyProperty, arrValues				  				

				For Each strValue in arrValues

					LogMessage "Registry value is: " & strValue

					' Build an array out and pass it back

					PropertyFromRegistry = PropertyFromRegistry & strValue & "$$$$"

				Next

				PropertyFromRegistry = MID(PropertyFromRegistry,1,InStrRev(PropertyFromRegistry,"$$$$")-1)

				PropertyFromRegistry = PropertyFromRegistry & "&&&&Array"

			End Select 

		End If

	Next

End Function

Function PropertyFromWMI(fnameSpace,fPath)

	' Get a connection to the namespace on the local system.

	Set objWMIService = GetObject("winmgmts:\\.\" & fnameSpace)

	If Err.Number <> 0 Then 

			WScript.Quit(11) ' Invalid namespace

			LogMessage "Invalid WMI Namespace"

	End If

	' Get the specific class

	Err.Clear

	' Look for where clause and put it to the side

	If Instr(1, LCase(fPath), "where") Then

		LogMessage "WMI Where clause detected, attempting to separate out the elements without mangling the query"

		whereClause = LTRIM(RTRIM(MID(fPath,InStr(1,LCase(fPath),"where"),LEN(fPath)-InStr(1,LCase(fPath),"where")+1)))

		If InStr(1, whereClause,"=") Then ' Look for the where clauses value and surround it in qoutation marks to allow spaces

			LogMessage "WMI Where Clause operator = detected and will be used"

			whereValue = CHR(34) & LTRIM(RTRIM(MID(whereClause, InStr(1, whereClause,"=")+1, LEN(whereClause)-InStr(1, whereClause,"=")))) & CHR(34)

			whereClause = MID(whereClause, 1, InStr(1, whereClause,"=")) & " " & whereValue

		ElseIf InStr(1, whereClause,"<") Then

			LogMessage "WMI Where Clause operator < detected and will be used"

			whereValue = CHR(34) & LTRIM(RTRIM(MID(whereClause, InStr(1, whereClause,"<")+1, LEN(whereClause)-InStr(1, whereClause,"<")))) & CHR(34)

			whereClause = MID(whereClause, 1, InStr(1, whereClause,"<")) & " " & whereValue

		ElseIf InStr(1, whereClause,">") Then

			LogMessage "WMI Where Clause operator > detected and will be used"

			whereValue = CHR(34) & LTRIM(RTRIM(MID(whereClause, InStr(1, whereClause,">")+1, LEN(whereClause)-InStr(1, whereClause,">")))) & CHR(34)

			whereClause = MID(whereClause, 1, InStr(1, whereClause,">")) & " " & whereValue

		ElseIf InStr(1, LCase(whereClause),"like") Then

			LogMessage "WMI Where Clause operator LIKE detected and will be used and encased in qoute marks"

			whereValue = CHR(34) & LTRIM(RTRIM(MID(whereClause, InStr(1, LCase(whereClause),"like")+5, LEN(whereClause)-InStr(1, LCase(whereClause),"like")))) & CHR(34)

			whereClause = MID(whereClause, 1, InStr(1, LCase(whereClause),"like")+4) & " " & whereValue

		End If

		LogMessage "Where clause post-modification: " & whereClause

		fPath = LTRIM(RTRIM(MID(fPath,1,InStr(1,LCase(fPath),"where")-1)))

	End If

	LogMessage "Full WMI Query: " & "Select " & MID(fPath,instr(1, fPath,"\")+1,LEN(fPath)-instr(1, fPath,"\")) & " from " & MID(fPath,1,INSTR(1,fPath,"\")-1) & " " & whereClause

	On Error Resume Next

	Err.Clear

	set colClass = objWMIService.ExecQuery("Select " & MID(fPath,instr(1, fPath,"\")+1,LEN(fPath)-instr(1, fPath,"\")) & " from " & MID(fPath,1,INSTR(1,fPath,"\")-1) & " " & whereClause)

	LogMessage "Count of instances returned: " & colClass.Count 

	If Err.Number <> 0 Then

		LogMessage "We encountered a WMI error: " & Err.Number & " - " & Err.Description

		WScript.Quit(13) ' Invalid Query or WMI error

	End If

	On Error Goto 0

	If colClass.Count <> 0 Then

		' Loop through the available class objects and return the property

		For Each oClass in colClass

			For each oProperty in oClass.Properties_

				On Error Resume Next

				'LogMessage fPath
				'LogMessage "Property: " & oProperty
				LogMessage "WMI Property Type: " & oProperty.CIMType
				'LogMessage "Array1: " & oProperty.IsArray
				'LogMessage "Array2: " & IsArray(oProperty)

				If IsArray(oProperty) Then ' Array

					'LogMessage oProperty(1)

					For ii = 1 to UBOUND(oProperty)

						'LogMessage oProperty(ii)

						PropertyFromWMI = PropertyFromWMI & oProperty(ii) & "$$$$"

					Next

					PropertyFromWMI = MID(PropertyFromWMI,1,InStrRev(PropertyFromWMI,"$$$$")-1)

					wmiPropertyType = "Array"

				Else

					If oProperty.CIMType = CIM_REFERENCE Then ' Treat as String

						PropertyFromWMI = oProperty

						wmiPropertyType = "String"

					ElseIf oProperty.CIMType = CIM_STRING Then ' String

						PropertyFromWMI = oProperty

						wmiPropertyType = "String"

					ElseIf oProperty.CIMType = CIM_BOOLEAN or oProperty.CIMType = CIM_SINT8 or oProperty.CIMType = CIM_UINT8 or oProperty.CIMType = CIM_SINT16 Then ' Integer

						PropertyFromWMI = oProperty

						wmiPropertyType = "Integer"

					ElseIf oProperty.CIMType = CIM_DATETIME Then ' Convert Date to String

						Set dateTime = CreateObject("WbemScripting.SWbemDateTime")

						datetime.Value = oProperty

						PropertyFromWMI = datetime.GetVarDate

						Set dateTime = Nothing

						wmiPropertyType = "String"

					ElseIf  oProperty.CIMType = CIM_UINT16 or oProperty.CIMType = CIM_SINT32 or oProperty.CIMType = CIM_UINT32 or Property.CIMType = CIM_SINT64 or oProperty.CIMType = CIM_UINT64 Then ' Integer

						PropertyFromWMI = oProperty

						wmiPropertyType = "Integer"

					Else 'Treat anything else as a string
						PropertyFromWMI = oProperty

						wmiPropertyType = "String"

					End If

				End If

				On Error Goto 0

				Exit For ' Exit after processing the first returned object

			Next

			On Error Resume Next

			LogMessage "Class Name: " & oClass.Name
			LogMessage "Property: " & PropertyFromWMI

'			LogMessage "Tag: " & oClass.Tag

			On Error Goto 0

			If oProperty.Name = PropertyFromWMI Then ' The instance tag is being returned so replace with empty string, 		

				LogMessage "WMI returned the Tag instead of the Property"

				PropertyFromWMI = ""
				
				LogMessage "The WMI object's instance tag was returned instead of a value, is considered failure"

			End If

			Exit For

		Next

		If PropertyFromWMI <> "" Then PropertyFromWMI = PropertyFromWMI & "&&&&" & wmiPropertyType

	End If

End Function

Function DDRFileName ' Prepare DDR Filename

	Randomize

	If InStr(1,Date,"\",1) Then

		DDRFileName = ComputerName & "-" & REPLACE(Date,"\","-")

	ElseIf InStr(1,Date,"/",1) Then

		DDRFileName = ComputerName & "-" & REPLACE(Date,"/","-")

	End If

	DDRFileName = REPLACE(DDRFileName,".","-") & "-" & HEX(Int((100000 - 1 + 1) * Rnd + 1)) & ".DDR"

End Function

Sub LogMessage(logscriptError)

	On Error Resume Next

	Set fso = CreateObject("Scripting.FileSystemObject")

	Set logFile = fso.OpenTextFile(DDRDestPath & "\CustomDDR.LOG", ForAppending, True, tristateFalse)

	logFile.WriteLine Date & vbTab & vbTab & logscriptError	

	logfile.Close

	Set fso = Nothing

	On Error Goto 0

End Sub