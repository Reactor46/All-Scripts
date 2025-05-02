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

On Error Resume Next
Const HKEY_CLASSES_ROOT = &H80000000   
Const HKEY_CURRENT_USER = &H80000001   
Const HKEY_LOCAL_MACHINE = &H80000002   
Const HKEY_USERS = &H80000003   
Const HKEY_CURRENT_CONFIG = &H80000005 

Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7

'**********************************************************************************
' Function to set the IE TabProcGrowth property 
'**********************************************************************************
Function SetTabProcGrowth(siTabProcGrowth)
	Dim regValueType,reguRule,ieVersion,ieMajorVersion
	
	reguRule = "^[0-9]{1,}$"
	ieMajorVersion = 7
	ieVersion = GetIEVersion()
	
	If ieVersion <> "" Then 
		ieMajorVersion = Left(ieVersion,InStr(ieVersion,".")-1)
	End If 
	 
	If CInt(ieMajorVersion) >= 8 Then 
		If LCase(siTabProcGrowth) = "small" Then  
			SetRegistryValue HKEY_CURRENT_USER,"Software\Microsoft\Internet Explorer\Main","TabProcGrowth",siTabProcGrowth,1
			WScript.Echo "TabProcGrowth of the Internet Explorer was set to " & siTabProcGrowth & " successfuly."
			WScript.Echo "TabProcGrowth='small': Maximum 5 tab processes in a logon session, requires 15 tabs to get the 3rd tab process."
		ElseIf LCase(siTabProcGrowth) = "medium" Then
			SetRegistryValue HKEY_CURRENT_USER,"Software\Microsoft\Internet Explorer\Main","TabProcGrowth",siTabProcGrowth,1
			WScript.Echo "TabProcGrowth of the Internet Explorer was set to " & siTabProcGrowth & " successfuly."
			WScript.Echo "TabProcGrowth='medium': Maximum 9 tab processes in a logon session, requires 17 tabs to get the 5th tab process."
		ElseIf LCase(siTabProcGrowth) = "large" Then
			SetRegistryValue HKEY_CURRENT_USER,"Software\Microsoft\Internet Explorer\Main","TabProcGrowth",siTabProcGrowth,1
			WScript.Echo "TabProcGrowth of the Internet Explorer was set to " & siTabProcGrowth & " successfuly."
			WScript.Echo "TabProcGrowth='large': Maximum 16 tab processes in a logon session, requires 21 tabs to get the 9th tab process."
		ElseIf RegExpTest(reguRule,siTabProcGrowth) Then 
			If CInt(siTabProcGrowth) = 0 Then 
				SetRegistryValue HKEY_CURRENT_USER,"Software\Microsoft\Internet Explorer\Main","TabProcGrowth",siTabProcGrowth,4
				WScript.Echo "TabProcGrowth of the Internet Explorer was set to " & siTabProcGrowth & " successfuly."
				WScript.Echo "TabProcGrowth=0 : Tabs and frames run within the same process; frames are not unified across MIC(mandatory integrity) levels."
			ElseIf CInt(siTabProcGrowth) = 1 Then 
				SetRegistryValue HKEY_CURRENT_USER,"Software\Microsoft\Internet Explorer\Main","TabProcGrowth",siTabProcGrowth,4
				WScript.Echo "TabProcGrowth of the Internet Explorer was set to " & siTabProcGrowth & " successfuly."
				WScript.Echo "TabProcGrowth=1 : All tabs for a given frame process run in a single tab process for a given MIC(mandatory integrity) level." 
			ElseIf CInt(siTabProcGrowth) > 1 And CInt(siTabProcGrowth) <= 16 Then 
				SetRegistryValue HKEY_CURRENT_USER,"Software\Microsoft\Internet Explorer\Main","TabProcGrowth",siTabProcGrowth,4
				WScript.Echo "TabProcGrowth of the Internet Explorer was set to " & siTabProcGrowth & " successfuly."
				WScript.Echo "TabProcGrowth>1 : Multiple tab processes will be used to execute the tabs at a given MIC(mandatory integrity) level for a single frame process. In general, new processes are created until the TabProcGrowth number is met, and then tabs are load balanced across the tab processes." 
			Else
				WScript.Echo "Cannot validate argument on parameter 'TabProcGrowth'. The valid argument must be in ""small,medium,large"" or a number from 0 to 16. Please supply an argument that is in the set and then try the command again."
			End If
		Else 
			WScript.Echo "Cannot validate argument on parameter 'TabProcGrowth'. The valid argument must be in ""small,medium,large"" or a number from 0 to 16. Please supply an argument that is in the set and then try the command again."
		End If 
	Else
		WScript.Echo "'TabProcGrowth' only works on IE 8 or later."
	End If 
End Function 

'**********************************************************************************
' Function to get the IE TabProcGrowth property 
'**********************************************************************************
Function GetTabProcGrowth()
	Dim ieVersion,ieMajorVersion,reguRule
	
	reguRule = "^[0-9]{1,}$"
	ieMajorVersion = 7
	ieVersion = GetIEVersion()
	If ieVersion <> "" Then 
		ieMajorVersion = Left(ieVersion,InStr(ieVersion,".")-1)
	End If 
	
	If CInt(ieMajorVersion) >= 8 Then 
		If Not IsAValueInReg(HKEY_CURRENT_USER,"Software\Microsoft\Internet Explorer\Main","TabProcGrowth") Then 
			strTabProcGrowth = "TabProcGrowth Not Set : The context-based algorithm is used and the curve is chosen based on the amount of physical memory on the machine."
		Else 
			siIETabProcGrowth = ReadingValueInReg(HKEY_CURRENT_USER,"Software\Microsoft\Internet Explorer\Main","TabProcGrowth")
			If RegExpTest(reguRule, siIETabProcGrowth) Then 
				If CInt(siIETabProcGrowth) = 0 Then 
					strTabProcGrowth = "TabProcGrowth=0 : Tabs and frames run within the same process; frames are not unified across MIC(mandatory integrity) levels." 
				ElseIf CInt(siIETabProcGrowth) = 1 Then 
					strTabProcGrowth = "TabProcGrowth=1 : All tabs for a given frame process run in a single tab process for a given MIC(mandatory integrity) level."
				ElseIf CInt(siIETabProcGrowth) > 1 Then 
					strTabProcGrowth = "(TabProcGrowth=" & siIETabProcGrowth & ")TabProcGrowth>1 : Multiple tab processes will be used to execute the tabs at a given MIC(mandatory integrity) level for a single frame process. In general, new processes are created until the TabProcGrowth number is met, and then tabs are load balanced across the tab processes." 
				End If 
			Else
				If siIETabProcGrowth = "small" Then 
					strTabProcGrowth = "TabProcGrowth='small': Maximum 5 tab processes in a logon session, requires 15 tabs to get the 3rd tab process."
				ElseIf siIETabProcGrowth = "medium" Then 
					strTabProcGrowth = "TabProcGrowth='medium': (=Default Setting)Maximum 9 tab processes in a logon session, requires 17 tabs to get the 5th tab process."
				ElseIf siIETabProcGrowth = "large" Then 
					strTabProcGrowth = "TabProcGrowth='large': Maximum 16 tab processes in a logon session, requires 21 tabs to get the 9th tab process."
				End If 
			End If 
		End If 
	Else
		strTabProcGrowth = "'TabProcGrowth' only works on IE 8 or later."
	End If 
	
	WScript.Echo strTabProcGrowth
End Function 

'**********************************************************************************
' Function to get the IE Version
'**********************************************************************************
Function GetIEVersion()
	Dim ieFilePath,query,systemDrive,ieFileVersion
	Dim objShell,objFSO,objWMIService,objIEFile,objFileProperty,strComputer
	
	Set objShell = CreateObject("WScript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	strComputer = "."	
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	
	systemDrive = GetSystemDrive()
	ieFilePath = systemDrive & "\\Program Files\\Internet Explorer\\iexplore.exe"
	query = "Select * from CIM_Datafile Where name = '" & ieFilePath & "'"
	Set objIEFile = objWMIService.ExecQuery(query)

	For Each objFileProperty in objIEFile
		ieFileVersion = objFileProperty.Version
	Next 
	
	GetIEVersion = ieFileVersion
End Function

'***************************************************************************************
' Function to get the system Drive 
'***************************************************************************************
Function GetSystemDrive()
	Dim objWMIService, objItem, colItems
	Dim strSystemDrive
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
	For Each objItem in colItems
		strSystemDrive = objItem.SystemDrive
	Next
	GetSystemDrive = strSystemDrive
End Function 

'***********************************************************************************
' Function to set reigstry value
'***********************************************************************************
Function SetRegistryValue(root,strKeyPath,strValueName,strValue,strValueType)
	Dim strComputer
	Dim oReg
	strComputer = "."
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
	strComputer & "\root\default:StdRegProv")
	
	oReg.DeleteValue root,strKeyPath,strValueName,strValue
	
	If strValueType = REG_SZ Then 
		oReg.SetStringValue root,strKeyPath,strValueName,strValue
	ElseIf strValueType = REG_DWORD Then 
		oReg.SetDWORDValue root,strKeyPath,strValueName,strValue
	End If 
End Function 

'***********************************************************************************
' Function to reading value in Registry
'***********************************************************************************
Function ReadingValueInReg(root,strKeyPath,strValueName)
	On Error Resume Next 
	Dim strComputer,isCreateOK,rregResult
	Dim oReg
	
	strComputer = "."
	rregResult = ""
	
	Set oReg=	GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
						strComputer & "\root\default:StdRegProv")						
	isCreateOK=oReg.CreateKey (root,strKeyPath)
	intValueType = GetRegistryValueType(root,strKeyPath,strValueName)
	If Err.Number > 0 Then 
		rregResult = ""
	Else 
		If isCreateOK=0 Then  
			Select Case intValueType
				Case REG_SZ 
					oReg.GetStringValue root,strKeyPath & "\" ,strValueName,arrValues
					rregResult=arrValues
				Case REG_EXPAND_SZ
					oReg.GetExpandedStringValue root,strKeyPath & "\" ,strValueName,arrValues
					rregResult=arrValues
				Case REG_BINARY
					oReg.GetBinaryValue root,strKeyPath & "\" ,strValueName,arrValues
					rregResult= arrValues
				Case REG_DWORD 
					oReg.GetDWORDValue root,strKeyPath & "\" ,strValueName,arrValues
					rregResult=arrValues
				Case REG_MULTI_SZ
					oReg.GetMultiStringValue root,strKeyPath & "\" ,strValueName,arrValues
					rregResult= arrValues
				Case Else rregResult=""
			End Select
		End If
	End If   
	ReadingValueInReg=rregResult
End Function 

'**********************************************************************************
' Function to Get a registry value type
'**********************************************************************************
Function GetRegistryValueType(root,strKeyPath,strValueName)
	Dim strComputer,i,valueType
	Dim oReg
	
	strComputer = "."
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
	strComputer & "\root\default:StdRegProv")
	oReg.EnumValues root, strKeyPath, arrValueNames, arrValueTypes
	 
	For i=0 To UBound(arrValueNames)
	    If arrValueNames(i)= strValueName Then 
	    	valueType = arrValueTypes(i)
	    	Exit For 
	    End If 
	Next
	
	GetRegistryValueType = valueType
End Function 

'**********************************************************************************
' Function to check whether the Registry has a value in a key 
'**********************************************************************************
Function IsAValueInReg(root,strKeyPath,checkkey)
	On Error Resume Next
	 
	Dim result,strComputer,i
	Dim oReg,arrValueNames
	
	result=False 
	strComputer = "." 
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
							strComputer & "\root\default:StdRegProv")
	oReg.EnumValues root,strKeyPath,arrValueNames,arrValueTypes
	For i=0 To UBound(arrValueNames)  
	    oReg.GetStringValue root, strKeyPath,arrValueNames(i),strvalue  
	    If LCase(arrValueNames(i))=LCase(checkkey) Then
	    	If Err.Number = 0 Then 
		    	result=True 
		    	Exit For
	    	End If 
		End If 
	Next
	IsAValueInReg=result
End Function

'**********************************************************************************
' Regular Expression Check
' re = "^[0-9]{1,}.[0-9]{1,}.[0-9]{1,}.[0-9]{1,}$"
' MsgBox(RegExpTest(re, "12.01.12.22"))
' Reference: http://www.jxln.info/wangzhan-zhizuo/asp-vbscript-regexp.html
'**********************************************************************************
Function RegExpTest(patrn, strng)
	RetStr = False 
	Dim regEx, Match, Matches 
	Set regEx = New RegExp 
	regEx.Pattern = patrn 
	regEx.IgnoreCase = True 
	regEx.Global = True 
	Set Matches = regEx.Execute(strng) 
	For Each Match in Matches 
		If Match.Value <> "" Then 
			RetStr = True 
			Exit For 
		End If 
	Next
	RegExpTest = RetStr
End Function