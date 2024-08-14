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

Const NOMINIMUM = 0
Const REQUIRE_NTLM_V2_SESSION_SECURITY = 524288
Const REQUIRE_128BIT_ENCRRYPTION = 536870912

'**********************************************************************************
' Function to check Lync Server Role
'**********************************************************************************
Function CheckIsLyncEdgeorFrontEndServer()
	Dim strComputer
	Dim IsLyncEdgeorFrontEndServer
	Dim objWMIService,colServiceList,objservice,strServicesDisplayName
	strComputer = "."
	IsLyncEdgeorFrontEndServer = False 
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colServiceList = objWMIService.ExecQuery _
	    ("Select * from Win32_Service")
	For Each objservice in colServiceList
		strServicesDisplayName = LCase(objService.DisplayName) 
	    If InStr(strServicesDisplayName,"lync server") > 0 Then 
		    If InStr(strServicesDisplayName,"edge") > 0 Or  InStr(strServicesDisplayName,"front-end") > 0 Then
		        IsLyncEdgeorFrontEndServer = True 
		        Exit For 
		    End If
	    End If  
	Next
	CheckIsLyncEdgeorFrontEndServer = IsLyncEdgeorFrontEndServer
End Function 

'**********************************************************************************
' Function to change the NTLM setting to allow pre-Windows 7 clients from joining online meetings
'**********************************************************************************
Function FixtheIssueDescribedInKB982021()
	Dim intNtlmMinClientSec,intNtlmMinServerSec
	If(CheckIsLyncEdgeorFrontEndServer) Then 
		intNtlmMinClientSec = ReadingValueInReg(HKEY_LOCAL_MACHINE,"SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0","NtlmMinClientSec")
		intNtlmMinServerSec = ReadingValueInReg(HKEY_LOCAL_MACHINE,"SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0","NtlmMinServerSec")
		If (intNtlmMinClientSec = REQUIRE_128BIT_ENCRRYPTION Or intNtlmMinClientSec = REQUIRE_NTLM_V2_SESSION_SECURITY + REQUIRE_128BIT_ENCRRYPTION) Or  _ 
			(intNtlmMinServerSec = REQUIRE_128BIT_ENCRRYPTION Or intNtlmMinServerSec = REQUIRE_NTLM_V2_SESSION_SECURITY + REQUIRE_128BIT_ENCRRYPTION) Then 
			If intNtlmMinClientSec = REQUIRE_NTLM_V2_SESSION_SECURITY + REQUIRE_128BIT_ENCRRYPTION Then 
				SetRegistryValue HKEY_LOCAL_MACHINE,"SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0","NtlmMinClientSec",REQUIRE_NTLM_V2_SESSION_SECURITY,REG_DWORD
			ElseIf intNtlmMinClientSec = REQUIRE_128BIT_ENCRRYPTION Then 
				SetRegistryValue HKEY_LOCAL_MACHINE,"SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0","NtlmMinClientSec",NOMINIMUM,REG_DWORD
			End If 
			
			If intNtlmMinServerSec = REQUIRE_NTLM_V2_SESSION_SECURITY + REQUIRE_128BIT_ENCRRYPTION Then 
				SetRegistryValue HKEY_LOCAL_MACHINE,"SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0","NtlmMinServerSec",REQUIRE_NTLM_V2_SESSION_SECURITY,REG_DWORD
			ElseIf intNtlmMinServerSec = REQUIRE_128BIT_ENCRRYPTION Then 
				SetRegistryValue HKEY_LOCAL_MACHINE,"SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0","NtlmMinServerSec",NOMINIMUM,REG_DWORD
			End If
			
			WScript.Echo "The Server has been set to allow pre-Windows 7 clients from joining online meetings successfully."
		Else
			WScript.Echo "The Server was already matched setting to allow pre-Windows 7 clients from joining online meetings."
		End If 
	Else
		WScript.Echo "The Server is not a Lync Front End Server or an Lync Edge Server."
	End If
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
	Dim strComputer,isCreateOK,rregResult,intValueType,arrValues
	Dim oReg
	
	strComputer = "."
	rregResult = ""
	
	Set oReg=	GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
						strComputer & "\root\default:StdRegProv")						
	isCreateOK=oReg.CreateKey (root,strKeyPath)
	intValueType = GetRegistryValueType(root,strKeyPath,strValueName)
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
	
	ReadingValueInReg=rregResult
End Function 

'**********************************************************************************
' Function to Get a registry value type
'**********************************************************************************
Function GetRegistryValueType(root,strKeyPath,strValueName)
	Dim strComputer,i,arrValueNames, arrValueTypes,valueType
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
' Run Main function to fix the issue
'**********************************************************************************
FixtheIssueDescribedInKB982021()