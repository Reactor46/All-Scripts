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

'**********************************************************************************
' Function to get the total count of printers hosted on the server.
'**********************************************************************************
Function GetPinterCountTheServerHosted()
	On Error Resume Next 
	Dim strComputer
	Dim oReg,oSubReg,arrSubKeys
	Dim arrSubKeys1,subkey
	Dim PinterCountTheServerHosted
	Dim IsFailOverCluster
	Dim ResourceType, ResourceState
	
	strComputer = "."
	PinterCountTheServerHosted = 0
	IsFailOverCluster = False 
	
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
							strComputer & "\root\default:StdRegProv")
													
	oReg.EnumKey HKEY_LOCAL_MACHINE, "Cluster\Resources", arrSubKeys

	If arrSubKeys <> Null Then 
		For Each subkey In arrSubKeys
	    	If IsAKeyInReg(HKEY_LOCAL_MACHINE,"Cluster\Resources\" & subkey,"Printers") Then
	    		ResourceType = ReadingValueInReg(HKEY_LOCAL_MACHINE,"Cluster\Resources\" & subkey,"Type")
	    		ResourceState = ReadingValueInReg(HKEY_LOCAL_MACHINE,"Cluster\Resources\" & subkey,"PersistentState")
	    		If ResourceType = "Print Spooler" Then 
		    		If ResourceState = "2" Then 
			    		IsFailOverCluster = True 
				    	oReg.EnumKey HKEY_LOCAL_MACHINE, "Cluster\Resources\" & subkey &"\Printers", arrSubKeys1 
				    	If arrSubKeys1 <> Null Then 
				    		PinterCountTheServerHosted = PinterCountTheServerHosted + UBound(arrSubKeys1) + 1
				    	End If
				    End If 
		    	End If  
			End If 
		Next
	End If 
	
	If Not IsFailOverCluster Then 
		oReg.EnumKey HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Print\Printers", arrSubKeys
		If arrSubKeys <> Null Then  
			PinterCountTheServerHosted = PinterCountTheServerHosted + UBound(arrSubKeys) + 1
		End If 
	End If 
	
	WScript.Echo "The total number of the printers hosted on this Server: " & PinterCountTheServerHosted
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
	Dim strComputer,i,valueType,arrValueNames,arrValueTypes
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
	Dim oReg,arrValueNames,arrValueTypes
	
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
'check whether the Registry has a key 
'**********************************************************************************
Function IsAKeyInReg(root,strKeyPath,checkkey)
	On Error Resume Next 
	
	Dim result,strComputer,i
	Dim oReg,arrSubKeys,subkey
	
	result=False  
	strComputer = "." 
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
							strComputer & "\root\default:StdRegProv")
	oReg.EnumKey root, strKeyPath, arrSubKeys
	For Each subkey In arrSubKeys
    	If LCase(subkey)=LCase(checkkey) Then
	    	If Err.Number = 0 Then 
		    	result=True 
		    	Exit For
	    	End If 
		End If 
	Next
	IsAKeyInReg=result
End Function

GetPinterCountTheServerHosted