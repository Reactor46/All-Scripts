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

Function GetOSCADOUDN(Name,ObjectCategory)
	'This function is used to get distinguishedName of an orgnizational unit.
	Dim objConnection,objCommand,objRecordSet,objRootDSE,strDefaultNamingContext
	Set objRootDSE = GetObject("LDAP://RootDSE")
	strDefaultNamingContext = objRootDSE.Get("defaultNamingContext")
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = _
	"<LDAP://" & strDefaultNamingContext & _
	">;(&(objectCategory=" & ObjectCategory & ")(name=" & Name & _
	"));distinguishedName;subtree"
	Set objRecordSet = objCommand.Execute
	If objRecordSet.RecordCount = 1 Then
		While Not objRecordSet.EOF
			GetOSCADOUDN = objRecordSet.Fields("distinguishedName")
			objRecordSet.MoveNext
		Wend
	Else
		GetOSCADOUDN = Null
	End If
	objConnection.Close
End Function

Function GetOSCADComputerName(OUDN,SearchScope)
	'This function is used to get computer names from one orgnizational unit. 
	Dim objConnection,objCommand,objRecordSet,objRootDSE,strDefaultNamingContext
	Dim arrNames(),i
	Set objRootDSE = GetObject("LDAP://RootDSE")
	strDefaultNamingContext = objRootDSE.Get("defaultNamingContext")
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = _
	"<LDAP://" & OUDN & _
	">;(&(sAMAccountType=805306369)(name=*));name;" & SearchScope
	Set objRecordSet = objCommand.Execute
	While Not objRecordSet.EOF
		ReDim Preserve arrNames(i)
		arrNames(i) = objRecordSet.Fields("name")
		objRecordSet.MoveNext
		i = i + 1
	Wend
	objConnection.Close
	GetOSCADComputerName = arrNames
End Function

Function SetOSCDNSServerSearchOrderForSingleNIC(ComputerName,NetConnectionID,OldDNSServerSearchOrder,NewDNSServerSearchOrder)
	'This function is used to modify the DNS server search order for a NIC with specified name.
	On Error Resume Next
	Dim objWMILocator,objWMIService,colItems,objItem,strOriginalDNServers
	Dim intNICIndex,objNIC,blnNeedModify,i,intReturnValue,arrOldDNServers,objInParams,objOutParams
	Dim arrReportItem(4)
	'Connect to remote computer
	Set objWMILocator = CreateObject("WbemScripting.SWbemLocator")
	Set objWMIService = objWMILocator.ConnectServer(ComputerName,"root\cimv2")
	'If error occured, generate the report item with error description.
	If Err.Number <> 0 Then
		arrReportItem(0) = ComputerName
		arrReportItem(1) = NetConnectionID
		arrReportItem(2) = Err.Description
		For i = 3 To 4
			arrReportItem(i) = "N/A"
		Next
		Err.Clear
	Else
		'Find the NIC with specified name by using WQL.
		Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapter Where NetConnectionID='" & NetConnectionID & "'")
		'If cannot find the NIC with specified name, generate the report item with a friendly message.
		If colItems.Count <> 0 Then
			For Each objItem In colItems
				'Populate report item
				arrReportItem(0) = ComputerName
				arrReportItem(1) = NetConnectionID
				intNICIndex = objItem.Index
				'Check DNS configuration for the specified NIC
				Set objNIC = objWMIService.Get("Win32_NetworkAdapterConfiguration.Index='" & intNICIndex & "'")
				If objNIC.Properties_.Item("IPEnabled").Value And Not objNIC.Properties_.Item("DHCPEnabled").Value Then
					'Get current DNS search order and convert it to a string.
					strOriginalDNServers = Join(objNIC.Properties_.Item("DNSServerSearchOrder"))
					'Compare to the user input, if current configuration does not meet the requirements, set a flag.
					If InStr(OldDNSServerSearchOrder,",") = 0 Then
						If InStr(strOriginalDNServers,OldDNSServerSearchOrder) > 0 Then
							blnNeedModify = True
						Else
							blnNeedModify = False
						End If
					Else
						arrOldDNServers = Split(OldDNSServerSearchOrder,",")
						For i = 0 To UBound(arrOldDNServers)
							If InStr(strOriginalDNServers,arrOldDNServers(i)) > 0 Then
								blnNeedModify = True
								Exit For
							Else
								blnNeedModify = False
							End If
						Next
					End If
				Else
					blnNeedModify = False
				End If
				'Begin to modify the DNS Server Search Order, and save the result to the report
				If blnNeedModify Then
					Set objInParams = objNIC.Methods_.Item("SetDNSServerSearchOrder").InParameters.SpawnInstance_()
					objInParams.Properties_.Item("DNSServerSearchOrder") = Split(NewDNSServerSearchOrder,",")
					Set objOutParams = objNIC.ExecMethod_("SetDNSServerSearchOrder",objInParams)
					If objOutParams.ReturnValue = 0 Then
						arrReportItem(2) = "Modified"
					Else
						arrReportItem(2) = "Failed to modify DNS server search order. Return Value: " & CStr(objOutParams.ReturnValue)
					End If
				Else
					arrReportItem(2) = "Not Modified"
				End If
				arrReportItem(3) = strOriginalDNServers
				If arrReportItem(2) = "Not Modified" Then
					arrReportItem(4) = "N/A"
				Else
					arrReportItem(4) = NewDNSServerSearchOrder
				End If
			Next
		Else
			arrReportItem(0) = ComputerName
			arrReportItem(1) = NetConnectionID
			arrReportItem(2) = "Cannot find a network adaptor with specified name."
			arrReportItem(3) = "N/A"
			arrReportItem(4) = "N/A"
		End If
	End If
	'Return the report
	SetOSCDNSServerSearchOrderForSingleNIC = Chr(34) & Join(arrReportItem,""",""") + Chr(34)
End Function

Function SetOSCDNSServerSearchOrderForMultipleNIC(ComputerName,OldDNSServerSearchOrder,NewDNSServerSearchOrder)
	'This function is used to modify the DNS server search order for multiple NICs.
	On Error Resume Next
	Dim objWMILocator,objWMIService,colItems,objItem,strOriginalDNServers
	Dim intNICIndex,colNics,objNIC,blnNeedModify,i,intReturnValue,arrOldDNServers,objInParams,objOutParams
	Dim arrReportItem(4),strReport
	Dim arrReportRows(),intReportRowCounter
	Set objWMILocator = CreateObject("WbemScripting.SWbemLocator")
	Set objWMIService = objWMILocator.ConnectServer(ComputerName,"root\cimv2")
	If Err.Number <> 0 Then
		arrReportItem(0) = ComputerName
		arrReportItem(1) = "N/A"
		arrReportItem(2) = Err.Description
		For i = 3 To 4
			arrReportItem(i) = "N/A"
		Next
		ReDim Preserve arrReportRows(intReportRowCounter)
		arrReportRows(intReportRowCounter) = Chr(34) & Join(arrReportItem,""",""") & Chr(34)
		intReportRowCounter = intReportRowCounter + 1
		Err.Clear
	Else
		Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled=True And DHCPEnabled=False")
		If colItems.Count <> 0 Then
			For Each objItem In colItems
				arrReportItem(0) = ComputerName
				intNICIndex = objItem.Index
				'Check DNS configuration for specfied NIC
				Set colNICs = objWMIService.ExecQuery("Select * From Win32_NetworkAdapter Where Index='" & intNICIndex & "'")
				For Each objNIC In colNics
					arrReportItem(1) = objNIC.Properties_.Item("NetConnectionID")
				Next
				strOriginalDNServers = Join(objItem.Properties_.Item("DNSServerSearchOrder"),",")
				If InStr(OldDNSServerSearchOrder,",") = 0 Then
					If InStr(strOriginalDNServers,OldDNSServerSearchOrder) > 0 Then
						blnNeedModify = True
					Else
						blnNeedModify = False
					End If
				Else
					arrOldDNServers = Split(OldDNSServerSearchOrder,",")
					For i = 0 To UBound(arrOldDNServers)
						If InStr(strOriginalDNServers,arrOldDNServers(i)) > 0 Then
							blnNeedModify = True
							Exit For
						Else
							blnNeedModify = False
						End If
					Next
				End If
				If blnNeedModify Then
					Set objInParams = objItem.Methods_.Item("SetDNSServerSearchOrder").InParameters.SpawnInstance_()
					objInParams.Properties_.Item("DNSServerSearchOrder") = Split(NewDNSServerSearchOrder,",")
					Set objOutParams = objItem.ExecMethod_("SetDNSServerSearchOrder",objInParams)
					If objOutParams.ReturnValue = 0 Then
						arrReportItem(2) = "Modified"
					Else
						arrReportItem(2) = "Failed to modify DNS serrver search order. Return Value: " & CStr(objOutParams.ReturnValue)
					End If
				Else
					arrReportItem(2) = "Not Modified"
				End If
				arrReportItem(3) = strOriginalDNServers
				If arrReportItem(2) = "Not Modified" Then
					arrReportItem(4) = "N/A"
				Else
					arrReportItem(4) = NewDNSServerSearchOrder
				End If
				ReDim Preserve arrReportRows(intReportRowCounter)
				arrReportRows(intReportRowCounter) = Chr(34) & Join(arrReportItem,""",""") & Chr(34)
				intReportRowCounter = intReportRowCounter + 1
			Next
		Else
			arrReportItem(0) = ComputerName
			arrReportItem(1) = "N/A"
			arrReportItem(2) = "Cannot find a network adaptor with IP address enabled."
			arrReportItem(3) = "N/A"
			arrReportItem(4) = "N/A"
			ReDim Preserve arrReportRows(intReportRowCounter)
			arrReportRows(intReportRowCounter) = Chr(34) & Join(arrReportItem,""",""") & Chr(34)
			intReportRowCounter = intReportRowCounter + 1
		End If
	End If
	strReport = Join(arrReportRows,vbCrLf)
	SetOSCDNSServerSearchOrderForMultipleNIC = strReport
End Function

Function GetOSCDNSServerSearchOrderForMultipleNIC(ComputerName)
	'This function is used to retrieve the DNS server search order.
	On Error Resume Next
	Dim objWMILocator,objWMIService,colItems,objItem,strOriginalDNServers
	Dim intNICIndex,colNics,objNIC,i,intReturnValue,arrOldDNServers,objInParams,objOutParams
	Dim arrReportItem(3),strReport
	Dim arrReportRows(),intReportRowCounter
	Set objWMILocator = CreateObject("WbemScripting.SWbemLocator")
	Set objWMIService = objWMILocator.ConnectServer(ComputerName,"root\cimv2")
	If Err.Number <> 0 Then
		arrReportItem(0) = ComputerName
		arrReportItem(1) = "N/A"
		arrReportItem(2) = Err.Description
		arrReportItem(3) = "N/A"
		ReDim Preserve arrReportRows(intReportRowCounter)
		arrReportRows(intReportRowCounter) = Chr(34) & Join(arrReportItem,""",""") & Chr(34)
		intReportRowCounter = intReportRowCounter + 1
		Err.Clear
	Else
		Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled=True And DHCPEnabled=False")
		If colItems.Count <> 0 Then
			For Each objItem In colItems
				arrReportItem(0) = ComputerName
				intNICIndex = objItem.Index
				Set colNICs = objWMIService.ExecQuery("Select * From Win32_NetworkAdapter Where Index='" & intNICIndex & "'")
				For Each objNIC In colNics
					arrReportItem(1) = objNIC.Properties_.Item("NetConnectionID")
				Next
				strOriginalDNServers = Join(objItem.Properties_.Item("DNSServerSearchOrder"),",")
				arrReportItem(2) = "OK"
				arrReportItem(3) = strOriginalDNServers
				ReDim Preserve arrReportRows(intReportRowCounter)
				arrReportRows(intReportRowCounter) = Chr(34) & Join(arrReportItem,""",""") & Chr(34)
				intReportRowCounter = intReportRowCounter + 1
			Next
		Else
			arrReportItem(0) = ComputerName
			arrReportItem(1) = "N/A"
			arrReportItem(2) = "Cannot find a network adaptor with IP address enabled."
			arrReportItem(3) = "N/A"
			ReDim Preserve arrReportRows(intReportRowCounter)
			arrReportRows(intReportRowCounter) = Chr(34) & Join(arrReportItem,""",""") & Chr(34)
			intReportRowCounter = intReportRowCounter + 1
		End If
	End If
	strReport = Join(arrReportRows,vbCrLf)
	GetOSCDNSServerSearchOrderForMultipleNIC = strReport
End Function

Sub OSCScriptUsage
	WScript.Echo "How to use this script:" & vbCrLf & vbCrLf _
	& "1. Logon to one utility server with an administrative account." & vbCrLf _
	& "2. Run one of the following commands:" & vbCrLf & vbCrLf  _
	& "cscript //nologo SetDNSServerSearchOrder.vbs /OUName:""OUName"" /NICName:""Local Area Connection"" " _
	& "/OldDNSServerSearchOrder:""W.X.Y.Z,W.X.Y.Z"" /NewDNSServerSearchOrder:""W.X.Y.Z,W.X.Y.Z""" _
	& vbCrLf & vbCrLf _
	& "cscript //nologo SetDNSServerSearchOrder.vbs /OUName:""OUName""" _
	& " /OldDNSServerSearchOrder:""W.X.Y.Z,W.X.Y.Z"" /NewDNSServerSearchOrder:""W.X.Y.Z,W.X.Y.Z""" _
	& vbCrLf & vbCrLf _
	& "cscript //nologo SetDNSServerSearchOrder.vbs /OUName:""OUName""" _
	& " /Retrieve"
End Sub

Sub Main
	Dim i,strOUName,strOUDN,arrComputerNames,strNICName,strOldDNSServerSearchOrder,strNewDNSServerSearchOrder,strReport
	Dim arrReportColumnHead,objRegExp,arrDNSServers,blnRetrieve,objArgs
	Set objArgs = WScript.Arguments
	Set objRegExp = New RegExp
	objRegExp.Global = True
	objRegExp.IgnoreCase = True
	If InStr(WScript.FullName,"cscript") = 0 Then
		Call OSCScriptUsage()
		WScript.Quit(1)
	End If	
	'Verify Arguments
	objRegExp.Pattern = "ouname|nicname|olddnsserversearchorder|newdnsserversearchorder|retrieve"
	If objArgs.Named.Count < 2 Or objArgs.Named.Count > 5 Then
		Call OSCScriptUsage()
		WScript.Quit(1)
	Else
		For i = 0 To objArgs.Count - 1
			If Not objRegExp.Test(objArgs(i)) Then
				Call OSCScriptUsage()
				WScript.Quit(1)
			Else
				With objArgs.Named
					If .Exists("ouname") Then strOUName = .Item("ouname")
					If .Exists("nicname") Then strNICName = .Item("nicname")
					If .Exists("olddnsserversearchorder") Then strOldDNSServerSearchOrder = .Item("olddnsserversearchorder")
					If .Exists("newdnsserversearchorder") Then strNewDNSServerSearchOrder = .Item("newdnsserversearchorder")
					If .Exists("retrieve") Then blnRetrieve = True
				End With
			End If
		Next
	End If
	If IsNull(GetOSCADOUDN(strOUName,"organizationalUnit")) Then
		WScript.Echo "Please use a valid organizational unit name."
		WScript.Quit(1)
	Else
		strOUDN = GetOSCADOUDN(strOUName,"organizationalUnit")
	End If
	If Not blnRetrieve Then
		'Ensure IP address is valid.
		objRegExp.Pattern = "^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$"
		If InStr(strOldDNSServerSearchOrder,",") > 0 Then
			arrDNSServers = Split(strOldDNSServerSearchOrder,",")
			For i = 0 To UBound(arrDNSServers)
				If Not objRegExp.Test(arrDNSServers(i)) Then
					WScript.Echo "Please use a valid IP address."
					WScript.Quit(1)
				End If
			Next
		Else
			If Not objRegExp.Test(strOldDNSServerSearchOrder) Then
				WScript.Echo "Please use a valid IP address."
				WScript.Quit(1)
			End If	
		End If
		If InStr(strNewDNSServerSearchOrder,",") > 0 Then
			arrDNSServers = Split(strNewDNSServerSearchOrder,",")
			For i = 0 To UBound(arrDNSServers)
				If Not objRegExp.Test(arrDNSServers(i)) Then
					WScript.Echo "Please use a valid IP address."
					WScript.Quit(1)
				End If
			Next
		Else
			If Not objRegExp.Test(strNewDNSServerSearchOrder) Then
				WScript.Echo "Please use a valid IP address."
				WScript.Quit(1)
			End If	
		End If
	End If
	'Run functions
	arrComputerNames = GetOSCADComputerName(strOUDN,"subtree")
	For i = 0 To UBound(arrComputerNames)
		If Not blnRetrieve Then
			If strNICName <> "" Then
				strReport = strReport + _
				SetOSCDNSServerSearchOrderForSingleNIC(arrComputerNames(i),strNICName,strOldDNSServerSearchOrder,strNewDNSServerSearchOrder) _
				+ vbCrLf
			Else
				strReport = strReport + _
				SetOSCDNSServerSearchOrderForMultipleNIC(arrComputerNames(i),strOldDNSServerSearchOrder,strNewDNSServerSearchOrder) _
				+ vbCrLf
			End If
		Else
			strReport = strReport + GetOSCDNSServerSearchOrderForMultipleNIC(arrComputerNames(i)) + vbCrLf
		End If
	Next
	If Not blnRetrieve Then
		arrReportColumnHead = Array("ComputerName","NIC Name","Status", _
		"DNSServerSearchOrder(Before)","DNSServerSearchOrder(After)")
	Else
		arrReportColumnHead = Array("ComputerName","NIC Name","Status","DNSServerSearchOrder")
	End If
	WScript.Echo Chr(34) + Join(arrReportColumnHead,""",""") + Chr(34)		
	WScript.Echo strReport	
End Sub

Call Main()