Const ADS_PROPERTY_APPEND = 3
Set WshShell = WScript.CreateObject("WScript.Shell")
'----Get Computer DN------

Set objADSysInfo = CreateObject("ADSystemInfo")
ComputerDN = objADSysInfo.ComputerName
strcomputerdn = "LDAP://" & computerDN
Set objADSysInfo = Nothing

'----Connect AD-----

Set oRoot = GetObject("LDAP://rootDSE")
strDomainPath = oRoot.Get("defaultNamingContext")
Set oConnection = CreateObject("ADODB.Connection")
oConnection.Provider = "ADsDSOObject"
oConnection.Open "Active Directory Provider"

Count = WScript.Arguments.Count
For i = 0 To  count-1  	
	Group = WScript.Arguments(i)
	Addgroup Group
Next 

'----Get Group DN------
Function Addgroup(groupname)
	Set oRs = oConnection.Execute("SELECT adspath FROM 'LDAP://" & strDomainPath & "'" & "WHERE objectCategory='group' AND " & "Name='" & GroupName & "'")
	If Not oRs.EOF Then
		strAdsPath = oRs("adspath")
	End If
	If IsEmpty(strAdsPath) = False  Then 
		Const ADS_SECURE_AUTHENTICATION = 1
		Set objGroup = GetObject(stradspath) 
		Set objComputer = GetObject(strComputerDN)
		If (objGroup.IsMember(objComputer.AdsPath) = False) Then
			objGroup.PutEx ADS_PROPERTY_APPEND, "member", Array(computerdn)
			objGroup.SetInfo
		End If
	End If 
End Function