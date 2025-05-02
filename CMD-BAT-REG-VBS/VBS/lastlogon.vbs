intBias = TimeZoneBias
arrAttributes = Array("lastLogonTimeStamp","sAMAccountName","mail")

Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

Set objRootDSE = GetObject("LDAP://RootDSE")
strBase = "<LDAP://" & objRootDSE.Get("defaultNamingContext") & ">"
Set objRootDSE = Nothing

strFilter = "(&(objectCategory=person)(objectClass=user))"
strAttributes = Join(arrAttributes,",")
Wscript.Echo Join(arrAttributes,";")
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False
Set adoRecordset = adoCommand.Execute
Do Until adoRecordset.EOF
	On Error Resume Next
	strTempOutput = ""
	For i = 1 To Ubound(arrAttributes)
		strTempOutput =  strTempOutput & " ; " & adoRecordset.Fields(arrAttributes(i)).Value
		strOutput = Mid(Ltrim(strTempOutput),3)
	Next
	Set objDate = adoRecordset.Fields(arrAttributes(0)).Value
	If (Err.Number <> 0) Then
        dtmDate = #1/1/1601#
    Else
		dtmDate = ((((objDate.Highpart * (2^32)) + objDate.LowPart)/(600000000 - intBias))/1440) + #1/1/1601#
	End If
	Set objDate = Nothing
	Wscript.Echo strOutput & " ; " & dtmDate
	adoRecordset.MoveNext
Loop
adoRecordset.Close
adoConnection.Close
Set adoRecordset = Nothing
Set adoConnection = Nothing
Set adoCommand = Nothing

Function TimeZoneBias
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colTimeZone = objWMIService.ExecQuery("Select * from Win32_TimeZone")
	For Each objTimeZone in colTimeZone
		TimeZoneBias = objTimeZone.Bias
	Next
	Set colTimeZone = Nothing
	Set objWMIService = Nothing
End Function
