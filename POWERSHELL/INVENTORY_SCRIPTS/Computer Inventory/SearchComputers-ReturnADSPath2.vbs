'=*=*=*=*=*=*=*=*=*=*=*=*=
' Created by Assaf Miron
' Date : 15/06/2008
' SearchComputers-ReturnADSPath.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*=

Function FindObject(strObj,ObjClass)
	Const ADS_SCOPE_SUBTREE = 2
	Dim objRootDSE,objConnection,objCommand,objRecordSet
	Dim strDomainLdap

	Set objRootDSE = GetObject ("LDAP://rootDSE")
	strDomainLdap  = objRootDSE.Get("defaultNamingContext")
	
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand = CreateObject("ADODB.Command")
	
	objConnection.Provider = "ADsDSOObject"
	objConnection.Open "Active Directory Provider"
	
	Set objCommand.ActiveConnection = objConnection
	
	objCommand.CommandText = _
		"SELECT AdsPath FROM 'LDAP://" & strDomainLdap & "' WHERE objectClass='" & ObjClass & "' and Name='" &_
			strObj & "'"
	
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Timeout") = 30
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
	objCommand.Properties("Cache Results") = False
	
	Set objRecordSet = objCommand.Execute
	
	If objRecordSet.RecordCount = 0 Then 
		FindObject= "No Computer Object Found"
	Else 	
		objRecordSet.Requery
		objRecordSet.MoveFirst
		Do Until objRecordSet.EOF
			FindObject= objRecordSet.Fields("AdsPath").Value
			objRecordSet.MoveNext
		Loop

	End If 	
End Function 

Wscript.echo FindObject(WScript.Arguments(0),"computer")