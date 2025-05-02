' Disclaimer
'--------------------------------------------------------------------------------- 
' The sample scripts are not supported under any Microsoft standard support 
' program or service. The sample scripts are provided AS IS without warranty  
' of any kind. Microsoft further disclaims all implied warranties including,  
' without limitation, any implied warranties of merchantability or of fitness for 
' a particular purpose. The entire risk arising out of the use or performance of  
' the sample scripts and documentation remains with you. In no event shall 
' Microsoft, its authors, or anyone else involved in the creation, production, or 
' delivery of the scripts be liable for any damages whatsoever (including, 
' without limitation, damages for loss of business profits, business interruption, 
' loss of business information, or other pecuniary loss) arising out of the use 
' of or inability to use the sample scripts or documentation, even if Microsoft 
' has been advised of the possibility of such damages 
'--------------------------------------------------------------------------------- 

' Objective
'===================================================================
' This sample VBScript is to show how to find an OU in a given domain
' This sample binds to the root of the current domain to search for the OU
'===================================================================

Dim adoCommand
Dim adoConnection
Dim strFilter
Dim strAttributes
Dim resRecSet
Dim strSrchPath
Dim strDefNamCont
Dim strQuery
Dim strDN

' Setup ADO object for AD search.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

' bind to the root of domain - defaultNamingContext: http://technet.microsoft.com/en-us/library/ee156506.aspx
Set objRoot = GetObject("LDAP://rootDSE")
strDefNamCont = objRoot.Get("defaultnamingcontext")

' define ADO search object
strSrchPath = "<LDAP://" & strDefNamCont & ">"
' Filter to find out the OU in the domain
strFilter = "(&(objectClass=organizationalUnit)(OU=UserAccounts))"
' Comma delimited list of attribute values to retrieve.
strAttributes = "distinguishedName,name"
' Construct the LDAP syntax query.
strQuery = strSrchPath & ";" & strFilter & ";" & strAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 1
adoCommand.Properties("Cache Results") = False
' Run the query.
Set resRecSet = adoCommand.Execute

Do Until resRecSet.EOF
	' Retrieve values
	strDN = resRecSet.Fields("distinguishedName").Value
	strName = resRecSet.Fields("name").Value
	'Display the distingusishedName of the OU
	WScript.Echo strDN
Loop
