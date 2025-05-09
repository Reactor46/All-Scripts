'Script deletes Users from a csv file.
'csv format is strsAMUserName,Whatever
'Written by Andrew hill and Carl Harrison - Microsoft MCS
'this script is offered with no warranty
'On Error Resume Next 'used in case user not found
Option Explicit

Const ForReading = 1

Dim strL, spl1, strOU, strUserCN, strName
Dim objFSO, objInputFile 

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objInputFile = objFSO.OpenTextFile(".\users.csv", ForReading) 'your csv file

wscript.echo "script started"

'extract from csv file
Do until objInputFile.AtEndOfStream
	strL = objInputFile.ReadLine
	spl1 = Split(strL, ",")
	strName = (spl1(0))
	If UserExists(strName) = True Then
		'WScript.Echo strName & " exists."
		DelUser
	End If			
Loop

Set objFSO = Nothing
Set objInputFile = Nothing

wscript.echo "script finished"


'user exist check
Function UserExists(strsAMUserName) 

Dim strDNSDomain, strFilter, strQuery
Dim objConnection, objCommand, objRootLDAP, objLDAPUser, objRecordSet


UserExists = False
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
Set objRootLDAP = GetObject("LDAP://RootDSE")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
'objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

strDNSDomain = objRootLDAP.Get("DefaultNamingContext")
strFilter = "(&(objectCategory=user)(sAMAccountName=" & strsAMUserName & "))"

strQuery = "<LDAP://" & strDNSDomain & ">;" & strFilter & ";sAMAccountName,adspath,CN;subTree"

objCommand.CommandText = strQuery
'WScript.Echo strFilter
'WScript.Echo strQuery
Set objRecordSet = objCommand.Execute

If objRecordSet.RecordCount = 1 Then

objRecordSet.MoveFirst
    'WScript.Echo "We got here " & strsAMGroupName      
	'WScript.Echo objRecordSet.Fields("sAMAccountname").Value
	'WScript.Echo objRecordSet.Fields("adspath").Value
	If objRecordSet.Fields("sAMAccountname").Value = strsAMUserName Then
		UserExists = True
		Set objLDAPUser = GetObject(objRecordSet.Fields("adspath").Value)
		strOU = objLDAPUser.Parent
		strUserCN = objRecordSet.Fields("CN").Value
	End If
Else
	WScript.Echo strsAMUserName & " User doesn't exist or Duplicate sAMAccountName"
	UserExists = False
	strUserCN = ""
	strOU = ""
End If

objRecordSet.Close
Set objConnection = Nothing
Set objCommand = Nothing
Set objRootLDAP = Nothing
Set objLDAPUser = Nothing
Set objRecordSet = Nothing

end function


Sub DelUser

Dim objOU

'WScript.Echo strOU
'WScript.Echo strGroupCN
Set objOU = GetObject(strOU)
objOU.Delete "User", "cn=" & strUserCN & ""
WScript.Echo strName & " (CN=" & strUserCN & ") has been deleted."

Set ObjOU = Nothing
strUserCN = ""

End Sub