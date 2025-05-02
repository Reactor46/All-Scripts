'Script deletes security groups from a csv file.
'csv format is strsAMGroupName,Whatever
'Written by Andrew hill and Carl Harrison - Microsoft MCS
'this script is offered with no warranty
'On Error Resume Next 'used in case group not found
Option Explicit

Const ForReading = 1

Dim strL, spl1, strOU, strGroupCN, strGroupName
Dim objFSO, objInputFile 

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objInputFile = objFSO.OpenTextFile(".\groups.csv", ForReading) 'your csv file

wscript.echo "script started"

'extract from csv file
Do until objInputFile.AtEndOfStream
	strL = objInputFile.ReadLine
	spl1 = Split(strL, ",")
	strGroupName = (spl1(0))
	If GroupExists(strGroupName) = True Then
		'WScript.Echo strGroupName & " exists."
		DelGroup
	End If			
Loop

Set objFSO = Nothing
Set objInputFile = Nothing

wscript.echo "script finished"


'group exist check
Function GroupExists(strsAMGroupName) 

Dim strDNSDomain, strFilter, strQuery
Dim objConnection, objCommand, objRootLDAP, objLDAPGroup, objRecordSet


GroupExists = False
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
Set objRootLDAP = GetObject("LDAP://RootDSE")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
'objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

strDNSDomain = objRootLDAP.Get("DefaultNamingContext")
strFilter = "(&(objectCategory=group)(sAMAccountName=" & strsAMGroupName & "))"

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
	If objRecordSet.Fields("sAMAccountname").Value = strsAMGroupName Then
		GroupExists = True
		Set objLDAPGroup = GetObject(objRecordSet.Fields("adspath").Value)
		strOU = objLDAPGroup.Parent
		strGroupCN = objRecordSet.Fields("CN").Value
	End If
Else
	WScript.Echo strsAMGroupName & " Group doesn't exist or Duplicate sAMAccountName"
	GroupExists = False
	strGroupCN = ""
	strOU = ""
End If

objRecordSet.Close
Set objConnection = Nothing
Set objCommand = Nothing
Set objRootLDAP = Nothing
Set objLDAPGroup = Nothing
Set objRecordSet = Nothing

end function


Sub DelGroup

Dim objOU

'WScript.Echo strOU
'WScript.Echo strGroupCN
Set objOU = GetObject(strOU)
objOU.Delete "Group", "cn=" & strGroupCN & ""
WScript.Echo strGroupName & " (CN=" & strGroupCN & ") has been deleted."

Set ObjOU = Nothing
strGroupCN = ""

End Sub