'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2009
'
' NAME: Unlock User accounts
'
' AUTHOR: Mohammed Alyafae , 
' DATE  : 9/22/2011
'
' COMMENT: this script is to search for locked accounts and unlock them 
'
'==========================================================================

Option Explicit
On Error Resume Next
Dim oQuery
Dim objConnection
Dim objCommand
Dim objRecordSet
Dim objUser
Dim objRoot
Dim NamingContext


set objRoot = getobject("LDAP://RootDSE")
NamingContext = objRoot.get("defaultNamingContext")
oQuery = "<LDAP://" & NamingContext & ">;" & "(objectClass=user);adspath;subtree"


'=======all the following lines are the same for every script====================
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Open "Provider=ADsDSOObject;"
objCommand.ActiveConnection = objConnection
objCommand.CommandText = oQuery
Set objRecordSet = objCommand.Execute
obj
'=================================================================================

While Not objRecordSet.EOF

Set objUser=GetObject(objRecordSet(0))

If objUser.Isaccountlocked=True Then
	WScript.Echo objRecordSet(0)
	objUser.IsAccountLocked=False
	objUser.SetInfo
End If
objRecordSet.MoveNext

Wend

objConnection.Close
Set objUser=Nothing

