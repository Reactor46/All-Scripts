' =================================================================================
' Get The List of All Domain User-Accounts Who Have an Alternate Recipient
' Usage: CScript /NoLogo AltRecipient.vbs > Result.txt
' =================================================================================

Option Explicit

Dim ObjConn, ObjRS, ObjRootDSE
Dim StrSQL, StrDomain

On Error Resume Next

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrDomain = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing
WScript.Echo vbNullString
StrSQL = "Select Name, AltRecipient From 'LDAP://" & StrDomain & "' Where ObjectCategory = 'User' And AltRecipient = '*'"
Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Provider = "ADsDSOObject":	ObjConn.Open "Active Directory Provider"
Set ObjRS = CreateObject("ADODB.Recordset")
ObjRS.Open StrSQL, ObjConn
If Not ObjRS.EOF Then
	ObjRS.MoveFirst
	WScript.Echo "User-Accounts Who Have an Alternate Recipient"
	WScript.Echo "================================================="
	While Not ObjRS.EOF
		WScript.Echo ObjRS.AbsolutePosition & vbTab & vbTab & "Name: " & Trim(ObjRS.Fields("Name").Value)
		GetProperName(Trim(ObjRS.Fields("AltRecipient").Value))
		WScript.Echo vbNullString
		ObjRS.MoveNext
	Wend
Else
	WScript.Echo "No User-Account in the domain has Alternate Recipient."
End If
ObjRS.Close:	Set ObjRS = Nothing
ObjConn.Close:	Set ObjConn = Nothing


Private Sub GetProperName(StrAlternate)
	Dim ObjUser
	Set ObjUser = GetObject("LDAP://" & StrAlternate)
	WScript.Echo vbTab & vbTab & "Alternate Recipient: " & Trim(ObjUser.Mail)
	Set ObjUser = Nothing
End Sub