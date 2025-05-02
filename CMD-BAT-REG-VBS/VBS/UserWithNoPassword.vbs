' ========================================================================================================================
' Detect All User-Accounts Which Do Not Require To Have Password
' Contact: monimoys@hotmail.com

' Reference: Active Directory LDAP Syntax Filters
' >> https://social.technet.microsoft.com/wiki/contents/articles/5392.active-directory-ldap-syntax-filters.aspx
' ========================================================================================================================

Option Explicit

Const ADS_SCOPE_SUBTREE = 2
Dim ObjRootDSE, StrDomName, ObjConn, ObjRS, ObjCommand
Dim StrSQL, ObjUser, Counter

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing

StrSQL = "<LDAP://" & StrDomName & ">;(&(objectCategory=person)(objectClass=user)(userAccountControl:1.2.840.113556.1.4.803:=32));ADSPath;SubTree"

WScript.Echo VbCrLf & "Detecting All User-Accounts Which"
WScript.Echo Chr(34) & "Do Not Require To Have Password" & Chr(34) & " Setting Enabled"
WScript.Echo "Please Wait ..."
WScript.Echo "==================================================================="

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Provider = "ADsDSOObject":	ObjConn.Open "Active Directory Provider"
Set ObjCommand = CreateObject("ADODB.Command")
ObjCommand.ActiveConnection = ObjConn
ObjCommand.Properties("Page Size") = 1000:	ObjCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
ObjCommand.CommandText = StrSQL:	Set ObjRS = ObjCommand.Execute
If Not ObjRS.EOF Then
	ObjRS.MoveLast:	ObjRS.MoveFirst:	Counter = 0
	WScript.Echo "Total User-Accounts: " & Trim(ObjRS.RecordCount)
	WScript.Echo vbNullString
	While Not ObjRS.EOF
		Set ObjUser = GetObject(Trim(ObjRS.Fields("ADsPath").Value))
		If ObjUser.AccountDisabled = False Then
			Counter = Counter + 1
			WScript.Echo Trim(ObjRS.AbsolutePosition) & vbTab & "User ID: " & Trim(ObjUser.SAMAccountName) & " -- " & Trim(ObjUser.UserPrincipalName)
			WScript.Echo vbTab & "User's Name: " & Trim(ObjUser.DisplayName)
		End If
		Set ObjUser = Nothing:	WScript.Echo vbNullString
		ObjRS.MoveNext
	Wend
End If
WScript.Echo "==================================================================="
WScript.Echo "Total User-Accounts: " & Counter
WScript.Echo "==================================================================="
StrSQL = vbNullString:	Counter = 0
ObjRS.Close:	Set ObjRS = Nothing
If ObjCommand.State <> 0 Then
	ObjCommand.State = 0
End If
Set ObjCommand = Nothing
ObjConn.Close:	Set ObjConn = Nothing
WScript.Quit