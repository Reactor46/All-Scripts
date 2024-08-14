' ====================================================================
' Get a List of All Locked-Out User Accounts in the Domain
' Usage: CScript /NoLogo CheckLock.vbs > C:\Result.txt
' ====================================================================

Option Explicit

Const ADS_UF_LOCKOUT = 16
Dim ObjRootDSE, ObjRS, ObjConn, ObjUser, LockFlag
Dim StrDomName, StrLDAPFilter, StrSQL

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing

StrLDAPFilter = "(&(SAMAccountType=805306368)(LockoutTime>=1))"
StrSQL = "<LDAP://" & StrDomName & ">;" & StrLDAPFilter & ";ADsPath;SubTree"

Set ObjConn = CreateObject("ADODB.Connection")
ObjConn.Provider = "ADsDSOObject":	ObjConn.Open "Active Directory Provider" 
Set ObjRS = CreateObject("ADODB.Recordset") 
ObjRS.Open StrSQL, ObjConn
If Not ObjRS.EOF Then
	ObjRS.MoveFirst
	WScript.Echo "Locked User Accounts are :"
	WScript.Echo "============================" & VbCrLf
	While Not ObjRS.EOF
		Set ObjUser = GetObject(Trim(ObjRS.Fields("ADsPath").Value))
		ObjUser.GetInfoEx Array("msDS-User-Account-Control-Computed"), 0
		LockFlag = ObjUser.Get("msDS-User-Account-Control-Computed")
		If (LockFlag and ADS_UF_LOCKOUT) Then
			WScript.Echo UCase(Trim(ObjUser.SAMAccountName)) & vbTab & vbTab & "---> " & Trim(ObjUser.UserPrincipalName)
		End If
		ObjRS.MoveNext
	Wend
Else
	WScript.Echo "There is No Locked User Account in the Domain."
	WScript.Echo vbNullString	
End If
ObjRS.Close:	Set ObjRS = Nothing
ObjConn.Close:	Set ObjConn = Nothing