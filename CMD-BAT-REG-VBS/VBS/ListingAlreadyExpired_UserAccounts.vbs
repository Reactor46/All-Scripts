Option Explicit

Const ADS_SCOPE_SUBTREE = 2
Dim ObjConn, ObjRS, ObjCommand, ObjRootDSE, ObjUser, StrSQL
Dim StrDomName, StrAccntReturn, DTMAccountExpiration, CurrLocale

On Error Resume Next

CurrLocale = GetLocale():	SetLocale(2057)

Set ObjRootDSE = GetObject("LDAP://RootDSE")
StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing

StrSQL = "<LDAP://" & StrDomName & ">;(sAMAccountType=805306368);ADsPath;SubTree"

Set ObjConn =  CreateObject("ADODB.Connection")
ObjConn.Provider = "ADsDSOObject":	ObjConn.Open "Active Directory Provider"
Set ObjCommand = CreateObject("ADODB.Command"):	ObjCommand.ActiveConnection = ObjConn
ObjCommand.Properties("Page Size") = 1000:	ObjCommand.Properties("SearchScope") = ADS_SCOPE_SUBTREE
ObjCommand.Properties("Cache Results") = False:	ObjCommand.CommandText = StrSQL
Set ObjRS = ObjCommand.Execute

If Not ObjRS.EOF Then
	ObjRS.MoveFirst
	WScript.Echo
	WScript.Echo "Listing Already Expired User-Accounts"
	WScript.Echo "---------------------------------------"
	While Not ObjRS.EOF
		Set ObjUser = GetObject(Trim(ObjRS.Fields("ADsPath").Value))
		DTMAccountExpiration = ObjUser.AccountExpirationDate
		If Err.Number = -2147467259 OR DTMAccountExpiration = #1/1/1970# OR DTMAccountExpiration = #1/1/1601# Then
			StrAccntReturn = "Account Never Expires"
		Else
			StrAccntReturn = Day(ObjUser.AccountExpirationDate) & "/" & Month(ObjUser.AccountExpirationDate) & "/" & Year(ObjUser.AccountExpirationDate)
			If DateDiff("d", #1/1/1601#, StrAccntReturn) <> 0 Then
				If DateDiff("d", Day(Date) & "/" & Month(Date) & "/" & Year(Date), StrAccntReturn) < 0 Then
					WScript.Echo ObjUser.SAMAccountName & " -- " & ObjUser.UserPrincipalName & " -- Already Expired On: " & StrAccntReturn
				End If
				If DateDiff("d", Day(Date) & "/" & Month(Date) & "/" & Year(Date), StrAccntReturn) = 0 Then
					WScript.Echo ObjUser.SAMAccountName & " -- " & ObjUser.UserPrincipalName & " -- Will Expire Today: " & StrAccntReturn
				End If
			End If
		End If
		Set ObjUser = Nothing
		ObjRS.MoveNext
	Wend
End If

ObjRS.Close:	Set ObjRS = Nothing
If ObjCommand.State <> 0 Then
	ObjCommand.State = 0
End If
Set ObjCommand = Nothing
ObjConn.Close:	Set ObjConn = Nothing
SetLocale(CurrLocale):	WScript.Quit