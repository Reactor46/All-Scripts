' ================================================================================================================
' Password Expiration Information: Know When User's Password Will Expire
' Usage: CScript /NoLogo ChkPasswordExpiry.vbs <UserName> OR CScript /NoLogo ChkPasswordExpiry.vbs "UserName"
' Example: CScript /NoLogo ChkPasswordExpiry.vbs msanyal OR CScript /NoLogo ChkPasswordExpiry.vbs "msanyal"
' ================================================================================================================

Option Explicit

Dim ObjDomain, ObjUser, PwdAgeDays, PwdDaysValue
Dim StrDomName, StrUserLDAP, ObjRootDSE, PassExpDate
Dim ObjConn, ObjRS, StrUserName, StrSQL, FromDate, CurrDate

If WScript.Arguments.Count <> 1 Then
	WScript.Echo vbNullString
	WScript.Echo "Error: NO USER-NAME ENTERED" 
    WScript.Echo "Usage:  ScriptName <UserName>" 
    WScript.Echo "Example: " & Trim(WScript.ScriptName) & " " & "BobS" 
    WScript.Quit
End If 
StrUserName = Trim(WScript.Arguments(0))

On Error Resume Next
Set ObjRootDSE = GetObject("LDAP://RootDSE")
If Err.Number <> 0 Then
	Err.Clear
	WScript.Echo "Error: This Computer is not connected to Domain."
	WScript.Quit
End If
StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext"))
Set ObjRootDSE = Nothing

StrSQL = "Select ADsPath From 'LDAP://" & StrDomName & "' Where ObjectCategory = 'User' AND SAMAccountName = '" & StrUserName & "'" 

Set ObjConn = CreateObject("ADODB.Connection") 
ObjConn.Provider = "ADsDSOObject":    ObjConn.Open "Active Directory Provider" 
Set ObjRS = CreateObject("ADODB.Recordset") 
ObjRS.Open StrSQL, ObjConn 
If ObjRS.EOF Then
	WScript.Echo vbNullString
	WScript.Echo "User: " & StrUserName & " does not exist in Active Directory !!"
	ObjRS.Close:	Set ObjRS = Nothing
	ObjConn.Close:	Set ObjConn = Nothing
	WScript.Quit
Else
	StrUserLDAP = Trim(ObjRS.Fields("ADsPath").Value)
End If
ObjRS.Close:	Set ObjRS = Nothing
ObjConn.Close:	Set ObjConn = Nothing

' --- Get Maximum Password Age as Set In Domain Policy
Set ObjDomain = GetObject("LDAP://" & StrDomName)
Set PwdAgeDays = ObjDomain.Get("MaxPwdAge")
' --- Know The Number of Days Of Maximum Password Age
PwdDaysValue = CCur((PwdAgeDays.HighPart * 2 ^ 32) + PwdAgeDays.LowPart) / CCur(-864000000000)
WScript.Echo VbCrLf & "Maximum Password Age as configured in Domain: " & PwdDaysValue & " Days"
WScript.Echo "---------------------------------------------"
WScript.Echo vbNullString
' --- Know the Last Time When User Changed Password
Set ObjUser = GetObject(StrUserLDAP)
If ObjUser.UserAccountControl And 65536 Then
	Set ObjUser = Nothing:	Set PwdAgeDays = Nothing:	Set ObjDomain = Nothing
	WScript.Echo "Username: " & Ucase(StrUserName)
	WScript.Echo "This User's Password Has Been Set To Never Expire !!"
	WScript.Quit
End If
' --- Add Total Days To The Last Time User Set The Password
PassExpDate = DateAdd("D", PwdDaysValue, ObjUser.PasswordLastChanged)
WScript.Echo "The Last Time This User Changed Password Was On: "
WScript.Echo "---------------------------------------------------"
WScript.Echo vbNullString
WScript.Echo "===>> " & FormatDateTime (Trim(ObjUser.PasswordLastChanged), 1) & " at " & FormatDateTime (Trim(ObjUser.PasswordLastChanged), 3)
WScript.Echo vbNullString
WScript.Echo "This User's Password Will Expire On: "
WScript.Echo "----------------------------------------"
WScript.Echo vbNullString
WScript.Echo "===>> " & FormatDateTime (Trim(PassExpDate), 1) & " at " & FormatDateTime (Trim(PassExpDate), 3)

FromDate = FormatDateTime(Trim(PassExpDate), 2)
CurrDate = FormatDateTime (Trim(Date), 2)

WScript.Echo vbNullString
WScript.Echo "Days Left Before Current Password Expires: " & DateDiff("d", CurrDate, FromDate) & " Days Remaining."

Set ObjUser = Nothing:	Set PwdAgeDays = Nothing:	Set ObjDomain = Nothing