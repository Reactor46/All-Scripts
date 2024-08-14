' ========================================================================================================================================================
' Contact --- monimoys@hotmail.com		
' Script to start notifying the currently Logged-On User via Email when his/her Windows Active Directory Password is going to expire.
' User will start to receive email from 10 Days before the date when his/her password expires.
' This script WILL NOT work if the computer, from where User has logged-in, in not joined to any domain OR if the computer is disconnected
' from the domain. This Script should work fine in Windows Active Directory environment running Windows 2008 R2 Domain Controllers and having
' Email or Messaging systems running on Microsoft Exchange 2010.
' ========================================================================================================================================================

Option Explicit

Const ADS_SCOPE_SUBTREE = 2
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000 
Dim ObjRootDSE, ObjNetwork, ObjConn, ObjRS
Dim StrDomName, StrUser, StrComputer, IntPassAge, StrSQL
Dim BlnDomJoined, BlnDomConn, StartFromDays

BlnDomJoined = False:	BlnDomConn = False
Set ObjNetwork = CreateObject("WScript.Network")
StrComputer = Trim(ObjNetwork.ComputerName):	StrUser = Trim(ObjNetwork.UserName)
Set ObjNetwork = Nothing
CheckDomainJoined
If BlnDomJoined = True Then
	CheckDomainConnection
End If
If BlnDomConn = True AND BlnDomJoined = True Then
	GetDomainPasswordAge:	StartFromDays = 10
	StrSQL = "Select ADsPath From 'LDAP://" & StrDomName & "' Where ObjectClass = 'User' AND SAMAccountName = '" & StrUser & "'"
	WorkToNotifyUser
End If
StrSQL = vbNullString:	WScript.Quit

' =================================================================================================================================================

Private Sub GetDomainPasswordAge
	Dim ObjDomain, ObjMaxAge
	Set ObjDomain = GetObject("LDAP://" & StrDomName)
	Set ObjMaxAge = ObjDomain.Get("maxPwdAge")
	IntPassAge = CCur((ObjMaxAge.HighPart * 2 ^ 32) + ObjMaxAge.LowPart) / CCur(-864000000000)
	Set ObjDomain = Nothing
End Sub

' =================================================================================================================================================

Private Sub WorkToNotifyUser
	
	Dim ObjUser, StrSAM, StrFullName, StrMail, dtmPassChng
	Dim WhenPasswordExpires, DaysRemain
	
	Set ObjConn = CreateObject("ADODB.Connection")
	ObjConn.Provider = "ADsDSOObject":	ObjConn.Open "Active Directory Provider"
	Set ObjRS = CreateObject("ADODB.Recordset")
	ObjRS.Open StrSQL, ObjConn
	Set ObjUser = GetObject(Trim(ObjRS.Fields("ADsPath").Value))
	If ObjUser.AccountDisabled = FALSE AND ObjUser.UserAccountControl AND ADS_UF_DONT_EXPIRE_PASSWD = 0 Then
		StrSAM = Trim(ObjUser.SAMAccountName):	StrMail = Trim(ObjUser.Mail)
		StrFullName = Trim(ObjUser.DisplayName):	dtmPassChng = FormatDateTime(ObjUser.PasswordLastChanged, 2)	
		WhenPasswordExpires = DateAdd("d", IntPassAge, dtmPassChng)
		DaysRemain = FormatDateTime(Date, 2)
		DaysRemain = DateDiff("d", DaysRemain, WhenPasswordExpires) 
		If (DaysRemain > 0) AND (DaysRemain <= StartFromDays) Then
			If StrMail <> vbNullString Then
				Call NowWeSendTheMail(StrDomName, StrMail, StrFullName, ObjUser.PasswordLastChanged, DaysRemain)
			' --- Else
				' --- WScript.Echo "Not Applicable"
			End If
		End If		
	End If
	Set ObjUser = Nothing
	ObjRS.Close:	Set ObjRS = Nothing
	ObjConn.Close:	Set ObjConn = Nothing
	
End Sub

' =================================================================================================================================================

Private Sub NowWeSendTheMail(StrThisDomain, ThisMailAddr, ThisName, PassDate, DaysLeft)

	Dim ObjEMail
	
	StrThisDomain = Replace(StrThisDomain, "DC=", vbNullString)
	StrThisDomain = Replace(StrThisDomain, ",", ".")

	Set ObjEmail = CreateObject("CDO.Message")
	ObjEmail.From = "Test@" & StrThisDomain:	ObjEmail.To = ThisMailAddr
	ObjEmail.Subject = ThisName & " -- Your Password Will Expire On " & PassDate
	ObjEmail.HTMLBody = VbCrLf & "********** THIS MAIL IS AUTO-GENERATED. DO NOT REPLY TO THIS MAIL **********" & VbCrLf & VbCrLf
	ObjEmail.HTMLBody = ObjEmail.HTMLBody & "<H3>Dear " & ThisName & ", </H3>" & VbCrLf & VbCrLf
	ObjEmail.HTMLBody = ObjEmail.HTMLBody & "<H4>Your Password Expires On: " & PassDate & ". </H4>" & VbCrLf & VbCrLf
	ObjEmail.HTMLBody = ObjEmail.HTMLBody & "<H4>Days Remaining To Change Your Password: " & DaysLeft & " day(s).</H4>" & VbCrLf & VbCrLf
	ObjEmail.HTMLBody = ObjEmail.HTMLBody & "<H4>Please Press CTRL-ALT-DEL and select the option CHANGE A PASSWORD</H4>" & VbCrLf & VbCrLf
	ObjEmail.HTMLBody = ObjEmail.HTMLBody & "<H4>You will continue to receive this reminder mail till you change your password.</H4>" & VbCrLf & VbCrLf
	ObjEmail.HTMLBody = ObjEmail.HTMLBody & "<H4>Yours Sincerely </H4>" & VbCrLf
	ObjEmail.HTMLBody = ObjEmail.HTMLBody & "Administrator@" & StrThisDomain & VbCrLf
	ObjEmail.HTMLBody = ObjEmail.HTMLBody & VbCrLf & "********** THIS MAIL IS AUTO-GENERATED. DO NOT REPLY TO THIS MAIL **********" & VbCrLf & VbCrLf

	ObjEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

	' ---- Provide the IP Address Or FQDN Of the SMTP Server Below
	ObjEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "xxx.xxx.xxx.xxx"

	ObjEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

	ObjEmail.Configuration.Fields.Update:	ObjEmail.Send
	Set ObjEmail = Nothing

End Sub

' =================================================================================================================================================

Private Sub CheckDomainJoined()
    Dim ObjWMI, ColItems, ObjItem
    Set ObjWMI = GetObject("WinMgmts:\\" & StrComputer & "\Root\CIMV2")
    Set ColItems = ObjWMI.ExecQuery("Select * From Win32_ComputerSystem", , 48)
	For Each ObjItem In ColItems
		Select Case Trim(ObjItem.DomainRole)
			Case 0
				BlnDomJoined = False
			Case 2
				BlnDomJoined = False
			Case Else
				BlnDomJoined = True
		End Select
	Next
    Set ColItems = Nothing:	Set ObjWMI = Nothing:	ObjItem = vbNullString
End Sub

Private Sub CheckDomainConnection()
    On Error Resume Next
    Set ObjRootDSE = GetObject("LDAP://RootDSE")
	If Err.Number <> 0 Then
		BlnDomConn = False:	Err.Clear
		Exit Sub
	Else
		StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext"))
		If StrDomName <> vbNullString Then
			BlnDomConn = True
		Else
			BlnDomConn = False
		End If
	End If    
    Set ObjRootDSE = Nothing
End Sub

' =================================================================================================================================================