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
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------


Function MailInactiveUser(EmailFrom,EmailTo,SMTPServer,SMTPLogon,SMTPPassword,InactiveDays)
'This funtion is to send mail to adminstrators about inactive users.
	EmailSubject = "Inactive Users in " & InactiveDays & "days"
	Const SMTPSSL = True
	Const SMTPPort = 25
	Const cdoSendUsingPickup = 1 	'Send message using local SMTP service pickup directory.
	Const cdoSendUsingPort = 2 	'Send the message using SMTP over TCP/IP networking.
	Const cdoAnonymous = 0 	' No authentication
	Const cdoBasic = 1 	' BASIC clear text authentication
	Const cdoNTLM = 2 	' NTLM, Microsoft proprietary authentication
	' First, create the message
	Set objMessage = CreateObject("CDO.Message")
	objMessage.Subject = EmailSubject
	objMessage.From = " <" & EmailFrom & ">"
	objMessage.To = EmailTo
	objMessage.HTMLBody = GetInactiveUsers(InactiveDays)
	' Second, configure the server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPLogon
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassword
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SMTPSSL
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	objMessage.Configuration.Fields.Update
	'Now send the message!
	objMessage.Send
End Function 


Function GetInactiveUsers(InactiveDays)
'This function is to get inactive users in a period time
	Dim objRootDSE, adoConnection, adoCommand, strQuery
	Dim adoRecordset, strDNSDomain, objShell, lngBiasKey
	Dim lngBias, k, strDN, dtmDate, objDate
	Dim strBase, strFilter, strAttributes, lngHigh, lngLow
	Dim Body
	'Dim the html body
	Body = "Hi Administrators,<br />The followings are inactive users in " & InactiveDays & " days:<br /><br />"_
	& "<TABLE style=""BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid"""_
	& "cellSpacing=1 cellPadding=1 width=""60%"" border=1><tr><td>AccountName</td><td>WhenCreated</td><td>LastLogonDate</td></tr>"

	' Obtain local Time Zone bias from machine registry.
	' This bias changes with Daylight Savings Time.
	Set objShell = CreateObject("Wscript.Shell")
	lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
	    & "TimeZoneInformation\ActiveTimeBias")
	If (UCase(TypeName(lngBiasKey)) = "LONG") Then
	    lngBias = lngBiasKey
	ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
	    lngBias = 0
	    For k = 0 To UBound(lngBiasKey)
	        lngBias = lngBias + (lngBiasKey(k) * 256^k)
	    Next
	End If
	Set objShell = Nothing
	' Determine DNS domain from RootDSE object.
	Set objRootDSE = GetObject("LDAP://RootDSE")
	strDNSDomain = objRootDSE.Get("defaultNamingContext")
	Set objRootDSE = Nothing
	' Use ADO to search Active Directory.
	Set adoCommand = CreateObject("ADODB.Command")
	Set adoConnection = CreateObject("ADODB.Connection")
	adoConnection.Provider = "ADsDSOObject"
	adoConnection.Open "Active Directory Provider"
	adoCommand.ActiveConnection = adoConnection	
	strQuery =  "<LDAP://" & strDNSDomain & ">;(&(objectCategory=person)(objectClass=user));Name,whenCreated,lastLogonTimeStamp;subtree"
	' Run the query.	
	adoCommand.CommandText = strQuery
	adoCommand.Properties("Page Size") = 0
	adoCommand.Properties("Timeout") = 60
	adoCommand.Properties("Cache Results") = False
	Set adoRecordset = adoCommand.Execute
	
	' Enumerate resulting recordset.
	Do Until adoRecordset.EOF
	   ' Retrieve attribute values for the user.
	    strDN = adoRecordset.Fields("Name").Value
	    WhenCreated = adoRecordset.Fields("whenCreated").Value
	    ' Convert Integer8 value to date/time in current time zone.
	    On Error Resume Next
	    Set objDate = adoRecordset.Fields("lastLogonTimeStamp").Value
	    If (Err.Number <> 0) Then
	        On Error GoTo 0
	        dtmDate = #1/1/1601#
	    Else
	        On Error GoTo 0
	        lngHigh = objDate.HighPart
	        lngLow = objDate.LowPart
	        If (lngLow < 0) Then
	            lngHigh = lngHigh + 1
	        End If
	        If (lngHigh = 0) And (lngLow = 0) Then
	            dtmDate = #1/1/1601#
	        Else
	            dtmDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
	                + lngLow)/600000000 - lngBias)/1440
	        End If
	    End If
	    ' Display values for the user.
	    If (dtmDate = #1/1/1601#) Then
	    	If DateDiff("d",whenCreated, DateNow) > InactiveDays Then 
	    		body = body & "<tr><td>" & strDN & "</td><td>" & whenCreated & "</td><td>" & "Never" & "</td></tr>"
	    	End If
	    Else
	        Dim DateNow
			DateNow = Now()
			If DateDiff("d",dtmDate, DateNow) > InactiveDays Then 
				body = body & "<tr><td>" & strDN & "</td><td>" & whenCreated & "</td><td>" & dtmDate & "</td></tr>"
			End If 
	    End If
	    adoRecordset.MoveNext
	Loop
		Body = Body & "</table>"	
	'Clean up.
	adoRecordset.Close
	GetInactiveUsers = Body 
End Function 


