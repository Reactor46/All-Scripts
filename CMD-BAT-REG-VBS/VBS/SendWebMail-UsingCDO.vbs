'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
' Created by Assaf Miron
' Http://sites.google.com/site/assafmiron
' Date : 20/10/2009
' SendWebMail-UsingCDO.vbs
'=*=*=*=*=*=*=*=*=*=*=*=*==*=*==*=*==*=
Option Explicit

Sub SendMail(strUserName, strPassword, strSendTo, strSubject, strBody, strAttachment, bCheck)
On Error Resume Next
	Dim iMsg 
	Dim iConf 
	Dim Flds 
	Set iMsg = CreateObject("CDO.Message")
	Set iConf = CreateObject("CDO.Configuration")
	Set Flds = iConf.Fields
	
	schema = "http://schemas.microsoft.com/cdo/configuration/"
	If InStr(LCase(strUserName), "@gmail.com") Then
		' Using Gmail SMTP Server
		' smtp_host: smtp.gmail.com
		' smtp_port: 465 or 587
		' smtp_ssl: 1
		With Flds
			.Item(schema & "sendusing") = 2
			.Item(schema & "smtpserver") = "smtp.gmail.com" 
			.Item(schema & "smtpserverport") = 465
			.Item(schema & "smtpauthenticate") = 1
			.Item(schema & "sendusername") = strUserName
			.Item(schema & "sendpassword") = strPassword
			.Item(schema & "smtpusessl") = 1
			.Update
		End With
	ElseIf InStr(LCase(strUserName), "@hotmail.com") Then
		' Using Hotmail SMTP Server
		' smtp_host: smtp.live.com
		' smtp_port: 25 or 587
		' smtp_start_tls: 1
		With Flds
			.Item(schema & "sendusing") = 2
			.Item(schema & "smtpserver") = "smtp.live.com" 
			.Item(schema & "smtpserverport") = 25
			.Item(schema & "smtpauthenticate") = 1
			.Item(schema & "sendusername") = strUserName
			.Item(schema & "sendpassword") = strPassword
			.Item(schema & "smtpusessl") = 1
			.Update
		End With
	ElseIf InStr(LCase(strUserName), "@yahoo.com") Then
		' Using Yahoo Mail Plus (currently a paid service) SMTP Server
		' smtp_host: smtp.mail.yahoo.com
		' smtp_port: 25
		With Flds
			.Item(schema & "sendusing") = 2
			.Item(schema & "smtpserver") = "smtp.mail.yahoo.com" 
			.Item(schema & "smtpserverport") = 25
			.Item(schema & "smtpauthenticate") = 1
			.Item(schema & "sendusername") = strUserName
			.Item(schema & "sendpassword") = strPassword
			.Item(schema & "smtpusessl") = 0
			.Update
		End With
	Else
	' SMTP Server not Supported
	End If
	
	' Set Email Message Configuration
	With iMsg
		.From = strUserName
		If bCheck Then
			.To = strUserName
		Else
			.To = strSendTo
		End If
		.CC = ""
		.BCC = ""
		.Subject = strSubject
		.AddAttachment strAttachment
		.TextBody = strBody
		.Sender = strUserName
		.Organization = ""
		.ReplyTo = strUserName
		'.DSNOptions = cdoDSNSuccessFailOrDelay
		Set .Configuration = iConf
		.Send
	End With
	
	' Check for Errors
	If Err.Number = 0 Then
		WScript.Echo "EMail Sent Succefully"
	Else
		WScript.Echo "Error Sending Mail" & vbNewLine & Err.Description
	End If
	
	' Clean Up
	Set iMsg = nothing
	Set iConf = nothing
	Set Flds = nothing
End Sub

Dim strText, strSubject, strAttachmentPath
strText = "Hi!" & vbNewLine & "This is a Test Message!"
strSubject = "Send Web Mail - Test Mail
strAttachmentPath = ""
SendMail "myGmail@gmail.com", "MyPassW0rd", "toSomeOne@gmail.com", strSubject, strText, strAttachmentPath, True