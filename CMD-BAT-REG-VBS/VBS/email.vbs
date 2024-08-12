If WScript.Arguments.Count >= 3 Then
    strSMTPTo = WScript.Arguments.Item(0)
    strTextBody = WScript.Arguments.Item(1)
    strSubject = WScript.Arguments.Item(2)
    If WScript.Arguments.Count = 4 Then
        strAttachment = WScript.Arguments.Item(3)
    End If
Else
    Wscript.Echo "Usage: email.vbs SMTPTo TextBody Subject Attachment"
    Wscript.Quit
End If


strSMTPFrom = "itsupport@usonv.com"
strSMTPRelay = "192.168.1.5"


Set oMessage = CreateObject("CDO.Message")
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPRelay
oMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
oMessage.Configuration.Fields.Update

oMessage.Subject = strSubject
oMessage.From = strSMTPFrom
oMessage.To = strSMTPTo
oMessage.TextBody = strTextBody
oMessage.AddAttachment strAttachment


oMessage.Send