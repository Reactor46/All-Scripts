'********************** email send script ****************************

Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = "Windows Updates Availability: <SERVERNAME>" 
objMessage.From = "<SERVERNAME>@.net" 
objMessage.To = "myemail@a.com"
'objMessage.Cc = "CC To Others"
objMessage.TextBody = "Windows Updates Availability"
objMessage.AddAttachment "add attachment path here"
objMessage.Send
wscript.quit