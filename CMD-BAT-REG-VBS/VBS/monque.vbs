cComputerName = "."
MessageThreshold = 5
LastAlertSent = dateadd("h",-1,now())
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & _
cComputerName & "\root\MicrosoftExchangeV2")
Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
("SELECT * FROM __InstanceOperationEvent WITHIN 10 WHERE " _
& "Targetinstance ISA 'Exchange_SMTPQueue' and TargetInstance.MessageCount >= " & MessageThreshold)
Do
Set objLatestEvent = colMonitoredEvents.NextEvent
	Wscript.echo now() & "	" & objLatestEvent.TargetInstance.LinkName & "	" & objLatestEvent.TargetInstance.MessageCount & "	" & objLatestEvent.TargetInstance.Size 
	if LastAlertSent < dateadd("h",-1,now()) then
		call EnumSMTPQueues()
		LastAlertSent = now()
	end if
Loop


sub EnumSMTPQueues()
Const cWMINameSpace = "root/MicrosoftExchangeV2"
Const cWMIInstance = "Exchange_SMTPQueue"
HtmlMsgbody = 	"<table border=""1"" width=""100%"" cellpadding=""0"" bordercolor=""#000000""><tr><td bordercolor=""#FFFFFF"" align=""center"" bgcolor=""#000080"">" _
& "<b><font color=""#FFFFFF"">Queue Name</font></b></td><td bordercolor=""#FFFFFF"" align=""center"" bgcolor=""#000080""<b><font color=""#FFFFFF"">Message Count</font></b></td>" _
& "<td bordercolor=""#FFFFFF"" align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Queue Size</font></b></td></tr>"
strWinMgmts = "winmgmts:{impersonationLevel=impersonate}!//"& _
cComputerName&"/"&cWMINameSpace
Set objWMIExchange =  GetObject(strWinMgmts)
If Err.Number <> 0 Then
  WScript.Echo "ERROR: Unable to connect to the WMI namespace."
Else
  Set listExchange_PublicFolders = objWMIExchange.InstancesOf(cWMIInstance)
  For Each objExchange_SMTPQueue in listExchange_PublicFolders 
       HtmlMsgbody = HtmlMsgbody & "<tr><td>" &   objExchange_SMTPQueue.LinkName & "</td><td>" & objExchange_SMTPQueue.MessageCount _
	& "</td><td>" & objExchange_SMTPQueue.size & "</td></tr>" 
       WScript.echo objExchange_SMTPQueue.LinkName & "	" & objExchange_SMTPQueue.MessageCount  & "	" &   objExchange_SMTPQueue.size
       if objExchange_SMTPQueue.MessageCount >=  MessageThreshold then
	  wql ="Select * From Exchange_QueuedSMTPMessage Where LinkId='" & objExchange_SMTPQueue.LinkID 
      	  wql = wql & "' And LinkName='" & objExchange_SMTPQueue.Linkname & "' And ProtocolName='SMTP' And "
      	  wql = wql &  "QueueId='" & objExchange_SMTPQueue.QueueID & "' And QueueName='" & objExchange_SMTPQueue.Queuename  &"' And"
      	  wql = wql &  "  VirtualMachine='" & objExchange_SMTPQueue.VirtualMachine & "'"
       	  wql = wql &  "  And VirtualServerName='" & objExchange_SMTPQueue.VirtualServerName  & "'"  
	  quehtml = quehtml & getmess(wql) 
       end if 
  next
End If
HtmlMsgbody = HtmlMsgbody & "</table><BR><B>Message Queues</B><BR>" & quehtml
Set objEmail = CreateObject("CDO.Message")
objEmail.From = "Queuewarnings@yourdomain.com"
objEmail.To = "somebody@yourdomain.com"
objEmail.Subject = "Queue Threshold Exceeded"
objEmail.HTMLbody = HtmlMsgbody
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "Servername"
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Send
wscript.echo "message sent"
End sub

function getmess(wql)
quehtml = "<table border=""1"" width=""100%""><tr><td bgcolor=""#008000"" align=""center""><b><font color=""#FFFFFF"">Date Sent</font></b></td>" _
           & "<td bgcolor=""#008000"" align=""center""><b><font color=""#FFFFFF"">Sent By</font></b></td>"_
	   & "	<td bgcolor=""#008000"" align=""center""><b><font color=""#FFFFFF"">Recipients</font></b></td>"_
	   & "	<td bgcolor=""#008000"" align=""center""><b><font color=""#FFFFFF"">Subject</font></b></td>"_
	   & "	<td bgcolor=""#008000"" align=""center""><b><font color=""#FFFFFF"">Size</font></b></td></tr>"
Const cWMINameSpace = "root/MicrosoftExchangeV2"
strWinMgmts = "winmgmts:{impersonationLevel=impersonate}!//" & cComputerName & "/" & cWMINameSpace
Set objWMIExchange =  GetObject(strWinMgmts)
Set listExchange_MessageQueueEntries = objWMIExchange.ExecQuery(wql)
For each objExchange_MessageQueueEntries in listExchange_MessageQueueEntries
 recieved = dateadd("h",toffset,cdate(DateSerial(Left(objExchange_MessageQueueEntries.Received, 4), Mid(objExchange_MessageQueueEntries.Received, 5, 2), Mid(objExchange_MessageQueueEntries.Received, 7, 2)) & " " & timeserial(Mid(objExchange_MessageQueueEntries.Received, 9, 2),Mid(objExchange_MessageQueueEntries.Received, 11, 2),Mid(objExchange_MessageQueueEntries.Received,13, 2))))
 Wscript.echo recieved & " " & objExchange_MessageQueueEntries.Sender & " " & objExchange_MessageQueueEntries.Subject _
 & " " &  objExchange_MessageQueueEntries.size & " " & replace(replace(objExchange_MessageQueueEntries.Recipients(0),vbcrlf,""),"Envelope Recipients:","") 
 quehtml = quehtml & "<tr><td>" & recieved &"</td><td>" & objExchange_MessageQueueEntries.Sender & "</td><td>" & replace(replace(objExchange_MessageQueueEntries.Recipients(0),vbcrlf,""),"Envelope Recipients:","") & "</td><td>" _
 & objExchange_MessageQueueEntries.Subject & "</td><td>" & objExchange_MessageQueueEntries.size & "</td></tr>"
next
quehtml = quehtml & "</table>"
getmess = quehtml
end function