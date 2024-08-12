snServername = "servername"
mnMailboxname = "mailbox"
SourceURL = "http://" & snServername & "/exchange/" & mnMailboxname & "/inbox/"

strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"", ""http://schemas.microsoft.com/mapi/proptag/x65E90003"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & SourceURL & """') Where ""DAV:ishidden"" = True AND ""DAV:isfolder"" = False AND "
strQuery = strQuery & """http://schemas.microsoft.com/exchange/outlookmessageclass"" = 'IPM.ExtendedRule.Message' "
strQuery = strQuery & "AND ""http://schemas.microsoft.com/mapi/proptag/0x65EB001E"" = 'JunkEmailRule' </D:sql></D:searchrequest>"
set req = createobject("microsoft.xmlhttp")
req.open "SEARCH", SourceURL, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.send strQuery
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:href")
   set ostateList = oResponseDoc.getElementsByTagName("d:x65E90003")
   For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	wscript.echo "Rule Found"
	Rule_State = ostateList(i).nodetypedvalue
	select case Rule_State
		case 48 wscript.echo "OWA Junk Email Setting Off"
		case 49 wscript.echo "OWA Junk Email Setting Enabled"
   	end Select
   Next	
   if oNodeList.length = 0 then
	wscript.echo "No Junk Email rules Exists"
   End if
Else
End If
wscript.echo
