server = "servername"
mailbox = "mailbox"
strURL = "http://" & server & "/exchange/" & mailbox & "/inbox"
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"" "
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = True AND ""DAV:isfolder"" = False AND "
strQuery = strQuery & """http://schemas.microsoft.com/exchange/outlookmessageclass"" = 'IPM.ExtendedRule.Message' "
strQuery = strQuery & "AND ""http://schemas.microsoft.com/mapi/proptag/0x65EB001E"" = 'JunkEmailRule' </D:sql></D:searchrequest>"
set req = createobject("microsoft.xmlhttp")
req.open "SEARCH", strURL, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.send strQuery
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:href")
   For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	wscript.echo oNode.text
	xmlstr = xmlstr & "<D:href>" & oNode.text & "</D:href>"
   Next	
Else
End If
xmlstr1 = "<?xml version=""1.0"" ?>"
xmlstr1 = xmlstr1 & "<D:delete xmlns:D=""DAV:"">"
xmlstr1 = xmlstr1 & "<D:target>"
xmlstr1 = xmlstr1 & xmlstr
xmlstr1 = xmlstr1 & "</D:target>"
xmlstr1 = xmlstr1 & "</D:delete>"

wscript.echo

if oNodeList.length <> 0 then
	req.open "BDELETE", strURL, false
	req.setrequestheader "Content-Type", "text/xml"
	req.setRequestHeader "Translate","f"
	req.send xmlstr1
	If req.status >= 500 Then
	ElseIf req.status = 207 Then
		Wscript.echo "Success"
	Else
	wscript.echo "Request Failed. Results = " & req.Status & ": " & req.statusText
	End If
Else
	wscript.echo "No Rule Found"
end if

