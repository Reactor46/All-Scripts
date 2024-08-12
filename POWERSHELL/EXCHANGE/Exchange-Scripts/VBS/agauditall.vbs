snServername = "servername.com"
mnMailboxname = "mailboxname"
username = "username"
domain = "domain"
strpassword = "password"

report = "<table border=""1"" width=""100%"">" & vbcrlf
report = report & "  <tr>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Folder Name</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080"" colspan=""2""><b><font color=""#FFFFFF"">Over 2 Years Old</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080"" colspan=""2""><b><font color=""#FFFFFF"">1-2 Years Old</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080"" colspan=""2""><b><font color=""#FFFFFF"">0-1 Years Old</font></b></td>" & vbcrlf
report = report & "</tr>" & vbcrlf
report = report & "  <tr>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">&nbsp;</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">#Messages</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Size(MB)</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">#Messages</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Size(MB)</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">#Messages</font></b></td>" & vbcrlf
report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Size(MB)</font></b></td>" & vbcrlf
report = report & "</tr>" & vbcrlf
strusername =  domain & "\" & username
szXml = "destination=https://" & snServername & "/exchange/&flags=0&username=" & strusername 
szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"
set req = createobject("microsoft.xmlhttp")
req.Open "post", "https://" & snServername & "/exchweb/bin/auth/owaauth.dll", False
req.send szXml
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for c = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(c)),"set-cookie: sessionid=") then reqsessionID = right(reqhedrarry(c),len(reqhedrarry(c))-12)
	if instr(lcase(reqhedrarry(c)),"set-cookie: cadata=") then reqcadata= right(reqhedrarry(c),len(reqhedrarry(c))-12)
Next
baseurl = "https://" & snServername & "/exchange/" & mnMailboxname
call RecurseFolder(baseurl)
report = report & "</table>" & vbcrlf
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\temp\mailboxagereport.htm",2,true) 
wfile.write report
wfile.close
set wfile = nothing
set fso = Nothing

Public Sub RecurseFolder(sUrl)
  
   req.open "SEARCH", sUrl, False, "", ""
   sQuery = "<?xml version=""1.0""?>"
   sQuery = sQuery & "<g:searchrequest xmlns:g=""DAV:"">"
   sQuery = sQuery & "<g:sql>SELECT ""http://schemas.microsoft.com/"
   sQuery = sQuery & "mapi/proptag/x0e080003"", ""DAV:hassubs"" FROM SCOPE "
   sQuery = sQuery & "('SHALLOW TRAVERSAL OF """ & sUrl & """') "
   sQuery = sQuery & "WHERE ""DAV:isfolder"" = true and ""DAV:ishidden"" = false and ""http://schemas.microsoft.com/mapi/proptag/x36010003"" = 1"
   sQuery = sQuery & "</g:sql>"
   sQuery = sQuery & "</g:searchrequest>"
   req.setRequestHeader "Content-Type", "text/xml"
   req.setRequestHeader "Translate", "f"
   req.setRequestHeader "Depth", "0"
   req.SetRequestHeader "cookie", reqsessionID
   req.SetRequestHeader "cookie", reqCadata
   req.setRequestHeader "Content-Length", "" & Len(sQuery)
   req.send sQuery
   Set oXMLDoc = req.responseXML
   Set oXMLSizeNodes = oXMLDoc.getElementsByTagName("d:x0e080003")
   Set oXMLHREFNodes = oXMLDoc.getElementsByTagName("a:href")
   Set oXMLHasSubsNodes = oXMLDoc.getElementsByTagName("a:hassubs")
   For i = 0 to oXMLSizeNodes.length - 1
      call procfolder(oXMLHREFNodes.Item(i).nodeTypedValue,sUrl)
      wscript.echo oXMLHREFNodes.Item(i).nodeTypedValue
      If oXMLHasSubsNodes.Item(i).nodeTypedValue = True Then
         call RecurseFolder(oXMLHREFNodes.Item(i).nodeTypedValue)
      End If
   Next
End Sub

sub procfolder(strURL,pfname)
wscript.echo strURL
ReDim resarray(1,6)
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:""  xmlns:b=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"">"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"",  ""urn:schemas:httpmail:subject"", "
strQuery = strQuery & """DAV:creationdate"", ""DAV:getcontentlength"", "
strQuery = strQuery & """urn:schemas:httpmail:fromemail"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = False</D:sql></D:searchrequest>"
req.open "SEARCH", strURL, false
req.setrequestheader "Content-Type", "text/xml"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Translate","f"
req.send strQuery
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:displayname")
   set oNodeList1 = oResponseDoc.getElementsByTagName("a:href")
   set oSize = oResponseDoc.getElementsByTagName("a:getcontentlength")
   set odatereceived = oResponseDoc.getElementsByTagName("a:creationdate")
   For i = 0 To (oNodeList.length -1)
		set oNode = oNodeList.nextNode
		set oNode1 = oNodeList1.nextNode
		set oNode2 = oSize.nextNode
		set oNode3 = odatereceived.nextNode
		wscript.echo oNode3.text
		If CDate(DateSerial(mid(oNode3.text,1,4), mid(oNode3.text,6,2),mid(oNode3.text,9,2))) < dateadd("m",-24,now()) Then
			resarray(0,1) = resarray(0,1) + 1
			resarray(1,1) = resarray(1,1) + Int(oNode2.text)
		End if
		If CDate(DateSerial(mid(oNode3.text,1,4), mid(oNode3.text,6,2),mid(oNode3.text,9,2))) > dateadd("m",-24,now()) And CDate(DateSerial(mid(oNode3.text,1,4), mid(oNode3.text,6,2),mid(oNode3.text,9,2))) < dateadd("m",-12,now()) Then
			resarray(0,2) = resarray(0,2) + 1
			resarray(1,2) = resarray(1,2) + Int(oNode2.text)
		End if
		If CDate(DateSerial(mid(oNode3.text,1,4), mid(oNode3.text,6,2),mid(oNode3.text,9,2))) > dateadd("m",-12,now()) And CDate(DateSerial(mid(oNode3.text,1,4), mid(oNode3.text,6,2),mid(oNode3.text,9,2))) < now() Then
			resarray(0,3) = resarray(0,3) + 1
			resarray(1,3) = resarray(1,3) + Int(oNode2.text)
		End if
	Next
Else
End If
report = report & "<tr>" & vbcrlf
report = report & "<td align=""center"">" & unescape(Replace(strURL,baseurl,"")) & "&nbsp;</td>" & vbcrlf
report = report & "<td align=""center"">" & resarray(0,1) & "&nbsp;</td>" & vbcrlf
report = report & "<td align=""center"">" & FormatNumber(resarray(1,1)/1024/1024,2) & "&nbsp;</td>" & vbcrlf
report = report & "<td align=""center"">" &  resarray(0,2) & "&nbsp;</td>" & vbcrlf
report = report & "<td align=""center"">" & FormatNumber(resarray(1,2)/1024/1024,2) & "&nbsp;</td>" & vbcrlf
report = report & "<td align=""center"">" &  resarray(0,3) & "&nbsp;</td>" & vbcrlf
report = report & "<td align=""center"">" & FormatNumber(resarray(1,3)/1024/1024,2) & "&nbsp;</td>" & vbcrlf
report = report & "</tr>" & vbcrlf
end sub
