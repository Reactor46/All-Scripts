Set reqfba = CreateObject("Microsoft.xmlhttp")
domain = "domain"
strpassword = "password"
username = "username"
servername = "servername.com"
sUrl = "https://" & servername & "/exchange/" & username
strusername =  domain & "\" & username
szXml = "destination=https://" & servername & "/exchange&flags=0&username=" & strusername 
szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"
reqfba.Open "post", "https://" & servername & "/exchweb/bin/auth/owaauth.dll", False
reqfba.send szXml
reqhedrarry = split(reqfba.GetAllResponseHeaders(), vbCrLf,-1,1)
for c = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(c)),"set-cookie: sessionid=") then reqsessionID = right(reqhedrarry(c),len(reqhedrarry(c))-12)
	if instr(lcase(reqhedrarry(c)),"set-cookie: cadata=") then reqcadata= right(reqhedrarry(c),len(reqhedrarry(c))-12)
next
reqfba.open "SEARCH", sUrl, False, "", ""
sQuery = "<?xml version=""1.0""?>"
sQuery = sQuery & "<g:searchrequest xmlns:g=""DAV:"">"
sQuery = sQuery & "<g:sql>SELECT ""http://schemas.microsoft.com/"
sQuery = sQuery & "mapi/proptag/x0e080003"", ""DAV:hassubs"" FROM SCOPE "
sQuery = sQuery & "('SHALLOW TRAVERSAL OF """ & sUrl & """') "
sQuery = sQuery & "WHERE ""DAV:isfolder"" = true and ""DAV:ishidden"" = false and ""http://schemas.microsoft.com/mapi/proptag/x36010003"" = 1"
sQuery = sQuery & "</g:sql>"
sQuery = sQuery & "</g:searchrequest>"
reqfba.setRequestHeader "Content-Type", "text/xml"
reqfba.setRequestHeader "Translate", "f"
reqfba.setRequestHeader "Depth", "0"
reqfba.SetRequestHeader "cookie", reqsessionID
reqfba.SetRequestHeader "cookie", reqCadata
reqfba.setRequestHeader "Content-Length", "" & Len(sQuery)
reqfba.send sQuery
Set oXMLDoc = reqfba.responseXML
Set oXMLSizeNodes = oXMLDoc.getElementsByTagName("d:x0e080003")
Set oXMLHREFNodes = oXMLDoc.getElementsByTagName("a:href")
Set oXMLHasSubsNodes = oXMLDoc.getElementsByTagName("a:hassubs")
For i = 0 to oXMLSizeNodes.length - 1
      wscript.echo oXMLHREFNodes.Item(i).nodeTypedValue
Next
