Set req = CreateObject("Microsoft.XMLhttp")
servername = "servername"
mailbox = "mailbox"
domain = "domain"
strpassword = "password"
strusername =  domain & "\" & mailbox
szXml = "destination=https://" & servername & "/exchange&flags=0&username=" & strusername 
szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"
req.Open "post", "https://" & servername & "/exchweb/bin/auth/owaauth.dll", False
req.send szXml
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for i = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(i)),"set-cookie: sessionid=") then reqsessionID = right(reqhedrarry(i),len(reqhedrarry(i))-12)
	if instr(lcase(reqhedrarry(i)),"set-cookie: cadata=") then reqcadata= right(reqhedrarry(i),len(reqhedrarry(i))-12)
next

xmlstr = ""
xmlstr = xmlstr & "Cmd=options" & vbLf
xmlstr = xmlstr & "junkemailstate=1" & vbLf
xmlstr = xmlstr & "cmd=savejunkemailrule" & vbLf
xmlstr = xmlstr & "addtots=user@domain;@domain1.com;@domain2.com;"
req.Open "POST", "https://" & servername & "/exchange/" & mailbox & "/", False, "", ""
req.setRequestHeader "Accept-Language:", "en-us"
req.setRequestHeader "Content-type:", "application/x-www-UTF8-encoded"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Content-Length:", Len(xmlstr)
req.Send xmlstr
Wscript.echo req.responseText 

