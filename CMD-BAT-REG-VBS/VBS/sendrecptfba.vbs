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

stmsgbody ="Hello Fred <BR<BR>"
stmsgbody = stmsgbody & "What to get some coffee Mate, Meet me downstairs in ten<BR><BR>"
stmsgbody = stmsgbody & "Cheers<BR>"
stmsgbody = stmsgbody & "Barney<BR>"
szXml = ""
szXml = szXml & "Cmd=send" & vbLf
szXml = szXml & "MsgTo=address@youdomain.com" & vbLf
szXml = szXml & "MsgCc=" & vbLf
szXml = szXml & "MsgBcc=" & vbLf
szXml = szXml & "urn:schemas:httpmail:importance=1" & vbLf
szXml = szXml & "http://schemas.microsoft.com/exchange/readreceiptrequested=1" & vbLf
szXml = szXml & "http://schemas.microsoft.com/exchange/deliveryreportrequested=1" & vbLf
szXml = szXml & "http://schemas.microsoft.com/exchange/sensitivity-long=" & vbLf
szXml = szXml & "urn:schemas:httpmail:subject=Coffee ???" & vbLf
szXml = szXml & "urn:schemas:httpmail:htmldescription=<!DOCTYPE HTML PUBLIC " _
& """-//W3C//DTD HTML 4.0 Transitional//EN""><HTML DIR=ltr><HEAD><META HTTP-EQUIV" _
& "=""Content-Type"" CONTENT=""text/html; charset=utf-8""></HEAD><BODY><DIV>" _
& "<FONT face='Arial' color=#000000 size=2>" & stmsgbody & "</font>" _
& "</DIV></BODY></HTML>" & vbLf
req.Open "POST", "https://" & servername & "/exchange/" & mailbox & "/", False, "", ""
req.setRequestHeader "Accept-Language:", "en-us"
req.setRequestHeader "Content-type:", "application/x-www-UTF8-encoded"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Content-Length:", Len(szXml)
req.Send szXml
Wscript.echo req.responseText 

