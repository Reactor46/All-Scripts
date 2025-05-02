' ----Config----
Toaddress = wscript.arguments(0)
Message = wscript.arguments(1)
Servername = "servername.domain.com"
LogonSipAddress = "user@domain.com"
Username = "domain\username"
password = "password"
' ----End Config----
set req = createobject("microsoft.xmlhttp")
' ----Logon
Logonstr = "https://" & Servername & "/iwa/logon.html?uri=" & LogonSipAddress & "&signinas=1&language=en&epid="
req.Open "GET",Logonstr , False, Username ,password
req.send
wscript.echo req.status
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for c = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(c)),"set-cookie:") then reqsessionID = right(reqhedrarry(c),len(reqhedrarry(c))-12)
next
chk =  left(reqsessionID,36)
' ---Send Message---
Sendmsgcmd =  "https://" & Servername & "/cwa/MainCommandHandler.ashx?Ck=" & chk
Messagestr = "cmdPkg=1,LcwStartImRequest,sip:" & ToAddress & "," & Message & ",X-MMS-IM-Format: FN=Arial%253B EF=%253B CO=000000%253B CS=1%253B PF=00"
req.open "POST", Sendmsgcmd, False, Username ,password
req.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
req.setRequestHeader "Cookie:", reqsessionID 
req.send Messagestr
wscript.echo req.status
' ---Logoff---
logoffstr = "https://" & Servername &"/cwa/SignoutHandler.ashx?Ck=" & chk
req.Open "GET",logoffstr , False, Username ,password
req.setRequestHeader "Cookie:", reqsessionID 
req.send
wscript.echo req.status
