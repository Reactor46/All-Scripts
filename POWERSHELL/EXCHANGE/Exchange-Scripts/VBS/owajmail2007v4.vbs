snServername = "servername"
mnMailboxname = "user"
domain = "domain"
strpassword = "password"

Targetmailbox = "mailboxs@domain.com"

strusername =  domain & "\" & mnMailboxname
szXml = "destination=https://" & snServername & "/owa/&flags=0&username=" & strusername 
szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"


set req = createobject("MSXML2.ServerXMLHTTP.6.0")
req.Open "post", "https://" & snServername & "/owa/auth/owaauth.dll", False
req.SetOption 2, 13056 
req.send szXml
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for each ent in reqhedrarry
	wscript.echo ent
	if instr(ent,"UserContext=") then
		slen = instr(ent,"UserContext=")+12
		elen = instr(slen,ent,";")
		canary = mid(ent,slen,elen-slen)
		wscript.echo "parsed canary : " & canary 
	end if
Next

req.Open "GET", "https://" & snServername &  "/owa/" & Targetmailbox & "", False
req.SetOption 2, 13056 
req.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
req.setRequestHeader "Content-Length", Len(xmlstr)
req.send szXml
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for each ent in reqhedrarry
	wscript.echo ent
Next
cookie = req.getResponseHeader("Set-Cookie")
if instr(cookie,"=") then
	slen = instr(cookie,"=")+1
	elen = instr(slen,cookie,"&")
	canary = mid(cookie,slen,elen-slen)
	wscript.echo "parsed canery: " & canary 
end if

wscript.echo "Cookie : " & cookie

Call UpdateJunk(Targetmailbox,canary)

Sub UpdateJunk(mbMailbox,canary )

xmlstr = "<params><canary>" & canary & "</canary><fEnbl>1</fEnbl></params>" 
wscript.echo xmlstr
req.Open "POST", "https://" & snServername & "/owa/" & mbMailbox & "/ev.owa?oeh=1&ns=JunkEmail&ev=Enable", False
req.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
req.setRequestHeader "Content-Length", Len(xmlstr)
req.setRequestHeader "User-Agent","Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/4.0; GTB6.5; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; HPNTDF; .NET4.0C; InfoPath.3; .NET4.0E; Zune 4.0)"
req.send xmlstr
wscript.echo req.status
wscript.echo req.responsetext
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for each ent in reqhedrarry
	wscript.echo ent
Next
If InStr(req.responsetext,"name=lngFrm") Then
	wscript.echo "Mailbox has not been logged onto before via OWA"
	'Create a regular expression object
	Dim objRegExp
	Set objRegExp = New RegExp

	objRegExp.Pattern = "<option selected value=""(.*?)"">"
	objRegExp.IgnoreCase = True
	objRegExp.Global = True

	Dim objMatches
	Set objMatches = objRegExp.Execute(req.responsetext)
	If objMatches.count = 2 then
		lcidarry = Split(objMatches(0).Value,Chr(34))
		wscript.echo lcidarry(1)
		tzidarry = Split(objMatches(1).Value,Chr(34))
		wscript.echo tzidarry(1)
		pstring = "lcid=" & lcidarry(1) & "&tzid=" & tzidarry(1)
		req.Open "POST", "https://" & snServername & "/owa/" & mbMailbox & "/lang.owa", False
		req.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		req.setRequestHeader "Content-Length", Len(pstring)
'		req.SetRequestHeader "cookie", reqCadata
		req.send pstring
		if instr(req.responsetext,"errMsg") then 
			wscript.echo "Permission Error"
		else
			wscript.echo req.status
			If req.status = 200 and not instr(req.responsetext,"errMsg") Then 
				Call UpdateJunk(mbMailbox)
			Else
				wscript.echo "Failed to set Default OWA settings"
			End if
		end if

	Else
		wscript.echo "Script failed to retrieve default values"
	End if
Else
	wscript.echo "Junk Mail Setting Updated"
End if
End sub


