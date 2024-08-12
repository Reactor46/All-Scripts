snServername = "servername"
mnMailboxname = "UserName"
domain = "domain"
strpassword = "password"

strusername =  domain & "\" & mnMailboxname
szXml = "destination=https://" & snServername & "/owa/&flags=0&username=" & strusername 
szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"


set req = createobject("microsoft.xmlhttp")
req.Open "post", "https://" & snServername & "/owa/auth/owaauth.dll", False
req.send szXml
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for c = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(c)),"set-cookie: usercontext=") then reqcadata= right(reqhedrarry(c),len(reqhedrarry(c))-12)
next
for each ent in reqhedrarry
	wscript.echo ent
Next

Call UpdateJunk("user@smtpdomain.com")

Sub UpdateJunk(mbMailbox)

xmlstr = "<params><fEnbl>1</fEnbl></params>"
req.Open "POST", "https://" & snServername & "/owa/" & mbMailbox & "/ev.owa?oeh=1&ns=JunkEmail&ev=Enable", False
req.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
req.setRequestHeader "Content-Length", Len(xmlstr)
req.SetRequestHeader "cookie", reqCadata
req.send xmlstr
wscript.echo req.status
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for c = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(c)),"set-cookie: usercontext=") then reqcadata= right(reqhedrarry(c),len(reqhedrarry(c))-12)
next
for each ent in reqhedrarry
	wscript.echo ent
Next
If InStr(req.responsetext,"name=lngFrm") Then
	wscript.echo "Mailbox has not been logged onto before via OWA"
	'Create a regular expression object
	Dim objRegExp
	Set objRegExp = New RegExp

	'Set our pattern
	objRegExp.Pattern = "<option selected value=""(.*?)"">"
	objRegExp.IgnoreCase = True
	objRegExp.Global = True

	'Get the matches from the contents of our HTML file, strContents
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
		req.SetRequestHeader "cookie", reqCadata
		req.send pstring
		wscript.echo req.status
		If req.status = 200 Then 
			Call UpdateJunk(mbMailbox)
		Else
			wscript.echo "Failed to set Default OWA settings"
		End if

	Else
		wscript.echo "Script failed to retrieve default values"
	End if
Else
	wscript.echo "Junk Mail Setting Updated"
End if
End sub



