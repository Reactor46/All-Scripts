snServername = "servername"
userName = "domain\username"
strpassword = "password"

Targetmailbox = "User@domain.com"
set req = createobject("MSXML2.ServerXMLHTTP.6.0")

Call UpdateJunk(Targetmailbox)

Sub UpdateJunk(mbMailbox)

xmlstr = "<params><fEnbl>1</fEnbl></params>"
req.Open "POST", "https://" & snServername & "/owa/" & mbMailbox & "/ev.owa?oeh=1&ns=JunkEmail&ev=Enable", False,userName,strpassword 
req.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
req.setRequestHeader "Content-Length", Len(xmlstr)
req.send xmlstr
wscript.echo req.status
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
		req.Open "POST", "https://" & snServername & "/owa/" & mbMailbox & "/lang.owa", False,userName,strpassword
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



