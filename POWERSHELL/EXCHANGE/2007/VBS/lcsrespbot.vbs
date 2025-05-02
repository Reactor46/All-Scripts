SipURI = "user@domain.com"
Servername = "servername.domain.com"
userName = "domain\username"
Password = "password"

set req = createobject("microsoft.xmlhttp")
req.Open "GET", "https://" & Servername & "/iwa/logon.html?uri=" & SipURI & "&signinas=1&language=en&epid=", False, Username, Password
req.send
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for c = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(c)),"set-cookie:") then reqsessionID = right(reqhedrarry(c),len(reqhedrarry(c))-12)
next
wscript.echo reqsessionID
chk =  left(reqsessionID,36)
updatestr = "https://" & Servername & "/cwa/AsyncDataChannel.ashx?AckID=0&Ck=" & chk
req.Open "GET", updatestr, False, Username, Password
req.setRequestHeader "Cookie:", reqsessionID 
req.send
latupdate = mid(req.responsetext,instr(req.responsetext,"latestUpdate=")+14,instr(instr(req.responsetext,"latestUpdate=")+14,req.responsetext,chr(34))-(instr(req.responsetext,"latestUpdate=")+14))

while i <> 1
	updatestr = "https://" & Servername & "/cwa/AsyncDataChannel.ashx?AckID=" & latupdate  & "&Ck=" & chk
	req.Open "GET", updatestr, False, Username, Password
	req.setRequestHeader "Cookie:", reqsessionID 
	req.send
	wscript.echo req.status
	if instr(req.responsetext,"div id=""exception""") then i = 1
	oldlat = latupdate
	latupdate = mid(req.responsetext,instr(req.responsetext,"latestUpdate=")+14,instr(instr(req.responsetext,"latestUpdate=")+14,req.responsetext,chr(34))-(instr(req.responsetext,"latestUpdate=")+14))
	wscript.echo req.responsetext
	wscript.echo latupdate
	if latupdate = "yTimeout" then latupdate = oldlat
	if instr(req.responsetext,"message=""") then 
		Imid = mid(req.responsetext,instr(req.responsetext,"imId=""")+6,instr(instr(req.responsetext,"imId=""")+6,req.responsetext,chr(34))-(instr(req.responsetext,"imId=""")+6))
		message = mid(req.responsetext,instr(req.responsetext,"message=""")+9,instr(instr(req.responsetext,"message=""")+9,req.responsetext,chr(34))-(instr(req.responsetext,"message=""")+9))
		wscript.echo "************************Message Recieved****************************"
		wscript.echo message
		wscript.echo "************************Message Ends********************************"
		wscript.echo Imid
		exist = 0
		if instr(1,lcase(message),"what day is it") then 
			SendMessage("Today is " & weekdayname(weekday(now())))
		else
			SendMessage("Im a little simple and can only answer the question what day is it")	
		end if
		message = ""
		Invite = ""
	end if
wend


function SendMessage(message)

' ---Send Message---
Sendmsgcmd =  "https://" & Servername & "/cwa/MainCommandHandler.ashx?Ck=" & chk
Messagestr = "cmdPkg=1,LcwAcceptImRequest," & Imid 
req.open "POST", Sendmsgcmd, False, Username ,password
req.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
req.setRequestHeader "Cookie:", reqsessionID 
req.send Messagestr
wscript.echo req.status
Sendmsgcmd =  "https://" & Servername & "/cwa/MainCommandHandler.ashx?Ck=" & chk
Messagestr = "cmdPkg=2,LcwSendMessageRequest," & Imid & "," & Message & ",X-MMS-IM-Format: FN=Arial%253B EF=%253B CO=000000%253B CS=1%253B PF=00"
req.open "POST", Sendmsgcmd, False, Username ,password
req.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
req.setRequestHeader "Cookie:", reqsessionID 
req.send Messagestr
wscript.echo req.status

end function