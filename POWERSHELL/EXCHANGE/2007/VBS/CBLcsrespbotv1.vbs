SipURI = "user@domain.com"
Servername = "servername.domain.com"
userName = "domain\username"
Password = "password"

feedfilename = "E:\inetpub\wwwroot\LCSfeed.xml"
Maxitemnumber = 20

set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
Com.ActiveConnection = Conn
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")

set objdom = CreateObject("MICROSOFT.XMLDOM")
Set fso = CreateObject("Scripting.FileSystemObject")


if fso.FileExists(feedfilename) then
	objdom.async="false"
	objdom.load(feedfilename)
	Set xnItemNodes = objdom.getElementsByTagName("item")
	Set xnChannelNodes = objdom.getElementsByTagName("channel")
	wscript.echo xnItemNodes.length
	if  xnItemNodes.length > Maxitemnumber then
		for i = Maxitemnumber to xnItemNodes.length
			set parentnode = xnItemNodes(i-1).parentnode
			parentnode.removechild(xnItemNodes(i-1))
		next 
	end if
	set xnChannelNode = xnChannelNodes(0)
	fptrack = 0
else
	' ************ Create Root XML Elemements ************************
	fptrack = 1
	Set objField = objDom.createElement("rss")
	Set objattID = objDom.createAttribute("version")
	objattID.Text = "2.0"
	objField.setAttributeNode objattID
	objDom.appendChild objField
	Set xnChannelNode = objDom.createElement("channel")
	objfield.appendChild xnChannelNode
	Set objField4 = objDom.createElement("title")
	objfield4.text = "Company Instant Message Notice Board"
	xnChannelNode.appendChild objField4
	Set objField5 = objDom.createElement("description")
	objfield5.text = "Company Instant Message Notice Board"
	xnChannelNode.appendChild objField5
	Set objField6 = objDom.createElement("language")
	objfield6.text = "en-us"
	xnChannelNode.appendChild objField6
	Set objField7 = objDom.createElement("lastBuildDate")
	objfield7.text = WeekdayName(weekday(now),3) & ", " & day(now()) & " " & Monthname(month(now()),3) & " " & year(now()) & " " & formatdatetime(now(),4) & ":00 GMT"
	xnChannelNode.appendChild objField7
	Set objPI = objDom.createProcessingInstruction("xml", "version='1.0'")
	objDom.insertBefore objPI, objDom.childNodes(0)
	' ************ Create Root XML Elemements ************************
end if

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
		if instr(req.responsetext,"inviters=""") then 
			Invite = mid(req.responsetext,instr(req.responsetext,"inviters=""")+10,instr(instr(req.responsetext,"inviters=""")+10,req.responsetext,chr(34))-(instr(req.responsetext,"inviters=""")+10))
		else
			exist = 1
			elen = instr((instr(instr(req.responsetext,"action=""receive"""),req.responsetext,"uri=""")+5),req.responsetext,chr(34)) - (instr(instr(req.responsetext,"action=""receive"""),req.responsetext,"uri=""")+5)
			Invite = mid(req.responsetext,instr(instr(req.responsetext,"action=""receive"""),req.responsetext,"uri=""")+5,elen)
		end if
		SIPQueryFilter = "(&(objectCategory=person)(msRTCSIP-PrimaryUserAddress=" & Invite & "))"
		strQuery1 = "<LDAP://"  & strDefaultNamingContext & ">;" & SIPQueryFilter & ";DisplayName;subtree"
		com.Properties("Page Size") = 100
		Com.CommandText = strQuery1
		Set Rs1 = Com.Execute
		if rs1.recordcount <> 0 then 
			FromDisplay = rs1.fields("DisplayName")
		else
			FromDisplay = replace(Invite,"sip:","")
		end if
		call CreateItem(message,FromDisplay,fptrack)
		Set xnItemNodes = objdom.getElementsByTagName("item")
		Set xnChannelNodes = objdom.getElementsByTagName("channel")
		wscript.echo xnItemNodes.length
		if  xnItemNodes.length > Maxitemnumber then
			for i = Maxitemnumber to xnItemNodes.length
				set parentnode = xnItemNodes(i-1).parentnode
				parentnode.removechild(xnItemNodes(i-1))
			next 
		end if
	        set xnChannelNode = xnChannelNodes(0)
	        fptrack = 0
		objdom.save(feedfilename)
		SendMessage("Thanks for your message it has been added to the board")
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

Sub CreateItem(PostSubject,PostFrom,Fp)
PostBody = ""
if instr(PostSubject,chr(10)) then
	PostBody = mid(PostSubject,instr(PostSubject,chr(10)),len(postsubject)-instr(PostSubject,chr(10)))
	postbody = replace(postbody,chr(10),"<br>")
	PostSubject = left(PostSubject,instr(PostSubject,chr(10)))
end if
rndval = Int((20000000000 * Rnd) + 1) 
rval = day(now) & month(now) & year(now) & hour(now) & minute(now) & rndval 
Set objField2 = objDom.createElement("item")
if fp = 1 then
	xnChannelNode.appendChild objField2
else
	xnChannelNode.insertbefore objField2, xnItemNodes(0)
end if
Set objField8 = objDom.createElement("guid")
Set objattID8 = objDom.createAttribute("isPermaLink")
objattID8.Text = "false"
objField8.setAttributeNode objattID8
objfield8.text = rval
objfield2.appendChild objField8
Set objField9 = objDom.createElement("title")
objfield9.text = PostSubject
objfield2.appendChild objField9
if PostBody <> "" then
	Set objField11 = objDom.createElement("description")
	objfield11.text = PostBody
	objfield2.appendChild objField11
end if
Set xnChannelNode0 = objDom.createElement("author")
xnChannelNode0.text = PostFrom
objfield2.appendChild xnChannelNode0
Set xnChannelNode1 = objDom.createElement("pubDate")
xnChannelNode1.text = WeekdayName(weekday(now),3) & ", " & day(now()) & " " & Monthname(month(now()),3) & " " & year(now()) & " " & formatdatetime(now(),4) & ":00 GMT"
objfield2.appendChild xnChannelNode1
set objfield2 = nothing
set objfield8 = nothing
set objfield9 = nothing
set xnChannelNode0 = nothing
set xnChannelNode1 = nothing

End Sub
