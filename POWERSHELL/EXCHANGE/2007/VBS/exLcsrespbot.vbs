SipURI = "user@domain.com"
Servername = "servername.domain.com"
userName = "domain\username"
Password = "password"

set shell = CreateObject("WScript.Shell")
set req = createobject("microsoft.xmlhttp")
set exreq = createobject("microsoft.xmlhttp")
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
Com.ActiveConnection = Conn
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())

req.Open "GET", "https://" & Servername & "/iwa/logon.html?uri=" & SipURI & "&signinas=1&language=en&epid=", False, Username, Password
req.send
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for c = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(c)),"set-cookie:") then reqsessionID = right(reqhedrarry(c),len(reqhedrarry(c))-12)
next
wscript.echo reqsessionID
chk =  left(reqsessionID,36)
updatestr = "https://" & Servername & "/cwa/AsyncDataChannel.ashx?AckID=0&Ck=" & chk
req.Open "GET", updatestr, False, username, password 
req.setRequestHeader "Cookie:", reqsessionID 
req.send
latupdate = mid(req.responsetext,instr(req.responsetext,"latestUpdate=")+14,instr(instr(req.responsetext,"latestUpdate=")+14,req.responsetext,chr(34))-(instr(req.responsetext,"latestUpdate=")+14))

while iloop <> 1
	updatestr = "https://" & Servername & "/cwa/AsyncDataChannel.ashx?AckID=" & latupdate  & "&Ck=" & chk
	req.Open "GET", updatestr, False, Username, Password
	req.setRequestHeader "Cookie:", reqsessionID 
	req.send
	wscript.echo req.status
	if instr(req.responsetext,"div id=""exception""") then 
		iloop = 1
		wscript.echo req.responsetext
	end if
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
		if instr(1,lcase(message),"where is") then 
			if len(message) > 9 then
				wscript.echo Invite & " is looking for " & right(message,len(message)-9)
				SendMessage("Hang on a minute I'll just check for you")
				findperson = whereis(right(message,len(message)-9),Invite)
				if findperson = "" then findperson = "Im Sorry I cant find the person your looking for make sure you use their email address"
				SendMessage(findperson)
			else
				responetemp = "I am a where is response bot to ask me a question you need to use the following format" & vbcrlf 
				responetemp = responetemp & "Where is emailaddress@domain.com"
				SendMessage(responetemp)
			end if
		else 
			responetemp = "I am a where is response bot to ask me a question you need to use the following format" & vbcrlf 
			responetemp = responetemp & "Where is emailaddress@domain.com"
			SendMessage(responetemp)
		end if
		message = ""
		Invite = ""
	end if
wend


function whereis(emailaddress,fromsip)
SIPQueryFilter = "(&(objectCategory=person)(msRTCSIP-PrimaryUserAddress=" & fromsip & "))"
GALQueryFilter =  "(&(objectCategory=person)(proxyAddresses=SMTP:" & emailaddress & "))"
strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,msExchHomeServerName,displayname,legacyExchangeDN,homemdb;subtree"
strQuery1 = "<LDAP://"  & strDefaultNamingContext & ">;" & SIPQueryFilter & ";distinguishedName,proxyAddresses;subtree"
com.Properties("Page Size") = 100
Com.CommandText = strQuery
Set Rs = Com.Execute
Com.CommandText = strQuery1
Set Rs1 = Com.Execute
if rs1.recordcount <> 0 then 
	for each adr in rs1.fields("proxyAddresses").value
		if instr(adr,"SMTP:") then fromemail = adr 
	next
	fromemail = replace(fromemail,"SMTP:","")
end if
while not rs.eof
	legdn = rs.fields("legacyExchangeDN")
	exservername = right(rs.fields("msExchHomeServerName"),len(rs.fields("msExchHomeServerName"))-(instr(rs.fields("msExchHomeServerName"),"cn=Servers/cn=")+13))
	calinfo = CalendarQuery(exservername,emailaddress)
	calinfo = rs.fields("displayname").value &  vbcrlf & calinfo 
	mailinfo = emailQuery(exservername,emailaddress,fromemail)
	Logoninfo = LogonQuery(exservername,emailaddress,legdn)
	rs.movenext
wend
whereis = calinfo & mailinfo & Logoninfo
end function

function CalendarQuery(exservername,Alias)

wscript.echo exservername & "	" & Alias
dtListFrom = DateAdd("n", minTimeOffset, now())
dtListTo = isodateit(DateAdd("h",8,dtListFrom))
dtListFrom = isodateit(dtListFrom)
frun = 1 

strURL = "http://" & exservername & "/exchange/" & Alias & "/calendar/"
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:Displayname"", ""urn:schemas:calendar:dtend"", ""urn:schemas:calendar:dtstart"", ""urn:schemas:httpmail:subject"", "
strQuery = strQuery & """urn:schemas:calendar:location"" FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') WHERE NOT ""urn:schemas:calendar:instancetype"" = 1 AND (""DAV:ishidden"" is Null OR ""DAV:ishidden"" = false) AND (""urn:schemas:calendar:dtend"" &gt;= CAST(""" & dtListFrom & """ as 'dateTime')) "
strQuery = strQuery &  "AND (""urn:schemas:calendar:dtstart"" &lt; CAST(""" & dtListTo & """ as 'dateTime'))" 
strQuery = strQuery &  " ORDER BY ""urn:schemas:calendar:dtstart"" ASC</D:sql></D:searchrequest>"
exreq.open "SEARCH", strURL, false, username, password
exreq.setrequestheader "Content-Type", "text/xml"
exreq.setRequestHeader "Translate","f"
exreq.send strQuery
If exreq.status >= 500 Then
ElseIf exreq.status = 207 Then
   set oResponseDoc = exreq.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:href")
   set oNodedtstartList = oResponseDoc.getElementsByTagName("d:dtstart")
   set oNodedtendList = oResponseDoc.getElementsByTagName("d:dtend")
   set oNodeLocationList = oResponseDoc.getElementsByTagName("d:location")
   set oNodeSubjectList = oResponseDoc.getElementsByTagName("e:subject")
   if oNodeList.length = 0 then calapt = "No Appointments in the next 8 hours" & vbcrlf 
   For i = 0 To (oNodeList.length -1)
	if dateadd("h",toffset,oNodedtstartList(i).nodetypedvalue) < now() then
		calapt = calapt & "Currently is at the following appoinment " & oNodeSubjectList(i).text & vbcrlf 
		calapt = calapt & "This Appointment started at " & dateadd("h",toffset,oNodedtstartList(i).nodetypedvalue) & " and will finish at " & dateadd("h",toffset,oNodedtendList(i).nodetypedvalue) & vbcrlf
		if oNodeLocationList(i).text <> "" then	calapt = calapt & "This Appointment is Located at " & oNodeLocationList(i).text  & vbcrlf
	else
		if frun = 1 then 
			calapt = calapt & vbcrlf & "There Movements for the next 8 hours are" & vbcrlf 
			frun = 0 
		end if
		calapt = calapt & dateadd("h",toffset,oNodedtstartList(i).nodetypedvalue) & " To " & dateadd("h",toffset,oNodedtendList(i).nodetypedvalue) & " " & oNodeSubjectList(i).text & " " & oNodeLocationList(i).text & vbcrlf 
	end if
   Next	
Else
End If
calapt = calapt &  vbcrlf
CalendarQuery = calapt
end function

function EmailQuery(exservername,Alias,fromemail)


dtListFrom = isodateit(DateAdd("d",-3,now()))
strURL = "http://" & exservername & "/exchange/" & Alias & "/inbox/"
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:Displayname"", ""urn:schemas:httpmail:subject"", ""urn:schemas:httpmail:datereceived"" "
strQuery = strQuery & "FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') WHERE (""urn:schemas:httpmail:datereceived"" &gt;= CAST(""" & dtListFrom & """ as 'dateTime')) AND ""urn:schemas:httpmail:fromemail"" = '" & fromemail  & "' "
strQuery = strQuery & " AND ""urn:schemas:httpmail:read"" = FALSE ORDER BY ""urn:schemas:httpmail:datereceived"" DESC</D:sql></D:searchrequest>"
exreq.open "SEARCH", strURL, false, username, password
exreq.setrequestheader "Content-Type", "text/xml"
exreq.setRequestHeader "Range", "rows=0-10"
exreq.setRequestHeader "Translate","f"
exreq.send strQuery
If exreq.status >= 500 Then
ElseIf exreq.status = 207 Then
   set oResponseDoc = exreq.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:href")
   set oNoderecievedList = oResponseDoc.getElementsByTagName("d:datereceived")
   set oNodeSubjectList = oResponseDoc.getElementsByTagName("d:subject")
   if oNodeList.length = 0 then 
	mailrep = "There are no unread emails in there inbox from you over the past 3 days" & vbcrlf & vbcrlf
   else
	mailrep = mailrep & "The following Emails sent by you are still marked as Unread in their Inbox" & vbcrlf & vbcrlf
   end if
   For i = 0 To (oNodeList.length -1)
	mailrep  = mailrep & dateadd("h",toffset,oNoderecievedList(i).nodetypedvalue) & " " & oNodeSubjectList(i).text & vbcrlf
   Next	
   if oNodeList.length <> 0 then mailrep = mailrep & vbcrlf
Else
End If
strURL = "http://" & exservername & "/exchange/" & Alias & "/sent items/"
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:Displayname"", ""urn:schemas:httpmail:subject"", ""urn:schemas:httpmail:datereceived"" "
strQuery = strQuery & "FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """')</D:sql></D:searchrequest>"
exreq.open "SEARCH", strURL, false, username, password
exreq.setrequestheader "Content-Type", "text/xml"
exreq.setRequestHeader "Range", "rows=0-1"
exreq.setRequestHeader "Translate","f"
exreq.send strQuery
If exreq.status >= 500 Then
ElseIf exreq.status = 207 Then
   set oResponseDoc = exreq.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:href")
   set oNoderecievedList = oResponseDoc.getElementsByTagName("d:datereceived")
   set oNodeSubjectList = oResponseDoc.getElementsByTagName("d:subject")
   if oNodeList.length = 0 then 
	mailrep = mailrep &  "No emails appear to have been sent for the past 3 days" & vbcrlf
   else
	mailrep = mailrep &  "The Last Sent Email in users Sent Items is dated "  & dateadd("h",toffset,oNoderecievedList(0).nodetypedvalue) & vbcrlf & vbcrlf
   end if
Else
End If

EmailQuery = mailrep
end function

function Logonquery(exservername,emailaddress,exchangedn)

showarg = "All"

set conn1 = createobject("ADODB.Connection")
strConnString = "Data Provider=NONE; Provider=MSDataShape"
conn1.Open strConnString		
set objParentRS = createobject("adodb.recordset")
set objChildRS = createobject("adodb.recordset")
strSQL = "SHAPE APPEND" & _
           "  NEW adVarChar(255) AS ADDisplayName, " & _
           "  NEW adVarChar(255) AS ADLegacyDN, " & _
           " ((SHAPE APPEND  " & _
           "      NEW adVarChar(255) AS WMIDisplayName, " & _
           "      NEW adVarChar(255) AS WMILegacyDN, " & _
           "      NEW adVarChar(255) AS WMILoggedOnUserAccount, " & _
           "      NEW adVarChar(255) AS WMIClientVersion, " & _
           "      NEW adVarChar(255) AS WMIClientIP, " & _
	   "      NEW adVarChar(255) AS WMILogonTime, " & _		
           "      NEW adVarChar(255) AS WMIClientMode " & _
	   ")" & _
           "      RELATE ADLegacyDN TO WMILegacyDN) AS rsADWMI " 
objParentRS.LockType = 3
objParentRS.Open strSQL, conn1
Com.ActiveConnection = Conn
GALQueryFilter =  "(&(objectCategory=person)(proxyAddresses=SMTP:" & emailaddress & "))"
strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,displayname,legacyExchangeDN,homemdb;subtree"
com.Properties("Page Size") = 100
Com.CommandText = strQuery
Set Rs2 = Com.Execute
while not rs2.eof
	objParentRS.addnew 
	objParentRS("ADDisplayName") = rs2.fields("displayname")
	objParentRS("ADLegacyDN") = rs2.fields("legacyExchangeDN")
	objParentRS.update	
	rs2.movenext
wend
strWinMgmts ="winmgmts:{impersonationLevel=impersonate}!//"& exservername &"/root/MicrosoftExchangeV2"
Select case showarg
	case else sqlstate = "Select * FROM Exchange_Logon where StoreType = 1 and ClientVersion <> 'SMTP' AND ClientVersion <> 'OLEDB' AND MailboxLegacyDN ='" & exchangedn & "'"
end select
Set objWMIExchange = GetObject(strWinMgmts)
Set listExchange_ExchangeLogons = objWMIExchange.ExecQuery(sqlstate,,48)
objChildRS.LockType = 3
Set objChildRS = objParentRS("rsADWMI").Value
For each objExchange_ExchangeLogon in listExchange_ExchangeLogons
	if objExchange_ExchangeLogon.LoggedOnUserAccount <> "NT AUTHORITY\SYSTEM"  then
		if lcase(objExchange_ExchangeLogon.LoggedOnUserAccount) <> lcase(username) then
			objChildRS.addnew 
			objChildRS("WMIDisplayName") = objExchange_ExchangeLogon.MailboxDisplayName
			objChildRS("WMILegacyDN") = objExchange_ExchangeLogon.MailboxLegacyDN
			objChildRS("WMILoggedOnUserAccount") = objExchange_ExchangeLogon.LoggedOnUserAccount
			objChildRS("WMIClientVersion") = objExchange_ExchangeLogon.ClientVersion
			objChildRS("WMIClientIP") = objExchange_ExchangeLogon.ClientIP
			objChildRS("WMILogonTime") = objExchange_ExchangeLogon.LogonTime
			objChildRS("WMIClientMode") = objExchange_ExchangeLogon.ClientMode
			objChildRS.update
		end if
	end if
Next
objParentRS.MoveFirst
logqry = "The Current Email Logon status of the user is" & vbcrlf & vbcrlf
Do While Not objParentRS.EOF
    Set objChildRS = objParentRS("rsADWMI").Value 
    objChildRS.sort = "WMILoggedOnUserAccount"
    if objChildRS.recordcount = 0 then logqry = "User Not Currently Logged On" & vbcrlf
    if showarg <> "loggedout" then
    Do While Not objChildRS.EOF
	currec = objChildRS.fields("WMILoggedOnUserAccount") & objChildRS.fields("WMIClientVersion") & objChildRS.fields("WMIClientIP") & objChildRS.fields("WMIClientMode")
	if currec <> prevrec then 
		ltime = dateadd("h",toffset,cdate(DateSerial(Left(objChildRS.fields("WMILogonTime"), 4), Mid(objChildRS.fields("WMILogonTime"), 5, 2),Mid(objChildRS.fields("WMILogonTime"), 7, 2)) _
		& " " & timeserial(Mid(objChildRS.fields("WMILogonTime"), 9, 2),Mid(objChildRS.fields("WMILogonTime"), 11, 2),Mid(objChildRS.fields("WMILogonTime"),13, 2))))
		logqry = logqry & objChildRS.fields("WMILoggedOnUserAccount") & "   " & getversion(objChildRS.fields("WMIClientVersion")) & "  " &  _
		"Logged on at " & ltime & vbcrlf 
	end if
	prevrec = currec
	objChildRS.MoveNext
    Loop
    end if   		
    objParentRS.MoveNext
Loop
Logonquery = logqry
end function


function getversion(version)

Select case version
	case "HTTP" getversion = "Outlook Web Access"
	case else getversion =  "Outlook"
end select

end function

function getmode(clientmode)
	select case clientmode
		case 1 getmode = "Classic Online"
		case 2 getmode = "Cached Mode"
		case else getmode = " "
	end select
end function

function isodateit(datetocon)
	strDateTime = year(datetocon) & "-"
	if (Month(datetocon) < 10) then strDateTime = strDateTime & "0"
	strDateTime = strDateTime & Month(datetocon) & "-"
	if (Day(datetocon) < 10) then strDateTime = strDateTime & "0"
	strDateTime = strDateTime & Day(datetocon) & "T" & formatdatetime(datetocon,4) & ":00Z"
	isodateit = strDateTime
end function

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