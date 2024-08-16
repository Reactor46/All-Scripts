snServername = wscript.arguments(0)
mnMailbox = wscript.arguments(1)
Set sdMeetOrgs = CreateObject("Scripting.Dictionary")
datefrom = "2007-03-11T00:00:00Z"
dateto = "2007-04-01T00:00:00Z"

snOWAServername = "servername.com"
owaMailbox = "username"
domain = "domain"
strpassword = "password"
owOwaURL ="https://" & snOWAServername & "/exchange/" & owaMailbox  & "/Drafts/"

set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())

set req = createobject("microsoft.xmlhttp")
set com = createobject("ADODB.Command")
set conn = createobject("ADODB.Connection")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
polQuery = "<LDAP://" & strNameingContext &  ">;(&(objectCategory=msExchRecipientPolicy)(cn=Default Policy));distinguishedName,gatewayProxy;subtree"
Com.ActiveConnection = Conn
Com.CommandText = polQuery
Set plRs = Com.Execute
while not plRs.eof
	for each adrobj in plrs.fields("gatewayProxy").value
		if instr(adrobj,"SMTP:") then dpDefaultpolicy = right(adrobj,(len(adrobj)-instr(adrobj,"@")))
	next
	plrs.movenext
wend
DnameQuery = "<LDAP://" & strDefaultNamingContext &  ">;(mailnickname=" & mnMailbox & ");distinguishedName,DisplayName,mail;subtree"
Com.ActiveConnection = Conn
Com.CommandText = DnameQuery
Set dsRs = Com.Execute
while not dsRs.eof
	dnDisplayName = dsRs.fields("DisplayName")
	emEmailaddress =  dsRs.fields("mail")
	dsRs.movenext
Wend
wscript.echo dnDisplayName
mbMailboxURI = "http://" & snServername & "/exadmin/admin/" & dpDefaultpolicy & "/mbx/" & mnMailbox & "/Calendar/"
wscript.echo mbMailboxURI 
call procfolder(mbMailboxURI)

sub procfolder(strURL)
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:""  xmlns:b=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"">"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"",  ""urn:schemas:httpmail:subject"", "
strQuery = strQuery & """DAV:creationdate"", "
strQuery = strQuery & """http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"" As EntryID, "
strQuery = strQuery & """urn:schemas:httpmail:fromname"",  ""urn:schemas:calendar:dtstart"", ""urn:schemas:calendar:dtend"","
strQuery = strQuery & " ""urn:schemas:calendar:location"", ""http://schemas.microsoft.com/mapi/apptstateflags"" FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:contentclass"" = 'urn:content-classes:appointment' AND "
strQuery = strQuery & " NOT ""urn:schemas:calendar:instancetype"" = 1 AND " 
strQuery = strQuery & """urn:schemas:calendar:dtstart"" &lt;= CAST(""" & dateto & """ as 'dateTime') AND "
strQuery = strQuery & """urn:schemas:calendar:dtend"" &gt;= CAST(""" & datefrom & """ as 'dateTime')</D:sql></D:searchrequest>"
req.open "SEARCH", strURL, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.send strQuery
wscript.echo req.status
If req.status >= 500 Then
	wscript.echo "Error: " & req.responsetext
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oDisplayNameNodes = oResponseDoc.getElementsByTagName("a:displayname")
   set oHrefNodes = oResponseDoc.getElementsByTagName("a:href")
   set oSubject = oResponseDoc.getElementsByTagName("d:subject")
   set oEndTime = oResponseDoc.getElementsByTagName("e:dtend")
   Set oStartTime = oResponseDoc.getElementsByTagName("e:dtstart")
   Set oLocation = oResponseDoc.getElementsByTagName("e:location")
   Set oAppstate = oResponseDoc.getElementsByTagName("f:apptstateflags")
   Set oFromname = oResponseDoc.getElementsByTagName("d:fromname")
   Set oEntryID = oResponseDoc.getElementsByTagName("EntryID")
   For i = 0 To (oDisplayNameNodes.length -1)
		set oNode = oDisplayNameNodes.nextNode
		set oNode1 = oHrefNodes.nextNode
		set oNode2 = oSubject.nextNode
		set oNode3 = oEndTime.nextNode
		Set oNode4 = oStarttime.nextNode
		Set oNode5 = oLocation.nextNode
		Set oNode6 = oAppstate.nextNode
		Set oNode7 = oFromname.nextNode
		Set oNode8 = oEntryID.nextNode
		wscript.echo oNode2.text
		wscript.echo oNode3.text
		wscript.echo oNode4.text
		wscript.echo oNode5.text
		wscript.echo oNode6.text
		wscript.echo Octenttohex(oNode8.nodeTypedValue)
		soOrgnizer = ""
		soOrgnizer = oNode7.text
		sdStartDate = dateadd("h",toffset,DateSerial(Mid(oNode4.text,1,4),Mid(oNode4.text,6,2),Mid(oNode4.text,9,2)) & " " & Mid(oNode4.text,12,8))
		edEndDate = dateadd("h",toffset,DateSerial(Mid(oNode3.text,1,4),Mid(oNode3.text,6,2),Mid(oNode3.text,9,2)) & " " & Mid(oNode3.text,12,8))
		wscript.echo soOrgnizer
		wscript.echo 
		trReportBody = ""
		trReportBody = trReportBody  & "<tr>" & vbcrlf
		trReportBody  = trReportBody  & "<td align=""center"" width=""20%"">" & sdStartDate &  "&nbsp;</td>" & vbcrlf
		trReportBody  = trReportBody  & "<td align=""center"" width=""20%"">" & edEndDate & "&nbsp;</td>" & vbcrlf
		trReportBody  = trReportBody  & "<td align=""center"" width=""30%""><a href=""outlook:" & Octenttohex(oNode8.nodeTypedValue) & """>"  & oNode2.text & "</a>&nbsp;</td>" & vbcrlf
		trReportBody = trReportBody & "<td align=""center"" width=""15%"">" & oNode5.text & "&nbsp;</td>" & vbcrlf
		If oNode6.text <> 0 then 
			trReportBody = trReportBody & "<td align=""center"" width=""15%"">" & soOrgnizer & "&nbsp;</td>" & vbcrlf
			trReportBody  = trReportBody  & "</tr>" & vbcrlf
			If sdMeetOrgs.exists(soOrgnizer) Then
				sdMeetOrgs(soOrgnizer) = sdMeetOrgs(soOrgnizer) & trReportBody
			Else	
				sdMeetOrgs.Add soOrgnizer,trReportBody
			End if		
		Else
			trReportBody = trReportBody  & "<td align=""center"">NA&nbsp;</td>" & vbcrlf
			trReportBody  = trReportBody  & "</tr>" & vbcrlf
			If sdMeetOrgs.exists(dnDisplayName) Then
				sdMeetOrgs(dnDisplayName) = sdMeetOrgs(dnDisplayName) & trReportBody
			Else	
				sdMeetOrgs.Add dnDisplayName,trReportBody
			End If
		End if
	Next
Else
End If

Call WriteandSendReport()

end sub

Sub  WriteandSendReport()
vbVerbage = "<p><b><font face=""Arial"" color=""#000080"">Due to change blah blah the following " _ 
& "Meetings and Appointments scheduled between the 11th March and 1st of April may potential be 1" _ 
& "hour incorrect. The following is a list of appointments from your calender that may be " _ 
& "affected its recommended blah blah</font></b></p>"
rpReport = rpReport & vbVerbage  & vbcrlf
rpReport = rpReport &  "<p><b><font face=""Arial"" color=""#000080"">Meeting's and Appointments Organized by You</font></b></p>"  & vbcrlf
rpReport = rpReport & "<table border=""1"" width=""100%"">" & vbcrlf
rpReport = rpReport & "  <tr>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""20%""><b><font color=""#FFFFFF"">Start Time</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""20%""><b><font color=""#FFFFFF"">End time</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""30%""><b><font color=""#FFFFFF"">Subject</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""15%""><b><font color=""#FFFFFF"">Location</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""15%""><b><font color=""#FFFFFF"">Organizer</font></b></td>" & vbcrlf
rpReport = rpReport & "</tr>" & vbcrlf
rpReport = rpReport & sdMeetOrgs(dnDisplayName)
rpReport = rpReport & "</table>" & vbcrlf
rpReport = rpReport & "<p><b><font face=""Arial"" color=""#000080"">Meeting You are Scheduled to Attended</font></b></p>" & vbcrlf

For Each kyOrg In sdMeetOrgs.Keys
	If kyOrg <> dnDisplayName Then 
		rpReport = rpReport & "<p><b><font face=""Arial"" color=""#000080"">Organized By : " & kyOrg & "</font></b></p>"
		rpReport = rpReport & "<table border=""1"" width=""100%"">" & vbcrlf
		rpReport = rpReport & sdMeetOrgs(kyOrg)
		rpReport = rpReport & "</table>" & vbcrlf
	End if
Next


'Set objEmail = CreateObject("CDO.Message")
'objEmail.From = "user@domain"
'objEmail.To = "user@domain"
'objEmail.Subject = "Appointment Summary for DST change"
'objEmail.htmlbody = rpReport
'objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "servername"
'objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'objEmail.Configuration.Fields.Update
'objEmail.Send

strusername =  domain & "\" & owaMailbox 
szXml = "destination=https://" & snOWAServername & "/exchange&flags=0&username=" & strusername 
szXml = szXml & "&password=" & strpassword & "&SubmitCreds=Log On&forcedownlevel=0&trusted=0"
req.Open "post", "https://" & snOWAServername & "/exchweb/bin/auth/owaauth.dll", False
req.send szXml
reqhedrarry = split(req.GetAllResponseHeaders(), vbCrLf,-1,1)
for i = lbound(reqhedrarry) to ubound(reqhedrarry)
	if instr(lcase(reqhedrarry(i)),"set-cookie: sessionid=") then reqsessionID = right(reqhedrarry(i),len(reqhedrarry(i))-12)
	if instr(lcase(reqhedrarry(i)),"set-cookie: cadata=") then reqcadata= right(reqhedrarry(i),len(reqhedrarry(i))-12)
next

szXml = "" 
szXml = szXml & "Cmd=send" & vbLf 
szXml = szXml & "MsgTo=" &  emEmailaddress & vbLf
szXml = szXml & "MsgCc=" & vbLf 
szXml = szXml & "MsgBcc=" & vbLf 
szXml = szXml & "urn:schemas:httpmail:importance=1" & vbLf 
szXml = szXml & "http://schemas.microsoft.com/exchange/sensitivity-long=" & vbLf 
szXml = szXml & "urn:schemas:httpmail:subject=Appointment Summary for DST change" & vbLf 
szXml = szXml & "urn:schemas:httpmail:htmldescription=<!DOCTYPE HTML PUBLIC " _ 
& """-//W3C//DTD HTML 4.0 Transitional//EN""><HTML DIR=ltr><HEAD><META HTTP-EQUIV" _ 
& "=""Content-Type"" CONTENT=""text/html; charset=utf-8""></HEAD><BODY><DIV>" _ 
& "<FONT face='Arial' color=#000000 size=2>" & rpReport & "</font>" _ 
& "</DIV></BODY></HTML>" & vbLf 

req.Open "POST", owOwaURL, False, "", ""
req.setRequestHeader "Accept-Language:", "en-us"
req.setRequestHeader "Content-type:", "application/x-www-UTF8-encoded"
req.SetRequestHeader "cookie", reqsessionID
req.SetRequestHeader "cookie", reqCadata
req.setRequestHeader "Content-Length:", Len(szXml)
req.Send szXml
Wscript.echo req.responseText 

wscript.echo "Report Sent"

End Sub

Function Octenttohex(OctenArry) 
ReDim aOut(UBound(OctenArry)) 
For i = 1 to UBound(OctenArry) + 1 
	if len(hex(ascb(midb(OctenArry,i,1)))) = 1 then 
		aOut(i-1) = "0" & hex(ascb(midb(OctenArry,i,1)))
	else
		aOut(i-1) = hex(ascb(midb(OctenArry,i,1)))
	end if
Next 
Octenttohex = join(aOUt,"")
End Function 