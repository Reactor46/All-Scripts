Public Const CdoDefaultFolderCalendar = 0
Public Const CdoPR_START_DATE = &H600040
Public Const CdoPR_END_DATE = &H610040
Public Const CdoPR_EMAIL = &H39FE001E 
Set sdMeetOrgs = CreateObject("Scripting.Dictionary")

snOWAServername = "servername.com"
owaMailbox = "username"
domain = "domain"
strpassword = "password"
owOwaURL ="https://" & snOWAServername & "/exchange/" & owaMailbox  & "/Drafts/"

snServername = wscript.arguments(0)
mbMailboxName = wscript.arguments(1)


stStartTime = dateserial("2007","03","11")
etEndTime = dateserial("2007","04","01")



set csCDOSession = CreateObject("MAPI.Session")
pfProfile = "mgnms01" & vbLf & mbMailboxName
csCDOSession.Logon "","",False,True,0,True, pfProfile
set clCalendar = csCDOSession.getdefaultfolder(CdoDefaultFolderCalendar)
set acApptcol = clCalendar.messages
set ofApptFilter = acApptcol.Filter
Set afApptFltFld1 = ofApptFilter.Fields.Add(CdoPR_START_DATE, etEndTime)
Set afApptFltFld2 = ofApptFilter.Fields.Add(CdoPR_END_DATE, stStartTime)
For Each apApptoitment In acApptcol
	
  	wscript.echo apApptoitment.fields(CdoPR_START_DATE).value & " : " & apApptoitment.subject 
	trReportBody = ""
	trReportBody = trReportBody  & "<tr>" & vbcrlf
	trReportBody  = trReportBody  & "<td align=""center"" width=""18%"">" & apApptoitment.fields(CdoPR_START_DATE).value & "&nbsp;</td>" & vbcrlf
	trReportBody  = trReportBody  & "<td align=""center"" width=""18%"">" & apApptoitment.fields(CdoPR_END_DATE).value & "&nbsp;</td>" & vbcrlf
	If apApptoitment.IsRecurring Then
		trReportBody  = trReportBody  & "<td align=""center"" width=""34%"">" & apApptoitment.subject & "</a>&nbsp;</td>" & vbcrlf
	Else
		trReportBody  = trReportBody  & "<td align=""center"" width=""34%""><a href=""outlook:" & Right(apApptoitment.id,140) & """>"  & apApptoitment.subject & "</a>&nbsp;</td>" & vbcrlf
	End if
	trReportBody = trReportBody & "<td align=""center"" width=""15%"">" & apApptoitment.Location & "&nbsp;</td>" & vbcrlf
	If apApptoitment.MeetingStatus <> 0 then 
		Set orOrganizer = apApptoitment.Organizer
		trReportBody = trReportBody & "<td align=""center"" width=""15%"">" & orOrganizer.Name & "&nbsp;</td>" & vbcrlf
		trReportBody  = trReportBody  & "</tr>" & vbcrlf
		If sdMeetOrgs.exists(orOrganizer.Name) Then
			sdMeetOrgs(orOrganizer.Name) = sdMeetOrgs(orOrganizer.Name) & trReportBody
		Else	
			sdMeetOrgs.Add orOrganizer.Name,trReportBody
		End if		
	Else
		trReportBody = trReportBody  & "<td align=""center"">NA&nbsp;</td>" & vbcrlf
		trReportBody  = trReportBody  & "</tr>" & vbcrlf
		If sdMeetOrgs.exists(csCDOSession.currentuser.name) Then
			sdMeetOrgs(csCDOSession.currentuser.name) = sdMeetOrgs(csCDOSession.currentuser.name) & trReportBody
		Else	
			sdMeetOrgs.Add csCDOSession.currentuser.name,trReportBody
		End if
	End if
	
	wscript.echo	
Next
Call  WriteandSendReport

Sub  WriteandSendReport()
vbVerbage = "<p><b><font face=""Arial"" color=""#000080"">Due to change blah blah the following " _ 
& "Meetings and Appointments scheduled between the 11th March and 1st of April may potential be 1" _ 
& "hour incorrect. The following is a list of appointments from your calender that may be " _ 
& "affected its recommended blah blah</font></b></p>"
rpReport = rpReport & vbVerbage  & vbcrlf
rpReport = rpReport &  "<p><b><font face=""Arial"" color=""#000080"">Meeting's and Appointments Organized by You</font></b></p>"  & vbcrlf
rpReport = rpReport & "<table border=""1"" width=""100%"">" & vbcrlf
rpReport = rpReport & "  <tr>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""18%""><b><font color=""#FFFFFF"">Start Time</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""18%""><b><font color=""#FFFFFF"">End time</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""36%""><b><font color=""#FFFFFF"">Subject</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""15%""><b><font color=""#FFFFFF"">Location</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""15%""><b><font color=""#FFFFFF"">Organizer</font></b></td>" & vbcrlf
rpReport = rpReport & "</tr>" & vbcrlf
rpReport = rpReport & sdMeetOrgs(csCDOSession.currentuser.name)
rpReport = rpReport & "</table>" & vbcrlf
rpReport = rpReport & "<p><b><font face=""Arial"" color=""#000080"">Meeting You are Scheduled to Attended</font></b></p>" & vbcrlf

For Each kyOrg In sdMeetOrgs.Keys
	If kyOrg <> csCDOSession.currentuser.name Then 
		rpReport = rpReport & "<p><b><font face=""Arial"" color=""#000080"">Organized By : " & kyOrg & "</font></b></p>"
		rpReport = rpReport & "<table border=""1"" width=""100%"">" & vbcrlf
		rpReport = rpReport & sdMeetOrgs(kyOrg)
		rpReport = rpReport & "</table>" & vbcrlf
	End if
Next

wscript.echo csCDOSession.currentuser.fields(CdoPR_EMAIL).value

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

Set req = CreateObject("Microsoft.XMLhttp")

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
szXml = szXml & "MsgTo=" & csCDOSession.currentuser.fields(CdoPR_EMAIL).value & vbLf 
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