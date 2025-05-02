Mailboxurl = "http://server/exchange/mailbox/inbox"

Emailreport = "<table border=""1"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbcrlf
Emailreport = Emailreport & "<table border=""1"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbcrlf
Emailreport = Emailreport & "  <tr>" & vbcrlf
Emailreport = Emailreport & "<td align=""center"">Time</td>" & vbcrlf
Emailreport = Emailreport & "<td align=""center"">From</td>" & vbcrlf
Emailreport = Emailreport & "<td align=""center"">Subject</td>" & vbcrlf
Emailreport = Emailreport & "<td align=""center"">Attachment</td>" & vbcrlf
Emailreport = Emailreport & "</tr>" & vbcrlf

strComputer = "."
set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
dtListFrom = DateAdd("n", minTimeOffset, now())
dtListFrom = DateAdd("d",-1,dtListFrom)

set rs = createobject("ADODB.Recordset")
set conn = createobject("ADODB.Connection")
conn.Provider = "ExOLEDB.Datasource"
conn.Open Mailboxurl, "", "", -1 
Set Rs.ActiveConnection = conn
Rs.Source = "SELECT ""DAV:href"", " & _
" ""urn:schemas:mailheader:subject"", " & _
" ""urn:schemas:httpmail:fromname"", " & _
" ""urn:schemas:httpmail:datereceived"", " & _
" ""urn:schemas:httpmail:hasattachment"", " & _
" ""http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"" " & _
"FROM scope('shallow traversal of """ & Mailboxurl & """') " & _
"WHERE (""urn:schemas:httpmail:datereceived"" >= CAST(""" & isodateit(dtListFrom) & """ as 'dateTime'))" & _
" AND ""DAV:contentclass"" = 'urn:content-classes:message'"
Rs.Open
If Not (Rs.EOF) Then
	Rs.MoveFirst
	Do Until Rs.EOF
  	   Emailreport = Emailreport & "  <tr>" & vbcrlf
  	   Emailreport = Emailreport & "<td align=""center"">" & formatdatetime(dateadd("h",toffset,rs.fields("urn:schemas:httpmail:datereceived"))) & "</td>" & vbcrlf
  	   Emailreport = Emailreport & "<td align=""center"">" & rs.fields("urn:schemas:httpmail:fromname") & "</td>" & vbcrlf
  	   Emailreport = Emailreport & "<td align=""center""><a href=""outlook:" & Octenttohex(rs.fields("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102")) & """>"  & rs.fields("urn:schemas:mailheader:subject") & "</a></td>" & vbcrlf
  	   Emailreport = Emailreport & "<td align=""center"">" & rs.fields("urn:schemas:httpmail:hasattachment") & "</td>" & vbcrlf
  	   Emailreport = Emailreport & "</tr>" & vbcrlf
  	   rs.movenext
	loop
end if


Emailreport = Emailreport & "</table>" & vbcrlf
REm Email Bit
Set objEmail = CreateObject("CDO.Message")
objEmail.From = "source@yourdomain.com"
objEmail.To = "target@yourdomain.com"
objEmail.Subject = "Email Summary for " & formatdatetime(now(),2) 
objEmail.htmlbody = Emailreport
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "servername"
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Send

wscript.echo "Report Sent"

function isodateit(datetocon)
	strDateTime = year(datetocon) & "-"
	if (Month(datetocon) < 10) then strDateTime = strDateTime & "0"
	strDateTime = strDateTime & Month(datetocon) & "-"
	if (Day(datetocon) < 10) then strDateTime = strDateTime & "0"
	strDateTime = strDateTime & Day(datetocon) & "T" & formatdatetime(datetocon,4) & ":00Z"
	isodateit = strDateTime
end function

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