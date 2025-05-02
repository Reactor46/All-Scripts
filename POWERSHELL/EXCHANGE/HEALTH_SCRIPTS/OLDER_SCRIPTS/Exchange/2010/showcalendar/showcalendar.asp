<table border="0" id="table1" cellpadding="2" width="147">
<tr><b>
<% 
sdate = request.querystring("sdate")
if sdate = "" then
	wdate = now()
	mmonth = monthname(month(now())) & " " & Year(now())
else
	wdate = dateserial(mid(sdate,1,4),mid(sdate,6,2),mid(sdate,9,2))
	mmonth = monthname(month(wdate)) & " " & Year(wdate)
end if
pmonth = condate(dateadd("m",-1,wdate))
stime = condate(wdate)
etime = condate(dateadd("m",1,wdate))
response.write "<td style=""padding: 0"" width=""147"" align=""center"" colspan=""7""><b><font face=""Arial Narrow"" size=2 color=""#000080"">"
response.write "<a href=""showcalendar.asp?sdate=" & etime & """><img border=""0"" src=""pg-next.gif"" width=""16"" height=""16"" align=""right""></a>"
response.write "<a href=""showcalendar.asp?sdate=" & pmonth & """><img border=""0"" src=""pg-prev.gif"" width=""16"" height=""16"" align=""left""></a><B>" & mmonth & "</b></td>"
%>
</font></b> </tr>
<tr>
<td style="border-bottom-style: solid; padding: 0" width="16" align="center">
M</td>
<td style="border-bottom-style: solid; padding: 0" width="16" align="center">
T</td>
<td style="border-bottom-style: solid; padding: 0" width="16" align="center">
W</td>
<td style="border-bottom-style: solid; padding: 0" width="16" align="center">
T</td>
<td style="border-bottom-style: solid; padding: 0" width="16" align="center">
F</td>
<td style="border-bottom-style: solid; padding: 0" width="16" align="center">
S</td>
<td style="border-bottom-style: solid; padding: 0" width="16" align="center">
S</td>
</tr>
<%

urlstr = "http://servername/exchange/mailbox/calendar/?Cmd=monthfreebusy&start=" & stime & "T00:00:00+10:00&end=" & etime & "T00:00:00+10:00"
Set Objxml = Server.CreateObject("Microsoft.XMLhttp")
Objxml.Open "Get", urlstr, False, "", ""
Objxml.setRequestHeader "Accept-Language:", "en-us"
Objxml.setRequestHeader "Content-type:", "application/x-www-UTF8-encoded"
Objxml.setRequestHeader "Content-Length:", Len(szXml)
Objxml.Send szXml
set oResponseDoc = Objxml.responseXML
set oNodeList = oResponseDoc.getElementsByTagName("fbdata")
For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	caldata = oNode.text
next
cdatefday = cdate(dateserial(year(wdate),month(wdate),"1"))
sday = weekday(cdatefday,2)
cmonth = month(wdate)
for x = 1 to 6
response.write "<tr>"
for i = 1 to 7
if cmonth = month(cdatefday) then
if sday =< i then
sday = 0
if mid(caldata,day(cdatefday),1) = 1 then 
	response.write "<td style=""padding: 0"" width=""16"" align=""center""><b>" & day(cdatefday) & "</b></td>" 
else
	response.write "<td style=""padding: 0"" width=""16"" align=""center"">" & day(cdatefday) & "</td>" 
end if
cdatefday = dateadd("d",1,cdatefday)
else
response.write "<td style=""padding: 0"" width=""16"" align=""center""> </td>" 
end if
else
response.write "<td style=""padding: 0"" width=""16"" align=""center""> </td>" 
end if
next
response.write "</tr>"
next 


function condate(date2con)
dtcon = date2con
if month(dtcon) < 10 then 
	if day(dtcon) < 10 then
		qdat = year(dtcon) & "-" & "0" & month(dtcon) & "-" & "01"
	else
		qdat = year(dtcon) & "-" & "0" & month(dtcon) & "-" & "01"
	end if 
else
	if day(dtcon) < 10 then
		qdat = year(dtcon) & "-" & month(dtcon) & "-" & "01"
	else
		qdat = year(dtcon) & "-" & month(dtcon) & "-" & "01"
	end if 
end if
condate = qdat 
end function 
%>