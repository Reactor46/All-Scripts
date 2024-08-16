<%
on error resume next
set xmlfile1 = request.querystring("xmlfile")
xmlfile = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
xmlfile = left(xmlfile,(instr(xmlfile,"showmessage.asp")-1)) &  xmlfile1
uid = request.querystring("message")
set xmlobj = server.createobject("microsoft.xmlhttp")
xmlobj.Open "Get", xmlfile, False, "", ""
xmlobj.setRequestHeader "Accept-Language:", "en-us"
xmlobj.setRequestHeader "Content-type:", "text/xml"
xmlobj.Send
set oResponseDoc = xmlobj.responseXML
set oNodeList = oResponseDoc.getElementsByTagName("*")
For i = 0 To (oNodeList.length -2)
	set oNode = oNodeList.nextNode
	if oNode.Text = uid then
  		guid = oNode.Text 
		set oNode = oNodeList.nextNode
		Title = oNode.Text 
		set oNode = oNodeList.nextNode
		Link = oNode.Text 
		set oNode = oNodeList.nextNode
		Description = oNode.Text
		set oNode = oNodeList.nextNode
		Author = oNode.Text
		set oNode = oNodeList.nextNode
		pubdate =  oNode.Text
	end if
Next
pubdate = replace(pubdate,"GMT","")
pubdate = cdate(Mid(pubdate,(instr(pubdate,",")+1),len(pubdate)))
set shell = server.createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
pubdate = dateadd("h",toffset,pubdate)


%><table border="1" cellspacing="1" width="100%" id="table1">
	<tr>
		<td>Received: <%=pubdate%></td>
	</tr>
	<tr>
		<td>From: <%=author%></td>
	</tr>
	<tr>
		<td>Subject: <%=Title%></td>
	</tr>
	<tr>
		<td>Message: <%=Description%></td>
	</tr>
</table>