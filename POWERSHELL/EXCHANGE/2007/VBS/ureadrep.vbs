on error resume Next
set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())

Servername = wscript.arguments(0)
treport = "<table border=""1"" width=""100%"">" & vbcrlf
treport = treport & "  <tr>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Mailbox Name</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080"" colspan=""2""><b><font color=""#FFFFFF"">Less Than 6 Months</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080"" colspan=""2""><b><font color=""#FFFFFF"">Greator Than 6 Months</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">OOF Status</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Last Sent</font></b></td>" & vbcrlf
treport = treport & "</tr>" & vbcrlf
treport = treport & "  <tr>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">&nbsp;</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">#Messages</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Size(MB)</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">#Messages</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Size(MB)</font></b></td>" & vbcrlf
treport = treport & "</tr>" & vbcrlf
set req = createobject("microsoft.xmlhttp")
set com = createobject("ADODB.Command")
set conn = createobject("ADODB.Connection")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
polQuery = "<LDAP://" & strNameingContext &  ">;(&(objectCategory=msExchRecipientPolicy)(cn=Default Policy));distinguishedName,gatewayProxy;subtree"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = polQuery
Set plRs = Com.Execute
while not plRs.eof
	for each adrobj in plrs.fields("gatewayProxy").value
		if instr(adrobj,"SMTP:") then dpDefaultpolicy = right(adrobj,(len(adrobj)-instr(adrobj,"@")))
	next
	plrs.movenext
wend
wscript.echo dpDefaultpolicy 
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof	
	GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(msExchHomeServerName=" & rs.fields("legacyExchangeDN") & ")) )))))"
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";displayname,mail,distinguishedName,mailnickname,proxyaddresses;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		report = "<table border=""1"" width=""100%"">" & vbcrlf
		report = report & "  <tr>" & vbcrlf
		report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Folder Name</font></b></td>" & vbcrlf
		report = report & "<td align=""center"" bgcolor=""#000080"" colspan=""2""><b><font color=""#FFFFFF"">Less Than 6 Months</font></b></td>" & vbcrlf
		report = report & "<td align=""center"" bgcolor=""#000080"" colspan=""2""><b><font color=""#FFFFFF"">Greator Than 6 Months</font></b></td>" & vbcrlf	
		report = report & "</tr>" & vbcrlf
		report = report & "  <tr>" & vbcrlf
		report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">&nbsp;</font></b></td>" & vbcrlf
		report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">#Messages</font></b></td>" & vbcrlf
		report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Size(MB)</font></b></td>" & vbcrlf
		report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">#Messages</font></b></td>" & vbcrlf
		report = report & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Size(MB)</font></b></td>" & vbcrlf	
		report = report & "</tr>" & vbcrlf
		falias = "http://" & servername & "/exadmin/admin/" & dpDefaultpolicy & "/mbx/"
		for each paddress in rs1.fields("proxyaddresses").value
			if instr(paddress,"SMTP:") then falias = falias & replace(paddress,"SMTP:","")  
		next
		ReDim tresarray(1,6)
		wscript.echo  falias 
		offstat = chkoof(falias & "/" & "non_ipm_subtree/")
		call RecurseFolder(falias)
		LastSent = getlastsent(falias)
		report = report & "</table>" & vbcrlf
		Set fso = CreateObject("Scripting.FileSystemObject")
		set wfile = fso.opentextfile("c:\temp\" & rs1.fields("mail").value & ".htm",2,true) 
		wfile.write report
		wfile.close
		set wfile = nothing
		treport = treport & "<tr>" & vbcrlf
		treport = treport & "<td align=""center"">" & rs1.fields("mail").value & "&nbsp;</td>" & vbcrlf
		treport = treport & "<td align=""center"">" & tresarray(0,1) & "&nbsp;</td>" & vbcrlf
		treport = treport & "<td align=""center"">" & FormatNumber(tresarray(1,1)/1024/1024,2) & "&nbsp;</td>" & vbcrlf
		treport = treport & "<td align=""center"">" &  tresarray(0,2) & "&nbsp;</td>" & vbcrlf
		treport = treport & "<td align=""center"">" & FormatNumber(tresarray(1,2)/1024/1024,2) & "&nbsp;</td>" & vbcrlf
		treport = treport & "<td align=""center"">" & offstat & "&nbsp;</td>" & vbcrlf
		treport = treport & "<td align=""center"">" & LastSent & "&nbsp;</td>" & vbcrlf
		treport = treport & "</tr>" & vbcrlf
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
set conn = nothing
set com = nothing
treport = treport & "</table>" & vbcrlf
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\temp\Server-" & Servername & "-UnreadReport.htm",2,true) 
wfile.write treport
wfile.close
set wfile = nothing
set fso = nothing

Public Sub RecurseFolder(sUrl)
  
   req.open "SEARCH", sUrl, False, "", ""
   sQuery = "<?xml version=""1.0""?>"
   sQuery = sQuery & "<g:searchrequest xmlns:g=""DAV:"">"
   sQuery = sQuery & "<g:sql>SELECT ""http://schemas.microsoft.com/"
   sQuery = sQuery & "mapi/proptag/x0e080003"", ""DAV:hassubs"" FROM SCOPE "
   sQuery = sQuery & "('SHALLOW TRAVERSAL OF """ & sUrl & """') "
   sQuery = sQuery & "WHERE ""DAV:isfolder"" = true and ""DAV:ishidden"" = false and ""http://schemas.microsoft.com/mapi/proptag/x36010003"" = 1"
   sQuery = sQuery & "</g:sql>"
   sQuery = sQuery & "</g:searchrequest>"
   req.setRequestHeader "Content-Type", "text/xml"
   req.setRequestHeader "Translate", "f"
   req.setRequestHeader "Depth", "0"
   req.setRequestHeader "Content-Length", "" & Len(sQuery)
   req.send sQuery
   Set oXMLDoc = req.responseXML
   Set oXMLSizeNodes = oXMLDoc.getElementsByTagName("d:x0e080003")
   Set oXMLHREFNodes = oXMLDoc.getElementsByTagName("a:href")
   Set oXMLHasSubsNodes = oXMLDoc.getElementsByTagName("a:hassubs")
   For i = 0 to oXMLSizeNodes.length - 1
      call procfolder(oXMLHREFNodes.Item(i).nodeTypedValue,sUrl)
      wscript.echo oXMLHREFNodes.Item(i).nodeTypedValue
      If oXMLHasSubsNodes.Item(i).nodeTypedValue = True Then
         call RecurseFolder(oXMLHREFNodes.Item(i).nodeTypedValue)
      End If
   Next
End Sub

sub procfolder(strURL,pfname)
wscript.echo strURL
ReDim resarray(1,6)
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:""  xmlns:b=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"">"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"",  ""urn:schemas:httpmail:subject"", "
strQuery = strQuery & """DAV:creationdate"", ""DAV:getcontentlength"", "
strQuery = strQuery & """urn:schemas:httpmail:fromemail"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = False AND ""urn:schemas:httpmail:read"" = False</D:sql></D:searchrequest>"
req.open "SEARCH", strURL, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.send strQuery
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:displayname")
   set oNodeList1 = oResponseDoc.getElementsByTagName("a:href")
   set oSize = oResponseDoc.getElementsByTagName("a:getcontentlength")
   set odatereceived = oResponseDoc.getElementsByTagName("a:creationdate")
   For i = 0 To (oNodeList.length -1)
		set oNode = oNodeList.nextNode
		set oNode1 = oNodeList1.nextNode
		set oNode2 = oSize.nextNode
		set oNode3 = odatereceived.nextNode
		wscript.echo oNode3.text
		If CDate(DateSerial(mid(oNode3.text,1,4), mid(oNode3.text,6,2),mid(oNode3.text,9,2))) > dateadd("m",-6,now()) Then
			resarray(0,1) = resarray(0,1) + 1
			resarray(1,1) = resarray(1,1) + Int(oNode2.text)
		End if
		If CDate(DateSerial(mid(oNode3.text,1,4), mid(oNode3.text,6,2),mid(oNode3.text,9,2))) < dateadd("m",-6,now()) Then
			resarray(0,2) = resarray(0,2) + 1
			resarray(1,2) = resarray(1,2) + Int(oNode2.text)
		End if
	Next
Else
End If
tresarray(0,1) = tresarray(0,1) + resarray(0,1)
tresarray(1,1) = tresarray(1,1) + resarray(1,1)
tresarray(0,2) = tresarray(0,2) + resarray(0,2)
tresarray(1,2) = tresarray(1,2) + resarray(1,2)
report = report & "<tr>" & vbcrlf
report = report & "<td align=""center"">" & unescape(Replace(strURL,falias,"")) & "&nbsp;</td>" & vbcrlf
report = report & "<td align=""center"">" & resarray(0,1) & "&nbsp;</td>" & vbcrlf
report = report & "<td align=""center"">" & FormatNumber(resarray(1,1)/1024/1024,2) & "&nbsp;</td>" & vbcrlf
report = report & "<td align=""center"">" &  resarray(0,2) & "&nbsp;</td>" & vbcrlf
report = report & "<td align=""center"">" & FormatNumber(resarray(1,2)/1024/1024,2) & "&nbsp;</td>" & vbcrlf
report = report & "</tr>" & vbcrlf
end sub

Function chkoof(mburl)

xmlstr = "<?xml version='1.0' encoding='UTF-8' ?>" _ 
& "<a:propfind xmlns:a='DAV:' xmlns:b='urn:schemas-microsoft-com:datatypes'>" _ 
& "<a:prop xmlns:d='http://schemas.microsoft.com/exchange/'>" _ 
& "<d:oof-state/>" _ 
& "</a:prop>" _ 
& "</a:propfind>" 
req.open "PROPFIND", mburl, false
req.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8""" 
req.setRequestHeader "Depth", "0" 
req.setRequestHeader "Translate", "f" 
req.send xmlstr 
set oResponseDoc = req.responseXML 
set osubNodeList = oResponseDoc.getElementsByTagName("d:oof-state/") 
ntar = osubNodeList.length -1 
wscript.Echo "********** OOf-State :" & osubNodeList.Item(ntar).text 
chkoof = osubNodeList.Item(ntar).text
End Function 

function getlastsent(mburl)
getlastsent = "N/A"
strURL = mburl & "/Sent Items/"
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:""  xmlns:b=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"">"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"",  ""urn:schemas:httpmail:subject"", "
strQuery = strQuery & """DAV:creationdate"", ""DAV:getcontentlength"", "
strQuery = strQuery & """urn:schemas:httpmail:fromemail"", ""urn:schemas:mailheader:date"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = False</D:sql></D:searchrequest>"
req.open "SEARCH", strURL, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.setRequestHeader "Range", "rows=0-0"
req.send strQuery
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("e:date")
   For i = 0 To (oNodeList.length -1)
		set oNode = oNodeList.nextNode
		wscript.echo "*******Last Sent :" & oNode.text
		getlastsent = dateadd("h",toffset,DateSerial(Mid(oNode.text,1,4),Mid(oNode.text,6,2),Mid(oNode.text,9,2)) & " " & Mid(oNode.text,12,8))
   Next
Else
End If
end function
