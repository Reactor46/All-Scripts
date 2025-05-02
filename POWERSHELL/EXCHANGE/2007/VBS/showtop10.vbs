On Error Resume next
Servername = wscript.arguments(0)
set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
treport = "<table border=""1"" width=""100%"">" & vbcrlf
treport = treport & "  <tr>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Folder</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Date Recieved</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Mail From</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Subject</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Size KB</font></b></td>" & vbcrlf
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
		set conn1 = createobject("ADODB.Connection")
		strConnString = "Data Provider=NONE; Provider=MSDataShape"
		conn1.Open strConnString		
		set objParentRS = createobject("adodb.recordset")
		strSQL = "SHAPE APPEND" & _
				  "  NEW adVarChar(255) AS MailDate, " & _
				  "  NEW adVarChar(255) AS FolderName, " & _
				  "  NEW adVarChar(255) AS MailFrom, " & _
				  "  NEW adVarChar(255) AS Subject, " & _
				  "  NEW adBigInt AS Size"
		objParentRS.LockType = 3
		objParentRS.Open strSQL, conn1
		falias = "http://" & servername & "/exadmin/admin/" & dpDefaultpolicy & "/mbx/"
		for each paddress in rs1.fields("proxyaddresses").value
			if instr(paddress,"SMTP:") then falias = falias & replace(paddress,"SMTP:","")  
		next
		wscript.echo  falias 
		call RecurseFolder(falias)
		objParentRS.Sort = "Size DESC"
		objParentRS.movefirst
		Set fso = CreateObject("Scripting.FileSystemObject")
		set wfile = fso.opentextfile("c:\temp\" & rs1.fields("mail").value & ".htm",2,true) 
		report = ""
		For mrep = 1 To 10
			report = report & "<tr>" & vbcrlf
			report = report & "<td align=""center"">" & objParentRS.fields("FolderName") & "&nbsp;</td>" & vbcrlf
			report = report & "<td align=""center"">" & objParentRS.fields("MailDate") & "&nbsp;</td>" & vbcrlf
			report = report & "<td align=""center"">" & objParentRS.fields("MailFrom") & "&nbsp;</td>" & vbcrlf
			report = report & "<td align=""center"">" & objParentRS.fields("Subject") & "&nbsp;</td>" & vbcrlf
			report = report & "<td align=""center"">" & formatnumber(cdbl(objParentRS.fields("Size"))/1024,2) & "&nbsp;</td>" & vbcrlf
			report = report & "</tr>" & vbcrlf
			objParentRS.moveNext
		next
		objParentRS.close
		Set objParentRS = nothing
		wfile.writeline treport
		wfile.writeline report
		wfile.writeline "</table>"
		wfile.close
		set wfile = Nothing
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
set conn = nothing
set com = nothing

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
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:""  xmlns:b=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"">"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"",  ""urn:schemas:httpmail:subject"", "
strQuery = strQuery & """DAV:creationdate"", ""DAV:getcontentlength"", "
strQuery = strQuery & """urn:schemas:httpmail:fromemail"", ""urn:schemas:httpmail:datereceived"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = False Order by ""DAV:getcontentlength"" DESC</D:sql></D:searchrequest>"
req.open "SEARCH", strURL, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.setRequestHeader "Range", "rows=0-9"
req.send strQuery
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:displayname")
   set oNodeList1 = oResponseDoc.getElementsByTagName("a:href")
   set oSize = oResponseDoc.getElementsByTagName("a:getcontentlength")
   set odatereceived = oResponseDoc.getElementsByTagName("d:datereceived")
   set fromEmail = oResponseDoc.getElementsByTagName("d:fromemail")
   set subject = oResponseDoc.getElementsByTagName("d:subject")
   For i = 0 To (oNodeList.length -1)
		set oNode = oNodeList.nextNode
		set oNode1 = oNodeList1.nextNode
		set oNode2 = oSize.nextNode
		set oNode3 = odatereceived.nextNode
		set onode4 = fromEmail.nextNode
		set onode5 = subject.nextNode
		wscript.echo onode4.text & "	" & onode5.text
		objParentRS.addnew 
		objParentRS("MailDate") = dateadd("h",toffset,DateSerial(Mid(oNode3.text,1,4),Mid(oNode3.text,6,2),Mid(oNode3.text,9,2)) & " " & Mid(oNode3.text,12,8))
		objParentRS("FolderName") =  unescape(Replace(strURL,falias,"")) 
		objParentRS("MailFrom") = onode4.text
		objParentRS("Subject") = Right(onode5.text,255)
		objParentRS("Size") = oNode2.Text
		objParentRS.update	
	Next
Else
End If
end sub
