on error resume Next
fpath = "c:\exp\"
Servername = wscript.arguments(0)
domaintosearch = wscript.arguments(1)
datefrom = wscript.arguments(2) & "T00:00:00Z"
dateto = wscript.arguments(3) & "T00:00:00Z"
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
		falias = "http://" & servername & "/exadmin/admin/" & dpDefaultpolicy & "/mbx/"
		for each paddress in rs1.fields("proxyaddresses").value
			if instr(paddress,"SMTP:") then 
				falias = falias & replace(paddress,"SMTP:","")  
				cusername = replace(paddress,"SMTP:","")
			End if
		next
		ReDim tresarray(1,6)
		wscript.echo  falias 
		call RecurseFolder(falias)
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
set conn = nothing
set com = nothing
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
strQuery = strQuery & """urn:schemas:httpmail:fromemail"",  ""urn:schemas:httpmail:to"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = False AND " 
strQuery = strQuery & """urn:schemas:httpmail:datereceived"" &lt; CAST(""" & dateto & """ as 'dateTime') AND "
strQuery = strQuery & """urn:schemas:httpmail:datereceived"" &gt; CAST(""" & datefrom & """ as 'dateTime')</D:sql></D:searchrequest>"
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
   set fEmail = oResponseDoc.getElementsByTagName("d:fromemail")
   set TEmail = oResponseDoc.getElementsByTagName("d:to")
   For i = 0 To (oNodeList.length -1)
		set oNode = oNodeList.nextNode
		set oNode1 = oNodeList1.nextNode
		set oNode2 = oSize.nextNode
		set oNode3 = odatereceived.nextNode
		set oNode4 = fEmail.nextNode
		set oNode5 = TEmail.nextNode
		wscript.echo oNode3.text
		export = 0
		If InStr(LCase(oNode4.text),LCase(domaintosearch))Then
			export = 1
		End If
		if InStr(LCase(oNode5.text),LCase(domaintosearch))Then
			export = 1
		End If
		If export = 1 Then
			Call exportemail(oNode1.text,oNode.text)
			wscript.echo "Exporting : " & oNode4.text
		End if
	Next
Else
End If

end sub

sub exportemail(exporthref,subject)
req.open "GET", exporthref, false
req.setRequestHeader "Translate","f"
req.send
fname = replace(replace(replace(replace(replace((cusername & "-" & subject),":","-"),"\",""),"/",""),"?",""),chr(34),"")
fname = replace(replace(replace(replace(replace(replace(fname,"<",""),">",""),chr(11),""),"*",""),"|",""),"(","")
fname = replace(replace(replace(fname,")",""),chr(12),""),chr(15),"")
Randomize ' Initialize random-number generator.
rndval = Int((20000000000 * Rnd) + 1) 
fname = fpath & replace(lcase(fname),".eml",rndval & ".eml")
wscript.echo fname
set stm = createobject("ADODB.Stream")
stm.open
msgstring = req.responsetext
stm.type = 2
stm.Charset = "x-ansi"
stm.writetext msgstring,0
stm.Position = 0
stm.type = 1
stm.savetofile fname
set stm = nothing

End sub