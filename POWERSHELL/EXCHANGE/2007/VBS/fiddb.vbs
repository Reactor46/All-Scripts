rem on error resume Next

Servername = wscript.arguments(0)
csCurrentdbFileName = "c:\temp\fiddb.xml"

Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile(csCurrentdbFileName,2,true) 
wfile.writeline("<?xml version=""1.0""?>")
wfile.writeline("<SnappedFIDS SnapDate=""" & WeekdayName(weekday(now),3) & ", " & day(now()) & " " & Monthname(month(now()),3) & " " & year(now()) & " " & formatdatetime(now(),4) & ":00" & """>")

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
		wfile.writeline("<Mailbox displayName=""" & rs1.fields("mail").value & """>")
		for each paddress in rs1.fields("proxyaddresses").value
			if instr(paddress,"SMTP:") then falias = falias & replace(paddress,"SMTP:","")  & "/non_ipm_subtree/"
		Next
		Call  GetRootFolder(falias)
		call RecurseFolder(falias)
		wfile.writeline("</Mailbox>")
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
set conn = nothing
set com = nothing
wfile.writeline("</SnappedFIDS>")

Public Sub GetRootFolder(sUrl)

xmlreqtxt = "<?xml version='1.0'?><a:propfind xmlns:a='DAV:' xmlns:e='http://schemas.microsoft.com/exchange/'><a:prop><e:permanenturl/></a:prop></a:propfind>"
req.open "PROPFIND", sUrl, false , "", ""
req.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
req.setRequestHeader "Depth", "0"
req.setRequestHeader "Translate", "f"
req.send xmlreqtxt
set oResponseDoc = req.responseXML
set oNodeList = oResponseDoc.getElementsByTagName("d:permanenturl")
For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	 wfile.writeline("<Folder Name=""NON_IPM_SUBTREE/Root"" Path=""Root"" fid=""1"&  Mid(oNode.text,InStr(Len(oNode.text)-8,oNode.text,"-"),10) & """></Folder>")
Next

End sub


Public Sub RecurseFolder(sUrl)
  
   req.open "SEARCH", sUrl, False, "", ""
   sQuery = "<?xml version=""1.0""?>"
   sQuery = sQuery & "<g:searchrequest xmlns:g=""DAV:"">"
   sQuery = sQuery & "<g:sql>SELECT ""DAV:displayname"", ""http://schemas.microsoft.com/"
   sQuery = sQuery & "mapi/proptag/x6707001E"", ""http://schemas.microsoft.com/exchange/permanenturl"", ""DAV:hassubs"" FROM SCOPE "
   sQuery = sQuery & "('SHALLOW TRAVERSAL OF """ & sUrl & """') "
   sQuery = sQuery & "WHERE ""DAV:isfolder"" = true and NOT ""http://schemas.microsoft.com/mapi/proptag/x36010003"" = 3"
   sQuery = sQuery & "</g:sql>"
   sQuery = sQuery & "</g:searchrequest>"
   req.setRequestHeader "Content-Type", "text/xml"
   req.setRequestHeader "Translate", "f"
   req.setRequestHeader "Depth", "0"
   req.setRequestHeader "Content-Length", "" & Len(sQuery)
   req.send sQuery
   Set oXMLDoc = req.responseXML
   Set oXMLDavDisplayName = oXMLDoc.getElementsByTagName("a:displayname")
   Set oXMLHREFNodes = oXMLDoc.getElementsByTagName("a:href")
   Set oXMLHasSubsNodes = oXMLDoc.getElementsByTagName("a:hassubs")
   Set oXMLFIDNodes = oXMLDoc.getElementsByTagName("e:permanenturl")
   Set oXMLPathNodes = oXMLDoc.getElementsByTagName("d:x6707001E")
   For i = 0 to oXMLHREFNodes.length - 1
      wscript.echo oXMLHREFNodes.Item(i).nodeTypedValue
      wscript.echo oXMLDavDisplayName(i).nodeTypedValue & "	" & oXMLPathNodes(i).nodeTypedValue
      if  oXMLPathNodes(i).nodeTypedValue = "/" then
	strDispName = "root"
      else
	strDispName = oXMLDavDisplayName(i).nodeTypedValue
      end if
      wscript.echo  Mid(oXMLFIDNodes(i).text,InStr(Len(oXMLFIDNodes(i).text)-8,oXMLFIDNodes(i).text,"-"),10)
      wfile.writeline("<Folder Name=""" & escape(strDispName) & """ Path=""" & escape(oXMLPathNodes(i).nodeTypedValue)  & """ fid=""1"&  Mid(oXMLFIDNodes(i).text,InStr(Len(oXMLFIDNodes(i).text)-8,oXMLFIDNodes(i).text,"-"),10) & """></Folder>")
      If oXMLHasSubsNodes.Item(i).nodeTypedValue = True Then
         call RecurseFolder(oXMLHREFNodes.Item(i).nodeTypedValue)
      End If
   Next
End Sub



