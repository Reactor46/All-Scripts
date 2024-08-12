on error resume next
set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())

Servername = wscript.arguments(0)
treport = "<table border=""1"" width=""100%"">" & vbcrlf
treport = treport & "  <tr>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Mailbox Name</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Sent Items SubFolder</font></b></td>" & vbcrlf
treport = treport & "<td align=""center"" bgcolor=""#000080""><b><font color=""#FFFFFF"">Deleted Items SubFolder</font></b></td>" & vbcrlf
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
		falias = "http://" & servername & "/exadmin/admin/" & dpDefaultpolicy & "/mbx/"
		for each paddress in rs1.fields("proxyaddresses").value
			if instr(paddress,"SMTP:") then falias = falias & replace(paddress,"SMTP:","")  
		next
		ReDim tresarray(1,6)
		wscript.echo  falias 
		defold = ""
		sefold = ""
		defold = ProcFolder(falias & "/Deleted Items/")
		sefold =  ProcFolder(falias & "/Sent Items/")
		treport = treport & "<tr>" & vbcrlf
		treport = treport & "<td align=""center"">" & rs1.fields("mail").value & "&nbsp;</td>" & vbcrlf
		treport = treport & "<td align=""center"">" & sefold  & "&nbsp;</td>" & vbcrlf
		treport = treport & "<td align=""center"">" & defold & "&nbsp;</td>" & vbcrlf
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
set wfile = fso.opentextfile("c:\temp\Server-" & Servername & "-detectedSubs.htm",2,true) 
wfile.write treport
wfile.close
set wfile = nothing
set fso = nothing

function procfolder(sUrl)
wscript.echo sUrl
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
   if oXMLSizeNodes.length = 0 then
      procfolder = False
   else
      procfolder = True
   end if

end function



