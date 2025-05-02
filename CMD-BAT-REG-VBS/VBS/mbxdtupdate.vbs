servername = wscript.arguments(0)
set objhttp = createobject("Microsoft.XMLHTTP")
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof	
	GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(msExchHomeServerName=" & rs.fields("legacyExchangeDN") & ")) )))))"
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,mail,displayname,mailnickname;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		wscript.echo "User: " & rs1.fields("displayname")
		call procmailboxes(servername,rs1.fields("mail"))
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
set conn = nothing
set com = nothing
wscript.echo "Done"


Sub procmailboxes(servername,mailboxname)
xmlreqtxt = "<?xml version='1.0'?><a:propfind xmlns:a='DAV:' xmlns:t='http://schemas.microsoft.com/exchange/'><a:prop><t:timezone/></a:prop></a:propfind>"
objhttp.open "PROPFIND", "http://" & servername & "/exchange/" & mailboxname, false, "", ""
objhttp.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
objhttp.setRequestHeader "Depth", "0"
objhttp.setRequestHeader "Translate", "f"
objhttp.send xmlreqtxt
wscript.echo "http://" & servername & "/exchange/" & mailboxname
If objhttp.status >= 500 Then
	wscript.echo objhttp.responsetext
ElseIf objhttp.status = 207 Then
   set oResponseDoc = objhttp.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("d:timezone")
   For j = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	ctimezone = oNode.text
	wscript.echo "Current TimeZone: " & ctimezone 

   next
end if
select case ctimezone 
	case "AUS Eastern Standard Time" uptimezone = "AUS Eastern Standard Time (Commonwealth Games 2006)"
	case "Tasmania Standard Time" uptimezone = "Tasmania Standard Time (Commonwealth Games 2006)"
	case "Cen. Australia Standard Time" uptimezone = "Cen. Australia Standard Time (CommonwealthGames 2006)"
	case else uptimezone = ""
end select
if uptimezone <> "" then 

' Open the request object, assigning it the method PROPPATCH.
	propupxml = "<?xml version=""1.0""?><g:propertyupdate xmlns:g='DAV:' xmlns:t='http://schemas.microsoft.com/exchange/'>" _
	& "<g:set><g:prop><t:timezone>" & uptimezone & "</t:timezone></g:prop></g:set></g:propertyupdate>"
	objhttp.open "PROPPATCH", "http://" & servername & "/exchange/" & mailboxname, false, "", ""
	objhttp.setRequestHeader "Content-Type", "text/xml; charset=""UTF-8"""
	objhttp.setRequestHeader "Translate", "f"
	objhttp.setRequestHeader "Content-Length:", Len(propupxml)
	objhttp.send propupxml
	if objhttp.status = 207 then 
		wscript.echo "Updated Timezone to :" & uptimezone
	else
		wscript.echo "Error Updating Timezone " & objhttp.status
	end if
else
	wscript.echo "No update needed"
end if
end sub



