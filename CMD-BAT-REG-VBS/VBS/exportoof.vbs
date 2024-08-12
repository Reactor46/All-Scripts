servername = wscript.arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\temp\offexport-" & servername & ".xml",2,true) 
wfile.writeline("<?xml version=""1.0""?>")
wfile.writeline("<ExportedOffs ExportDate=""" & WeekdayName(weekday(now),3) & ", " & day(now()) & " " & Monthname(month(now()),3) & " " & year(now()) & " " & formatdatetime(now(),4) & ":00" & """>")
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
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,mailnickname,mail;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		call procmailboxes(servername,rs1.fields("mail"))
		wscript.echo rs1.fields("mail")
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
wfile.writeline("</ExportedOffs>")
wfile.close
set fso = nothing
set conn = nothing
set com = Nothing

wscript.echo "Done"




sub procmailboxes(servername,MailboxAlias)

Set msMapiSession = CreateObject("MAPI.Session")
on error Resume next
msMapiSession.Logon "","",False,True,True,True,Servername & vbLF & MailboxAlias
if err.number = 0 then
	on error goto 0
	if msMapiSession.outofoffice = false and msMapiSession.outofofficetext = "" then
		wscript.echo "No OOF Data for user"
	else 
		wfile.writeline("<OOFSetting DisplayName=""" & msMapiSession.CurrentUser & """ EmailAddress=""" & MailboxAlias & """ Offset=""" _
		& msMapiSession.outofoffice & """><![CDATA[ " & msMapiSession.outofofficetext & "]]></OOFSetting>")
	End if
else
	Wscript.echo "Error Opening Mailbox"
end if
Set msMapiSession = Nothing
Set mrMailboxRules = Nothing

End Sub

