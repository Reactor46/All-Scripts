servername = wscript.arguments(0)
PR_NTSDModificationTime = &H3FD60040
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\mbCreationTime.csv",2,true)
wfile.writeline("Mailbox,CreationTime")
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
wfile.close
set fso = nothing
set conn = nothing
set com = nothing
wscript.echo "Done"




sub procmailboxes(servername,MailboxAlias)

Set msMapiSession = CreateObject("MAPI.Session")
on error Resume next
msMapiSession.Logon "","",False,True,True,True,Servername & vbLF & MailboxAlias
if err.number = 0 then
	on error goto 0
	set mbSentItems = msMapiSession.GetDefaultFolder(3)	
	Wscript.echo mbSentItems.fields.item(PR_NTSDModificationTime)
	wfile.writeline(mailboxAlias & "," & mbSentItems.fields.item(PR_NTSDModificationTime))
else
	wscript.echo =  "Error Opening Mailbox"
	wfile.writeline(mailboxAlias & "," & "Error Opening Mailbox")
end if
Set msMapiSession = Nothing
Set mrMailboxRules = Nothing

End Sub

