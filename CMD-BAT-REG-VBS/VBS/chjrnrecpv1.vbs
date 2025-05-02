Servername = wscript.arguments(0)
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set snrs = Com.Execute
mbQuery = "<LDAP://" & strNameingContext & ">;(&(&(objectCategory=msExchPrivateMDB)(msExchOwningServer=" & snrs.fields("distinguishedName") & ")(cn=" & wscript.arguments(1) & ")));name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = mbQuery
Set Rs = Com.Execute
While Not Rs.EOF
	exStoreDN = "LDAP://" & rs.fields("distinguishedName")
	rs.movenext
Wend
if exStoreDN = "" then wscript.echo "No Store Found"
usrQuery = "<LDAP://" & strDefaultNamingContext & ">;(mailnickname=" & wscript.arguments(2) & ");name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = usrQuery
Set RsUsr = Com.Execute
While Not RsUsr.EOF
	jnJournalDN = RsUsr.fields("distinguishedName")
	RsUsr.movenext
Wend
if jnJournalDN = "" then wscript.echo "No User Found"
if jnJournalDN <> "" and exStoreDN <> "" then
	wscript.echo "Can configure"
	set exExchangeStore = createobject("CDOEXM.MailboxStoreDB")
	exExchangeStore.datasource.open exStoreDN,,3
	wscript.echo "Current Journal Recipient set to " & exExchangeStore.fields("msExchMessageJournalRecipient")
	wscript.echo
	exExchangeStore.fields("msExchMessageJournalRecipient").value = jnJournalDN
	exExchangeStore.fields.update
	exExchangeStore.datasource.save
	set exExchangeStore = nothing
	wscript.echo "New Journal Recipient set to " & jnJournalDN
Else
	wscript.echo
	wscript.echo "one of the parameters passed in was not valid the script can't continue"
	wscript.echo "use the following Syntax cscript chjrnrecpv1.vbs servername ""Mailbox Store (servername)"" mailboxalias"
end if