set conn = createobject("ADODB.Connection")
set mdbobj = createobject("CDOEXM.MailboxStoreDB")
set pdbobj = createobject("CDOEXM.PublicStoreDB")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
mbQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchPrivateMDB);name,distinguishedName;subtree"
pfQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchPublicMDB);name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = mbQuery
Set Rs = Com.Execute
Wscript.echo "Mailbox Stores"
Wscript.echo
While Not Rs.EOF
		mdbobj.datasource.open "LDAP://" & Rs.Fields("distinguishedName")
		Wscript.echo Rs.Fields("name") & " Last Backed Up : " & mdbobj.LastFullBackupTime
		wscript.echo 
		Rs.MoveNext

Wend
Wscript.echo "Public Folder Stores"
Wscript.echo
Com.CommandText = pfQuery
Set Rs1 = Com.Execute
While Not Rs1.EOF
		pdbobj.datasource.open "LDAP://" & Rs1.Fields("distinguishedName")
		Wscript.echo Rs1.Fields("name") & " Last Backed Up : " & pdbobj.LastFullBackupTime
		wscript.echo 
		Rs1.MoveNext

Wend
Rs.Close
Rs1.close
Conn.Close
set mdbobj = Nothing
set pdbobj = Nothing
Set Rs = Nothing
Set Rs1 = Nothing
Set Com = Nothing
Set Conn = Nothing













