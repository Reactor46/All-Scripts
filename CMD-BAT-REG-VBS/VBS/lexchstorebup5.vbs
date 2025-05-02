servername = wscript.arguments(0)
set conn = createobject("ADODB.Connection")
set mdbobj = createobject("CDOEXM.MailboxStoreDB")
set pdbobj = createobject("CDOEXM.PublicStoreDB")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,distinguishedName,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set nRs = Com.Execute
while not nRs.EOF
	serverdn =  nRs.fields("distinguishedName")
	nRs.movenext
Wend
mbQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchPrivateMDB)(msExchOwningServer=" & serverdn & "));name,distinguishedName,msExchEDBFile;subtree"
pfQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchPublicMDB)(msExchOwningServer=" & serverdn & "));name,distinguishedName,msExchEDBFile;subtree"
Com.ActiveConnection = Conn
Com.CommandText = mbQuery
Set Rs = Com.Execute
Wscript.echo "Mailbox Stores"
Wscript.echo
While Not Rs.EOF
		mdbobj.datasource.open "LDAP://" & Rs.Fields("distinguishedName")
		sgname = mid(rs.fields("distinguishedName"),(instr(3,rs.fields("distinguishedName"),",CN=")+4),(instr(rs.fields("distinguishedName"),",CN=InformationStore,") - (instr(3,rs.fields("distinguishedName"),",CN=")+4)))
		if sgname <> psgname then
			wscript.echo "Strorage Group Name: " & sgname
			wscript.echo
		end if
		psgname = sgname
		edbarray = split(rs.fields("msExchEDBFile"),"\")
		Wscript.echo edbarray(ubound(edbarray)) & " Last Backed Up : " & mdbobj.LastFullBackupTime
		wscript.echo 
		Rs.MoveNext

Wend
Wscript.echo "Public Folder Stores"
Wscript.echo
Com.CommandText = pfQuery
Set Rs1 = Com.Execute
While Not Rs1.EOF
		sgname = mid(rs1.fields("distinguishedName"),(instr(3,rs1.fields("distinguishedName"),",CN=")+4),(instr(rs1.fields("distinguishedName"),",CN=InformationStore,") - (instr(3,rs1.fields("distinguishedName"),",CN=")+4))) 
		if sgname <> psgname then
			wscript.echo "Strorage Group Name: " & sgname
			wscript.echo
		end if
		pdbobj.datasource.open "LDAP://" & Rs1.Fields("distinguishedName")
		edbarray = split(rs1.fields("msExchEDBFile"),"\")
		Wscript.echo edbarray(ubound(edbarray)) & " Last Backed Up : " & pdbobj.LastFullBackupTime
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

