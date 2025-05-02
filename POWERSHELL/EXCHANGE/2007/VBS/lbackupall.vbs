On error resume next
Set fso = CreateObject("Scripting.FileSystemObject")
fname = "c:\BackupResults.csv"
set wfile = fso.opentextfile(fname,2,true)
wfile.writeline("Servername,StorageGroupName,EDB FileName,LastBackup Result")
set conn = createobject("ADODB.Connection")
set mdbobj = createobject("CDOEXM.MailboxStoreDB")
set pdbobj = createobject("CDOEXM.PublicStoreDB")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
mbQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchPrivateMDB);name,distinguishedName,msExchOwningServer,msExchEDBFile;subtree"
pfQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchPublicMDB);name,distinguishedName,msExchOwningServer,msExchEDBFile;subtree"
Com.ActiveConnection = Conn
Com.CommandText = mbQuery
Set Rs = Com.Execute
Wscript.echo "Mailbox Stores"
Wscript.echo
While Not Rs.EOF
		mdbobj.datasource.open "LDAP://" & Rs.Fields("distinguishedName")
		servername = mid(rs.fields("msExchOwningServer"),4,instr(rs.fields("msExchOwningServer"),",")-4)
		sgname = mid(rs.fields("distinguishedName"),(instr(3,rs.fields("distinguishedName"),",CN=")+4),(instr(rs.fields("distinguishedName"),",CN=InformationStore,") - (instr(3,rs.fields("distinguishedName"),",CN=")+4)))
		if sgname <> psgname then
			wscript.echo "Strorage Group Name: " & sgname
			wscript.echo
		end if
		psgname = sgname
		edbarray = split(rs.fields("msExchEDBFile"),"\")
		Wscript.echo edbarray(ubound(edbarray)) & " Last Backed Up : " & mdbobj.LastFullBackupTime
		wscript.echo 
		wfile.writeline(servername & "," & sgname & "," & edbarray(ubound(edbarray)) & "," & mdbobj.LastFullBackupTime)
		if datediff("d",mdbobj.LastFullBackupTime,now()) > 6 then call sendalert(servername,sgname,edbarray(ubound(edbarray)),mdbobj.LastFullBackupTime)
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
		servername = mid(rs1.fields("msExchOwningServer"),4,instr(rs1.fields("msExchOwningServer"),",")-4)
		edbarray = split(rs1.fields("msExchEDBFile"),"\")
		Wscript.echo edbarray(ubound(edbarray)) & " Last Backed Up : " & pdbobj.LastFullBackupTime
		wscript.echo 
		wfile.writeline(servername & "," & sgname & "," & edbarray(ubound(edbarray)) & "," & pdbobj.LastFullBackupTime)
		if datediff("d",pdbobj.LastFullBackupTime,now()) > 6 then call sendalert(servername,sgname,edbarray(ubound(edbarray)),pdbobj.LastFullBackupTime)
		Rs1.MoveNext

Wend
Rs.Close
Rs1.close
Conn.Close
wfile.close
set wfile = nothing
set mdbobj = Nothing
set pdbobj = Nothing
Set Rs = Nothing
Set Rs1 = Nothing
Set Com = Nothing
Set Conn = Nothing


sub sendalert(servername,sgname,edbname,lastbackup)
	Set objEmail = CreateObject("CDO.Message")
	objEmail.From = "backup@blah.com"
	objEmail.To = "blah@blah.com"
	objEmail.Subject = "Backup Alert no Exchange Backup in last 7 Days " & serverName & " " & edbname
	objEmail.Textbody = serverName & " " & sgname & " " & edbname & " Last Succesfull Backup " & lastbackup
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "Servername"
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objEmail.Configuration.Fields.Update
	objEmail.Send
end sub










