set conn = createobject("ADODB.Connection")
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
		objmailstorename = "LDAP://" & Rs.Fields("distinguishedName") 
		set objmailstore = getObject(objmailstorename)
		objmailstore.GetInfoEx Array("homeMDBBL"), 0
		varReports = objmailstore.GetEx("homeMDBBL")
		Set fso = CreateObject("Scripting.FileSystemObject")
		edbfilespec = "\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\" & left(objmailstore.msExchEDBFile,1) & "$" & mid(objmailstore.msExchEDBFile,3,len(objmailstore.msExchEDBFile)-2)
  		stmfilespec = "\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\" & left(objmailstore.msExchSLVFile,1) & "$" & mid(objmailstore.msExchSLVFile,3,len(objmailstore.msExchSLVFile)-2)
		Set efile = fso.GetFile(edbfilespec)
		set sfile = fso.GetFile(stmfilespec)
		edbsize =  formatnumber(efile.size/1073741824,2,0,0,0)
		stmsize = formatnumber(sfile.size/1073741824,2,0,0,0)
		Wscript.echo Rs.Fields("name") & "# Mailboxes: " & ubound(varReports)+1 & " EDBSize(GB): " & edbsize & " STMSize(GB): " & stmsize
		Rs.MoveNext

Wend
Wscript.echo
Wscript.echo "Public Folder Stores"
Wscript.echo
Com.CommandText = pfQuery
Set Rs1 = Com.Execute
While Not Rs1.EOF
		objmailstorename = "LDAP://" & Rs1.Fields("distinguishedName") 
		set objmailstore = getObject(objmailstorename)
		Set fso = CreateObject("Scripting.FileSystemObject")
		edbfilespec = "\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\" & left(objmailstore.msExchEDBFile,1) & "$" & mid(objmailstore.msExchEDBFile,3,len(objmailstore.msExchEDBFile)-2)
  		stmfilespec = "\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\" & left(objmailstore.msExchSLVFile,1) & "$" & mid(objmailstore.msExchSLVFile,3,len(objmailstore.msExchSLVFile)-2)
		Set efile = fso.GetFile(edbfilespec)
		set sfile = fso.GetFile(stmfilespec)
		edbsize =  formatnumber(efile.size/1073741824,2,0,0,0)
		stmsize = formatnumber(sfile.size/1073741824,2,0,0,0)
		Wscript.echo Rs1.Fields("name") & " EDBSize(GB): " & edbsize & " STMSize(GB): " & stmsize
		Rs1.MoveNext

Wend
Rs.Close
Rs1.close
Conn.Close
Set Rs = Nothing
Set Rs1 = Nothing
Set Com = Nothing
Set Conn = Nothing





