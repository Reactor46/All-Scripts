lfcount = 0
lfsize = 0
lfoldatenum = ""
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
sgQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchStorageGroup);name,distinguishedName,msExchESEParamBaseName,adminDisplayName,msExchESEParamLogFilePath,msExchEDBFile;subtree"
Com.ActiveConnection = Conn
Com.CommandText = sgQuery
Set Rs = Com.Execute
Wscript.echo "Storeage Groups"
Wscript.echo
While Not Rs.EOF
		slen = instr(rs.fields("distinguishedName"),"CN=InformationStore,") + 23
		elen = instr(rs.fields("distinguishedName"),"CN=Servers,")-1
		Set fso = CreateObject("Scripting.FileSystemObject")
		logfileunc = "\\" & mid(rs.fields("distinguishedName"),slen,elen-slen) & "\" & left(rs.fields("msExchESEParamLogFilePath").value,1) & "$" & mid(rs.fields("msExchESEParamLogFilePath").value,3,len(rs.fields("msExchESEParamLogFilePath").value)-2)
		set lfolder = fso.getfolder(logfileunc)	
		set lfiles = lfolder.files
		for each lfile in lfiles
			if left(lfile.name,3) = rs.fields("msExchESEParamBaseName").value and right(lfile.name,3) = "log" then
				lfcount = lfcount + 1
				lfsize = lfsize + lfile.size
				if lfcount = 1 then lfolddatenum = lfile.DateLastModified
				if lfolddatenum > lfile.DateLastModified then
					lfolddatenum = lfile.DateLastModified
				end if
			end if	
		next
		wscript.echo "ServerName : " & mid(rs.fields("distinguishedName"),slen,elen-slen)		
		wscript.echo "Storage Group Name : " & rs.fields("adminDisplayName")
		wscript.echo "Log file Path : " & rs.fields("msExchESEParamLogFilePath").value
		wscript.echo "Log file Prefix : " & rs.fields("msExchESEParamBaseName").value
		wscript.echo "Number of Log files in Directory : " & lfcount
		wscript.echo "Disk Space being used : " & formatnumber(lfsize/1048576,2,0,0,0) & " MB"
		wscript.echo "Oldest Log file in this Directory : " & lfolddatenum
		wscript.echo
		lfcount = 0
		lfsize = 0
		lfolddatenum = ""
		Rs.MoveNext
Wend
Rs.Close
Conn.Close
set mdbobj = Nothing
set pdbobj = Nothing
Set Rs = Nothing
Set Com = Nothing
Set Conn = Nothing













