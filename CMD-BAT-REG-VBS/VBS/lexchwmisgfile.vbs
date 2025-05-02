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
		drive = left(rs.fields("msExchESEParamLogFilePath").value,1) & ":"
		servername = mid(rs.fields("distinguishedName"),slen,elen-slen)
		lfilepath = replace(rs.fields("msExchESEParamLogFilePath").value,"\","\\")
		lfilepath = right(lfilepath,(len(lfilepath)-2)) & "\\"
		Set objWMIService = GetObject("winmgmts:" _
 		& "{impersonationLevel=impersonate}!\\" & servername & "\root\cimv2")
		Set lfiles = objWMIService.ExecQuery _
		 ("select * from CIM_DataFile where path = """ & lfilepath & """ and extension = ""log"" and drive = """ & drive & """")
		for each lfile in lfiles
			if lcase(left(lfile.filename,3)) = lcase(rs.fields("msExchESEParamBaseName").value) then 
				lfcount = lfcount + 1
				lfsize = lfsize + lfile.filesize
				if lfcount = 1 then lfolddatenum = lfile.LastModified
				if lfolddatenum > lfile.LastModified then
					lfolddatenum = cdate(DateSerial(Left(lfile.lastmodified, 4), Mid(lfile.lastmodified, 5, 2), Mid(lfile.lastmodified, 7, 2)) & " " & timeserial(Mid(lfile.lastmodified, 9, 2),Mid(lfile.lastmodified, 11, 2),Mid(lfile.lastmodified,13, 2)))
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













