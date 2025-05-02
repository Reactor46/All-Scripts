set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
rangeStep = 999
lowRange = 0
highRange = lowRange + rangeStep
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
		mbnum = 0
		rangeStep = 999
		lowRange = 0
		highRange = lowRange + rangeStep
		quit = false
		set objmailstore = getObject(objmailstorename)
		Do until quit = true
			on error resume next
			strCommandText = "homeMDBBL;range=" & lowRange & "-" & highRange
			objmailstore.GetInfoEx Array(strCommandText), 0
			if err.number <> 0 then quit = true
			varReports = objmailstore.GetEx("homeMDBBL")
			if quit <> true then mbnum = mbnum + ubound(varReports)+1
		        lowRange = highRange + 1
        		highRange = lowRange + rangeStep
		loop
		err.clear
		Set efile = GetObject("winMgmts:!\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\root\cimv2:CIM_DataFile.Name='" & objmailstore.msExchEDBFile & "'")
                set sfile = GetObject("winMgmts:!\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\root\cimv2:CIM_DataFile.Name='" & objmailstore.msExchSLVFile & "'")
		edbsize =  formatnumber(efile.filesize/1073741824,2,0,0,0)
		stmsize = formatnumber(sfile.filesize/1073741824,2,0,0,0)
		set edrive = GetObject("winMgmts:!\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\root\cimv2:Win32_LogicalDisk.DeviceID='" & left(objmailstore.msExchEDBFile,1) & ":'")
		set sdrive = GetObject("winMgmts:!\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\root\cimv2:Win32_LogicalDisk.DeviceID='" & left(objmailstore.msExchSLVFile,1) & ":'")
		edbtotalspace = round(edrive.size/1073741824)
		edbfreespace = round(edrive.FreeSpace/1073741824)
		edbpercentleft = FormatNumber((edbfreespace/edbtotalspace)*100, 0)
		stmtotalspace = round(sdrive.size/1073741824)
		stmfreespace = round(sdrive.FreeSpace/1073741824)
		stmpercentleft = FormatNumber((stmfreespace/stmtotalspace)*100, 0)
		Wscript.echo Rs.Fields("name") & "# Mailboxes: " & mbnum & " EDBSize(GB): " & edbsize & " STMSize(GB): " & stmsize
		wscript.echo "Freespace on EDB Drive (" & left(objmailstore.msExchEDBFile,1) & ":) :" & edbfreespace & " GB	Percent Left :" & edbpercentleft & " %"
		wscript.echo "Freespace on STM Drive (" & left(objmailstore.msExchSLVFile,1) & ":) :" & stmfreespace & " GB	Percent Left :" & stmpercentleft & " %"
		wscript.echo 
		Rs.MoveNext

Wend
Wscript.echo "Public Folder Stores"
Wscript.echo
Com.CommandText = pfQuery
Set Rs1 = Com.Execute
While Not Rs1.EOF
		objmailstorename = "LDAP://" & Rs1.Fields("distinguishedName") 
		set objmailstore = getObject(objmailstorename)
		Set efile = GetObject("winMgmts:!\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\root\cimv2:CIM_DataFile.Name='" & objmailstore.msExchEDBFile & "'")
                set sfile = GetObject("winMgmts:!\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\root\cimv2:CIM_DataFile.Name='" & objmailstore.msExchSLVFile & "'")
		edbsize =  formatnumber(efile.filesize/1073741824,2,0,0,0)
		stmsize = formatnumber(sfile.filesize/1073741824,2,0,0,0)
		set edrive = GetObject("winMgmts:!\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\root\cimv2:Win32_LogicalDisk.DeviceID='" & left(objmailstore.msExchEDBFile,1) & ":'")
		set sdrive = GetObject("winMgmts:!\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\root\cimv2:Win32_LogicalDisk.DeviceID='" & left(objmailstore.msExchSLVFile,1) & ":'")
		edbtotalspace = round(edrive.size/1073741824)
		edbfreespace = round(edrive.FreeSpace/1073741824)
		edbpercentleft = FormatNumber((edbfreespace/edbtotalspace)*100, 0)
		stmtotalspace = round(sdrive.size/1073741824)
		stmfreespace = round(sdrive.FreeSpace/1073741824)
		stmpercentleft = FormatNumber((stmfreespace/stmtotalspace)*100, 0)
		Wscript.echo Rs1.Fields("name") & " EDBSize(GB): " & edbsize & " STMSize(GB): " & stmsize
		wscript.echo "Freespace on EDB Drive (" & left(objmailstore.msExchEDBFile,1) & ":) :" &  edbfreespace & " GB	Percent Left :" & edbpercentleft & " %"
		wscript.echo "Freespace on STM Drive (" & left(objmailstore.msExchSLVFile,1) & ":) :" &  stmfreespace & " GB	Percent Left :" & stmpercentleft & " %"
		wscript.echo 
		Rs1.MoveNext

Wend
Rs.Close
Rs1.close
Conn.Close
Set Rs = Nothing
Set Rs1 = Nothing
Set Com = Nothing
Set Conn = Nothing






