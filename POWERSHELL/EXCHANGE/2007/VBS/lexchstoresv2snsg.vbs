Servername = wscript.arguments(0)
sgcount = 0
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
rangeStep = 999
lowRange = 0
highRange = lowRange + rangeStep
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set snrs = Com.Execute
mbQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchPrivateMDB)(msExchOwningServer=" & snrs.fields("distinguishedName") & "));name,distinguishedName;subtree"
pfQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchPublicMDB)(msExchOwningServer=" & snrs.fields("distinguishedName") & "));name,distinguishedName;subtree"
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
		Set fso = CreateObject("Scripting.FileSystemObject")
		edbfilespec = "\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\" & left(objmailstore.msExchEDBFile,1) & "$" & mid(objmailstore.msExchEDBFile,3,len(objmailstore.msExchEDBFile)-2)
  		stmfilespec = "\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\" & left(objmailstore.msExchSLVFile,1) & "$" & mid(objmailstore.msExchSLVFile,3,len(objmailstore.msExchSLVFile)-2)
		Set efile = fso.GetFile(edbfilespec)
		set sfile = fso.GetFile(stmfilespec)
		edbsize =  formatnumber(efile.size/1073741824,2,0,0,0)
		stmsize = formatnumber(sfile.size/1073741824,2,0,0,0)
		set edrive = fso.GetDrive("\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\" & left(objmailstore.msExchEDBFile,1) & "$")
		set sdrive = fso.GetDrive("\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\" & left(objmailstore.msExchSLVFile,1) & "$")
		edbtotalspace = round(edrive.totalsize/1073741824)
		edbfreespace = round(edrive.FreeSpace/1073741824)
		edbpercentleft = FormatNumber((edbfreespace/edbtotalspace)*100, 0)
		stmtotalspace = round(sdrive.totalsize/1073741824)
		stmfreespace = round(sdrive.FreeSpace/1073741824)
		stmpercentleft = FormatNumber((stmfreespace/stmtotalspace)*100, 0)
		sgname = mid(rs.fields("distinguishedName"),(instr(3,rs.fields("distinguishedName"),",CN=")+4),(instr(rs.fields("distinguishedName"),",CN=InformationStore,") - (instr(3,rs.fields("distinguishedName"),",CN=")+4)))
		if sgname <> psgname then
			if sgcount <> 0 then
				Wscript.echo "Storage Groups Totals"
				Wscript.echo "Total Number of Mailboxes :" & sgmailsum
				Wscript.echo "Total Size of MailStores within Storeage Group: :" & formatnumber(sgmailsizesum/1073741824,2,0,0,0)
				wscript.echo 
				sgmailsum = 0
				sgmailsizesum = 0
			end if
			wscript.echo "Strorage Group Name: " & sgname
			wscript.echo
			sgmailsum = sgmailsum + mbnum
			sgmailsizesum = efile.size + sfile.size + sgmailsizesum 
			sgcount = 1
		else
			sgmailsum = sgmailsum + mbnum
			sgmailsizesum = efile.size + sfile.size + sgmailsizesum
		end if
		psgname = sgname
		Wscript.echo Rs.Fields("name") & "# Mailboxes: " & mbnum & " EDBSize(GB): " & edbsize & " STMSize(GB): " & stmsize
		wscript.echo "Freespace on EDB Drive (" & left(objmailstore.msExchEDBFile,1) & ":) :" & edbfreespace & " GB	Percent Left :" & edbpercentleft & " %"
		wscript.echo "Freespace on STM Drive (" & left(objmailstore.msExchSLVFile,1) & ":) :" & stmfreespace & " GB	Percent Left :" & stmpercentleft & " %"
		wscript.echo 
		Rs.MoveNext
Wend
if sgcount <> 0 then
	Wscript.echo "Storage Groups Totals"
	Wscript.echo "Total Number of Mailboxes :" & sgmailsum
	Wscript.echo "Total Size of MailStores within Storeage Group: :" & formatnumber(sgmailsizesum/1073741824,2,0,0,0)
	wscript.echo 
	sgmailsum = 0
	sgmailsizesum = 0
end if
sgcount = 0
psgname = ""
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
		set edrive = fso.GetDrive("\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\" & left(objmailstore.msExchEDBFile,1) & "$")
		set sdrive = fso.GetDrive("\\" & mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) & "\" & left(objmailstore.msExchSLVFile,1) & "$")
		edbtotalspace = round(edrive.totalsize/1073741824)
		edbfreespace = round(edrive.FreeSpace/1073741824)
		edbpercentleft = FormatNumber((edbfreespace/edbtotalspace)*100, 0)
		stmtotalspace = round(sdrive.totalsize/1073741824)
		stmfreespace = round(sdrive.FreeSpace/1073741824)
		stmpercentleft = FormatNumber((stmfreespace/stmtotalspace)*100, 0)
		sgname = mid(rs1.fields("distinguishedName"),(instr(3,rs1.fields("distinguishedName"),",CN=")+4),(instr(rs1.fields("distinguishedName"),",CN=InformationStore,") - (instr(3,rs1.fields("distinguishedName"),",CN=")+4)))
		if sgname <> psgname then
			if sgcount <> 0 then
				Wscript.echo "Storage Groups Totals"
				Wscript.echo "Total Size of MailStores within Storeage Group: :" & formatnumber(sgmailsizesum/1073741824,2,0,0,0)
				wscript.echo 
				sgmailsum = 0
				sgmailsizesum = 0
			end if
			wscript.echo "Strorage Group Name: " & sgname
			wscript.echo
			sgmailsizesum = efile.size + sfile.size + sgmailsizesum 
			sgcount = 1
		else
			sgmailsizesum = efile.size + sfile.size + sgmailsizesum
		end if
		psgname = sgname
		Wscript.echo Rs1.Fields("name") & " EDBSize(GB): " & edbsize & " STMSize(GB): " & stmsize
		wscript.echo "Freespace on EDB Drive (" & left(objmailstore.msExchEDBFile,1) & ":) :" &  edbfreespace & " GB	Percent Left :" & edbpercentleft & " %"
		wscript.echo "Freespace on STM Drive (" & left(objmailstore.msExchSLVFile,1) & ":) :" &  stmfreespace & " GB	Percent Left :" & stmpercentleft & " %"
		wscript.echo 
		Rs1.MoveNext

Wend
if sgcount <> 0 then
	Wscript.echo "Storage Groups Totals"
	Wscript.echo "Total Size of MailStores within Storeage Group: :" & formatnumber(sgmailsizesum/1073741824,2,0,0,0)
	wscript.echo 
	sgmailsum = 0
	sgmailsizesum = 0
end if
Rs.Close
Rs1.close
Conn.Close
Set Rs = Nothing
Set Rs1 = Nothing
Set Com = Nothing
Set Conn = Nothing






