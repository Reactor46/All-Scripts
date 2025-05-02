Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\exchstoreJrn.csv",2,true)
wfile.writeline("""Servername"",""StoreName"",""JournalingEnabled"",""BCCJournalingEnabled"",""JournalMailoxDN"",""JournalMailboxDisplayName"",""JournalEmail""")
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
		strservername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
		Wscript.echo strservername & " " & Rs.Fields("name") 
		bccjrn = getbcc(strservername)
		if objmailstore.msExchMessageJournalRecipient <> "" then
			wscript.echo "Journaling enabled"
			wscript.echo "Journal Mailbox DN : " & objmailstore.msExchMessageJournalRecipient
			set objuser = getobject("LDAP://" & objmailstore.msExchMessageJournalRecipient)
			wscript.echo "Journal Mailbox Display Name : " & objuser.displayname
			wscript.echo "Journal Mailbox Email address : " & objuser.mail
			wscript.echo "BCC Journaling Enabled : " & bccjrn
			wscript.echo "Boo"
			wfile.writeline("""" & strservername & """,""" & rs.fields("name") & """,""Yes"",""" & bccjrn & """,""" & objmailstore.msExchMessageJournalRecipient &  """,""" & objuser.displayname & """,""" & objuser.mail & """")
			rem wfile.writeline("""" & strservername & """,""" & rs.fields("name") & """,""Yes"",""" & bccjrn & """,""" & objmailstore.msExchMessageJournalRecipient  & """,""" & objuser.displayname & """,""" & objuser.email & """")		
		else
			wscript.echo "Journaling Not enabled"
			wfile.writeline("""" & strservername & """,""" & rs.fields("name") & """,""No""")
		end if
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
		strservername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
		Wscript.echo strservername & " " & Rs1.Fields("name") 
		bccjrn = getbcc(strservername)
		if objmailstore.msExchMessageJournalRecipient <> "" then
			wscript.echo "Journaling enabled"
			wscript.echo "Journal Mailbox DN : " & objmailstore.msExchMessageJournalRecipient
			set objuser = getobject("LDAP://" & objmailstore.msExchMessageJournalRecipient)
			wscript.echo "Journal Mailbox Display Name : " & objuser.displayname
			wscript.echo "Journal Mailbox Email address : " & objuser.mail
			wscript.echo "BCC Journaling Enabled : " & bccjrn
		else
			wscript.echo "Journaling Not enabled"
		end if
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


function getbcc(strComputer)
const HKEY_LOCAL_MACHINE = &H80000002
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _ 
    & strComputer & "\root\default:StdRegProv")

strKeyPath = "SYSTEM\CurrentControlSet\Services\MSExchangeTransport\Parameters\"
objReg.getdwordvalue HKEY_LOCAL_MACHINE, strKeyPath, "JournalBCC", value
if value = "" then
	wscript.echo "Journalling Not Enabled"
else
	if value = 1 then
		getbcc = "Yes"
	else
		getbcc = "No"
	end if
end if

end function



