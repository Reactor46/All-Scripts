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
		objmailstorename = "LDAP://" & Rs.Fields("distinguishedName") 
		set objmailstore = getObject(objmailstorename)
		sgname = mid(rs.fields("distinguishedName"),(instr(3,rs.fields("distinguishedName"),",CN=")+4),(instr(rs.fields("distinguishedName"),",CN=InformationStore,") - (instr(3,rs.fields("distinguishedName"),",CN=")+4)))
		if sgname <> psgname then
			wscript.echo "Strorage Group Name: " & sgname
			wscript.echo
		end if
	        slvlastbackuped = queryeventlog(servername,objmailstore.msExchSLVFile)
		edblastbackuped = queryeventlog(servername,objmailstore.msExchEDBFile)
		edbarray = split(objmailstore.msExchEDBFile,"\")
		Wscript.echo edbarray(ubound(edbarray)) & " Last Backed Up EDB : " & edblastbackuped 
		slvarray = split(objmailstore.msExchSLVFile,"\")
                Wscript.echo slvarray(ubound(slvarray)) & " Last Backed Up STM : " & slvlastbackuped
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
		objmailstorename = "LDAP://" & Rs1.Fields("distinguishedName") 
		set objmailstore = getObject(objmailstorename)
		servername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
	        slvlastbackuped = queryeventlog(servername,objmailstore.msExchSLVFile)
		edblastbackuped = queryeventlog(servername,objmailstore.msExchEDBFile)
		edbarray = split(objmailstore.msExchEDBFile,"\")
		Wscript.echo edbarray(ubound(edbarray)) & " Last Backed Up EDB : " & edblastbackuped 
		slvarray = split(objmailstore.msExchSLVFile,"\")
                Wscript.echo slvarray(ubound(slvarray)) & " Last Backed Up STM : " & slvlastbackuped
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


function queryeventlog(servername,filename)
SB = 0
dtmStartDate = CDate(Date) - 7
dtmStartDate = Year(dtmStartDate) & Right( "00" & Month(dtmStartDate), 2) & Right( "00" & Day(dtmStartDate), 2)
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & servername & "\root\cimv2")
Set colLoggedEvents = objWMIService.ExecQuery("Select * from Win32_NTLogEvent Where Logfile='Application' and Eventcode = '221' and TimeWritten >= '" & dtmStartDate & "' ",,48)
For Each objEvent in colLoggedEvents
    SB = 1
    Time_Written = objEvent.TimeWritten
    Time_Written = left(Time_Written,(instr(Time_written,".")-1))
    if instr(objEvent.Message,filename) then
	queryeventlog = dateserial(mid(Time_Written,1,4),mid(Time_Written,5,2),mid(Time_Written,7,2)) & " " & timeserial(mid(Time_Written,9,2),mid(Time_Written,11,2),mid(Time_Written,13,2))
	exit for
    end if 
next
if SB = 0 then queryeventlog = "No Backup recorded in the last 7 Days"
end function










