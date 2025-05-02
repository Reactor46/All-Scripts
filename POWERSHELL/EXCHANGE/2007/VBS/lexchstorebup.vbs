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
		set objmailstore = getObject(objmailstorename)
		servername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
	        slvlastbackuped = queryeventlog(servername,objmailstore.msExchSLVFile)
		edblastbackuped = queryeventlog(servername,objmailstore.msExchEDBFile)
		Wscript.echo Rs.Fields("name") & " Last Backed up EDB : " & edblastbackuped 
		Wscript.echo Rs.Fields("name") & " Last Backed up STM : " & slvlastbackuped
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
		servername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
	        slvlastbackuped = queryeventlog(servername,objmailstore.msExchSLVFile)
		edblastbackuped = queryeventlog(servername,objmailstore.msExchEDBFile)
		Wscript.echo Rs1.Fields("name") & " Last Backed up EDB : " & edblastbackuped 
		Wscript.echo Rs1.Fields("name") & " Last Backed up STM : " & slvlastbackuped
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




