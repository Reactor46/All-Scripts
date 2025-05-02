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
wscript.echo 
While Not Rs.EOF
		objmailstorename = "LDAP://" & Rs.Fields("distinguishedName") 
		set objmailstore = getObject(objmailstorename)
		servername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
		wscript.echo Rs.Fields("name") 
		wscript.echo 
	        getbackupdata = queryeventlog(servername,objmailstore.msExchEDBFile,objmailstore.msExchSLVFile)
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
		wscript.echo Rs1.Fields("name") 
		wscript.echo 
	        getbackupdata = queryeventlog(servername,objmailstore.msExchEDBFile,objmailstore.msExchSLVFile)
		Rs1.MoveNext

Wend
Rs.Close
Rs1.close
Conn.Close
Set Rs = Nothing
Set Rs1 = Nothing
Set Com = Nothing
Set Conn = Nothing


function queryeventlog(servername,edbfilename,stmfilename)
days = 60
SB = 0
edbarraycnt = 0
stmarraycnt = 0
dtmStartDate = CDate(Date) - days
dtmStartDate = Year(dtmStartDate) & Right( "00" & Month(dtmStartDate), 2) & Right( "00" & Day(dtmStartDate), 2)
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & servername & "\root\cimv2")
Set colLoggedEvents = objWMIService.ExecQuery("Select * from Win32_NTLogEvent Where Logfile='Application' and Eventcode = '220' and TimeWritten >= '" & dtmStartDate & "'",,48)
For Each objEvent in colLoggedEvents
    SB = 1
    Time_Written = objEvent.TimeWritten
    Time_Written = left(Time_Written,(instr(Time_written,".")-1))
    if instr(objEvent.Message,edbfilename) then
	redim Preserve edbarray(3,edbarraycnt)
	if edbarraycnt = 0 then
		edbarray(0,edbarraycnt) = 0
		edbarray(1,edbarraycnt) = dateserial(mid(Time_Written,1,4),mid(Time_Written,5,2),mid(Time_Written,7,2))
		edbarray(2,edbarraycnt) = Mid(objEvent.Message,(InStr(objEvent.Message,"(size")+6),(InStr(objEvent.Message,"Mb)")-1)-(InStr(objEvent.Message,"(size")+6))	
	else
		edbarray(0,(edbarraycnt-1)) = edbarray(2,(edbarraycnt-1)) - Mid(objEvent.Message,(InStr(objEvent.Message,"(size")+6),(InStr(objEvent.Message,"Mb)")-1)-(InStr(objEvent.Message,"(size")+6))	
		edbarray(1,edbarraycnt) = dateserial(mid(Time_Written,1,4),mid(Time_Written,5,2),mid(Time_Written,7,2))
		edbarray(2,edbarraycnt) = Mid(objEvent.Message,(InStr(objEvent.Message,"(size")+6),(InStr(objEvent.Message,"Mb)")-1)-(InStr(objEvent.Message,"(size")+6))	
		edbarraysum = edbarray(0,(edbarraycnt-1)) + edbarraysum
	end if
	queryeventlog = dateserial(mid(Time_Written,1,4),mid(Time_Written,5,2),mid(Time_Written,7,2)) 
    	edbarraycnt = edbarraycnt + 1
    else 
	if instr(objEvent.Message,stmfilename) then
	    redim Preserve stmarray(3,stmarraycnt)
	    if stmarraycnt = 0 then
		stmarray(0,stmarraycnt) = 0
		stmarray(1,stmarraycnt) = dateserial(mid(Time_Written,1,4),mid(Time_Written,5,2),mid(Time_Written,7,2))
		stmarray(2,stmarraycnt) = Mid(objEvent.Message,(InStr(objEvent.Message,"(size")+6),(InStr(objEvent.Message,"Mb)")-1)-(InStr(objEvent.Message,"(size")+6))	
	    else
		stmarray(0,(stmarraycnt-1)) = stmarray(2,(stmarraycnt-1)) - Mid(objEvent.Message,(InStr(objEvent.Message,"(size")+6),(InStr(objEvent.Message,"Mb)")-1)-(InStr(objEvent.Message,"(size")+6))	
		stmarray(1,stmarraycnt) = dateserial(mid(Time_Written,1,4),mid(Time_Written,5,2),mid(Time_Written,7,2))
		stmarray(2,stmarraycnt) = Mid(objEvent.Message,(InStr(objEvent.Message,"(size")+6),(InStr(objEvent.Message,"Mb)")-1)-(InStr(objEvent.Message,"(size")+6))	
	    	stmarraysum = stmarray(0,(stmarraycnt-1)) + stmarraysum
	    end if
	    stmarraycnt = stmarraycnt + 1
	end if
    end if 
next
if SB = 0 then queryeventlog = "No Backup recorded in the last 7 Days"
if edbarraycnt > 1 then
	Wscript.echo "EDB Mailstore file Size on " & edbarray(1,(edbarraycnt-1)) & " Size :" & edbarray(2,(edbarraycnt-1)) &" MB"
	Wscript.echo "EDB Mailstore file Size on " & edbarray(1,0) & " Size :" & edbarray(2,0) &" MB"
        Wscript.echo "EDB Mailstore file size Growth over " & datediff("d",edbarray(1,(edbarraycnt-1)),edbarray(1,0)) & " Days " & edbarray(2,0) - edbarray(2,(edbarraycnt-1))  & " MB"
	wscript.echo 
	Wscript.echo "STM Mailstore file Size on " & stmarray(1,(stmarraycnt-1)) & " Size :" & stmarray(2,(stmarraycnt-1)) &" MB"
	Wscript.echo "STM Mailstore file Size on " & stmarray(1,0) & " Size :" & stmarray(2,0) &" MB"
 	Wscript.echo "STM Mailstore File Size Growth over " & datediff("d",stmarray(1,(stmarraycnt-1)),stmarray(1,0)) & " Days " & stmarray(2,0) - stmarray(2,(stmarraycnt-1))  & " MB"
	wscript.echo
	Wscript.echo "Weekly Growth Trends EDB File"
	wscript.echo
        sdate = edbarray(1,0)
	ssize = edbarray(2,0)
	for i = 0 to (edbarraycnt-1)
		if datediff("d",edbarray(1,i),sdate) > 6 then 
			wscript.echo "Growth on " & edbarray(1,i) & " to " & sdate & ":	 " & ssize-edbarray(2,i) & " MB"
			sdate = edbarray(1,i)
			ssize = edbarray(2,i)
		end if		
	next
	wscript.echo  "Daily Average : " & formatnumber(edbarraysum/(datediff("d",edbarray(1,(edbarraycnt-1)),edbarray(1,0))),2) & " MB"
	wscript.echo   
	Wscript.echo "Weekly Growth Trends STM File"
	wscript.echo
        sdate = stmarray(1,0)
	ssize = stmarray(2,0)
	for i = 0 to (stmarraycnt-1)
		if datediff("d",stmarray(1,i),sdate) > 6 then 
			wscript.echo "Growth on " & stmarray(1,i) & " to " & sdate & ":	 " & ssize-stmarray(2,i) & " MB"
			sdate = stmarray(1,i)
			ssize = stmarray(2,i)
		end if		
	next
	wscript.echo  "Daily Average : " & formatnumber(stmarraysum/(datediff("d",stmarray(1,(stmarraycnt-1)),stmarray(1,0))),2) & " MB"
	wscript.echo   
end if
end function




