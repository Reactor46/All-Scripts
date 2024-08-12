strFromAddress = "reports@yourdomain.com"
strToAddress = "user@yourdomain.com"
strSMTPServer = "servername"

strEmailBody = ""
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
strEmailBody = strEmailBody & "Mailbox Stores" & vbcrlf
strEmailBody = strEmailBody &  vbcrlf
While Not Rs.EOF
		objmailstorename = "LDAP://" & Rs.Fields("distinguishedName") 
		set objmailstore = getObject(objmailstorename)
		servername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
		dnarray = split(Rs.Fields("distinguishedName"),",",-1,1) 
                sgname = mid(dnarray(1),4) & "\" & mid(dnarray(0),4)
	        Dbfreespace = queryeventlog(servername,sgname,"1221")
		strEmailBody = strEmailBody &  Rs.Fields("name") & " Freespace after Defrag : " & Dbfreespace  & vbcrlf
		Dbreten = queryeventlog(servername,sgname,"1207")
		mbreten = queryeventlog(servername,sgname,"9535")
		strEmailBody = strEmailBody & vbcrlf
		Rs.MoveNext

Wend
strEmailBody = strEmailBody &  "Public Folder Stores" & vbcrlf
strEmailBody = strEmailBody &  vbcrlf
Com.CommandText = pfQuery
Set Rs1 = Com.Execute
While Not Rs1.EOF
		objmailstorename = "LDAP://" & Rs1.Fields("distinguishedName") 
		set objmailstore = getObject(objmailstorename)
		servername = mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4)
		dnarray1 = split(Rs1.Fields("distinguishedName"),",",-1,1) 
                sgname = mid(dnarray1(1),4) & "\" & mid(dnarray1(0),4)
	        Dbfreespace = queryeventlog(servername,sgname,"1221")
		strEmailBody = strEmailBody &  Rs1.Fields("name") & " Freespace after Defrag : " & Dbfreespace & vbcrlf
		Dbreten = queryeventlog(servername,sgname,"1207")		
		strEmailBody = strEmailBody & vbcrlf
		Rs1.MoveNext

Wend
Rs.Close
Rs1.close
Conn.Close
Set Rs = Nothing
Set Rs1 = Nothing
Set Com = Nothing
Set Conn = Nothing
Set objEmail = CreateObject("CDO.Message")
objEmail.From = strFromAddress
objEmail.To = strToAddress
objEmail.Subject = "Exchange Whitespace Report"
objEmail.textbody = strEmailBody 
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServer 
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Send
set objEmail = nothing


function queryeventlog(servername,sgname,event2s)
SB = 0
dtmStartDate = CDate(Date) - 7
dtmStartDate = Year(dtmStartDate) & Right( "00" & Month(dtmStartDate), 2) & Right( "00" & Day(dtmStartDate), 2)
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & servername & "\root\cimv2")
Set colLoggedEvents = objWMIService.ExecQuery("Select * from Win32_NTLogEvent Where Logfile='Application' and Eventcode = '" & event2s & "' and TimeWritten >= '" & dtmStartDate & "' ",,48)
For Each objEvent in colLoggedEvents
    SB = 1
    Time_Written = objEvent.TimeWritten
    Time_Written = left(Time_Written,(instr(Time_written,".")-1))
    if instr(objEvent.Message,sgname) then
	if event2s = "1221" then
		queryeventlog = Mid(objEvent.Message,InStr(15,objEvent.Message,chr(34))+6,(InStr(1,objEvent.Message,"megabytes")-1)-(InStr(15,objEvent.Message,chr(34))+6))
	else
		if event2s = "1207" then
			StartItems = Mid(objEvent.Message,InStr(82,objEvent.Message,chr(34))+13,(InStr(82,objEvent.Message,"items;")-(InStr(82,objEvent.Message,chr(34))+14)))
			StartSize =  Mid(objEvent.Message,(InStr(objEvent.Message,"items;")+7),InStr((InStr(objEvent.Message,"items;")+7),objEvent.Message," ")-(InStr(objEvent.Message,"items;")+7))
			End_Items =  Mid(objEvent.Message,(InStr(objEvent.Message,"End:")+5),InStr((InStr(objEvent.Message,"End:")+5),objEvent.Message," ")-(InStr(objEvent.Message,"End:")+5))
			End_Size =   Mid(objEvent.Message,(InStr((InStr(objEvent.Message,"End:")+5),objEvent.Message,"items;")+7),InStr((InStr((InStr(objEvent.Message,"End:")+5),objEvent.Message,"items;")+7),objEvent.Message," ")-(InStr((InStr(objEvent.Message,"End:")+5),objEvent.Message,"items;")+7))
			strEmailBody = strEmailBody &  "Retained StartItems : " & StartItems & "    StartSize : " & formatnumber(StartSize/1024,2) & vbcrlf
			strEmailBody = strEmailBody &  "Retained EndItems : "  & End_Items & "	EndSize : " & formatnumber(End_Size/1024,2) & vbcrlf
		else
			Deleted_Number = Mid(objEvent.Message,InStr(88,objEvent.Message,".")+5,InStr(88,objEvent.Message,"deleted")-1-(InStr(88,objEvent.Message,".")+5))
		        Deleted_Size = Mid(objEvent.Message,(InStr(88,objEvent.Message,"deleted")+19),InStr(InStr(88,objEvent.Message,"deleted")+19,objEvent.Message," ")-(InStr(88,objEvent.Message,"deleted")+19))
			Retained_Number = Mid(objEvent.Message,InStr(88,objEvent.Message,"removed.")+12,InStr(InStr(88,objEvent.Message,"removed.")+8,objEvent.Message,"deleted")-(InStr(88,objEvent.Message,"removed.")+12))
			Retained_Size =  Mid(objEvent.Message,InStr((InStr(88,objEvent.Message,"removed.")+8),objEvent.Message,"mailboxes")+11,InStr(InStr((InStr(88,objEvent.Message,"removed.")+8),objEvent.Message,"mailboxes")+11,objEvent.Message," ")-(InStr((InStr(88,objEvent.Message,"removed.")+8),objEvent.Message,"mailboxes")+11))
			strEmailBody = strEmailBody & "Number of Deleted Mailboxs Removed : " & Deleted_Number & "     Size : " & formatnumber(Deleted_Size/1024,2) & vbcrlf
			strEmailBody = strEmailBody & "Number of Deleted Mailboxs Retained : " & Retained_Number & "     Size : " & formatnumber(Retained_Size/1024,2) & vbcrlf
		end if
	end if
	exit for
    end if 
next
if SB = 0 then queryeventlog = "No Backup recorded in the last 7 Days"
end function



