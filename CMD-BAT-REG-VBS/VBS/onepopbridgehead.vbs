set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
Set fso = CreateObject("Scripting.FileSystemObject")
servername = "MGNUS03"
Set Cnxn1 = CreateObject("ADODB.Connection")
strCnxn1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\Trackinglog.mdb;"
Cnxn1.Open strCnxn1
SQL1 = "clearbridgeheadusers"
Cnxn1.Execute(SQL1)
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
mbQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchPrivateMDB);name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = mbQuery
Set Rs = Com.Execute
While Not Rs.EOF
		objmailstorename = "LDAP://" & Rs.Fields("distinguishedName") 
		set objmailstore = getObject(objmailstorename)
		if mid(objmailstore.msExchOwningServer,4,instr(objmailstore.msExchOwningServer,",")-4) = servername then
			objmailstore.GetInfoEx Array("homeMDBBL"), 0
			varReports = objmailstore.GetEx("homeMDBBL")
			for i = lbound(varReports) to ubound(varReports)
				set objuser = getobject("LDAP://" & varReports(i))
				if instr(objuser.objectCategory,"CN=Person") then
			        	proxyaddress = objuser.getex("proxyAddresses")
					for e = lbound(proxyaddress) to ubound(proxyaddress)
						if instr(lcase(proxyaddress(e)),"smtp:") then
						emailaddress =  replace(proxyaddress(e),"'","")
						emailaddress =  replace(lcase(emailaddress),"smtp:","")
						sqlstate1 = "INSERT INTO BridgeHeadUsers ( [EmailAddress], [SamAccountName]) values ('" & emailaddress & "','" & objuser.samaccountname & "')"
						Cnxn1.Execute(sqlstate1)
						end if
					next
				end if
			next
		end if
		Rs.MoveNext

Wend
Rs.Close
Conn.Close
Set Rs = Nothing
Set Com = Nothing
Set Conn = Nothing
set f = fso.getfolder("c:\logdir")
Set fc = f.Files
For Each file1 in fc
	importtodb(file1.path)
next
wscript.echo "Done"

function importtodb(fname)
set wfile = fso.opentextfile(fname,1,true)
for i = 1 to 5
	wfile.skipline
next
nline = wfile.readline
do until wfile.AtEndOfStream
if instr(nline,"	") then
inplinearray = Split(nline, "	", -1, 1)
if inplinearray(8) = "1020" or inplinearray(8) = "1028" then
	ClientIP = inplinearray(2)
   	if (isnull(ClientIP)) then ClientIP = "N/A"
   	Entrytype = inplinearray(8)
    	if (isnull(Entrytype)) then Entrytype = "N/A"
    	Subject = inplinearray(18)
   	if (isnull(Subject)) then Subject = "N/A"
   	RecipientAddress1 =  inplinearray(7)
    	if (isnull(RecipientAddress1)) then RecipientAddress1 = "N/A"
    	RecipientCount = inplinearray(13)
    	if (isnull(RecipientCount)) then RecipientCount = "N/A"
    	SenderAddress = inplinearray(19)
    	if (isnull(SenderAddress)) then SenderAddress = "N/A"
    	size1 = inplinearray(12)
    	if (isnull(size1)) then size1 = "N/A"
	datearray = split(inplinearray(0),"-",-1,1)
	timearray = split(inplinearray(1),":",-1,1)
	odate = dateserial(datearray(0),datearray(1),datearray(2))
	otime = timeserial(timearray(0),timearray(1),0)
	odate = dateadd("h",toffset,cdate(odate & " " & otime))
    	wtowrite = "('" & condate(odate) & "','" & formatdatetime(odate,4) & "','"  & ClientIP & "','"  & EntryType & "','" & RecipientCount & "','" & replace(SenderAddress,"'","") & "','" & replace(RecipientAddress1,"'","") & "','" & left(replace(subject,"'"," "),254) & "','" & size1 & "')"
	sqlstate1 = "INSERT INTO BridgeHeadTempImport ( [Date], [Time], [client-ip], [Event-ID], NoRecipients, [Sender-Address], [Recipient-Address], [Message-Subject], [total-bytes] ) values " & wtowrite
	Cnxn1.Execute(sqlstate1)
end if
end if
nline = wfile.readline
loop
SQL1 = "ImportbridgeheadRecieved"
Cnxn1.Execute(SQL1)
SQL1 = "importbridgeheadSent"
Cnxn1.Execute(SQL1)
SQL1 = "DeleteBridgeImport"
Cnxn1.Execute(SQL1)
wfile.close
set wfile = nothing
end function


function condate(date2con)
dtcon = date2con
if month(dtcon) < 10 then 
	if day(dtcon) < 10 then
		qdat = year(dtcon) & "0" & month(dtcon) & "0" & day(dtcon)
	else
		qdat = year(dtcon) & "0" & month(dtcon) & day(dtcon)
	end if 
else
	if day(dtcon) < 10 then
		qdat = year(dtcon) & month(dtcon) & "0" & day(dtcon)
	else
		qdat = year(dtcon) & month(dtcon) & day(dtcon)
	end if 
end if
condate = qdat 
end function 
