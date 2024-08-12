servername = wscript.arguments(0)
set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
set conn1 = createobject("ADODB.Connection")
strConnString = "Data Provider=NONE; Provider=MSDataShape"
conn1.Open strConnString		
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\deletedmbrep.csv",2,true)
wfile.writeline("Mailbox,MailStore,Mailbox Size(MB),Date Delete Noticed,Date when mailbox will be purged,Days left to Deletion")
set objParentRS = createobject("adodb.recordset")
set objChildRS = createobject("adodb.recordset")
strSQL = "SHAPE APPEND" & _
           "  NEW adVarChar(255) AS SOADDisplayName, " & _
           "  NEW adVarChar(255) AS SOADDistName, " & _
           "  NEW adVarChar(255) AS SOADmsExchMailboxRetentionPeriod, " & _
           " ((SHAPE APPEND  " & _
	   "      NEW adVarChar(255) AS WMILegacyDN, " & _
	   "      NEW adVarChar(255) AS WMIMailboxDisplayName, " & _
	   "      NEW adVarChar(255) AS WMISize, " & _
	   "      NEW adVarChar(255) AS WMIDateDiscoveredAbsentInDS, " & _
	   "      NEW adVarChar(255) AS WMIStorename) " & _
	   "   RELATE SOADDisplayName TO WMIStorename) AS MOWMI"
objParentRS.LockType = 3
objParentRS.Open strSQL, conn1
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,distinguishedName,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof	
	sgQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchPrivateMDB)(msExchOwningServer=" & rs.fields("distinguishedName") & "));cn,name,displayname,msExchMailboxRetentionPeriod,distinguishedName,legacyExchangeDN;subtree"
	Com.CommandText = sgQuery
	Set Rs1 = Com.Execute
	while not rs1.eof
		objParentRS.addnew 
		objParentRS("SOADDisplayName") = rs1.fields("displayname")
		objParentRS("SOADDistName") = left(rs1.fields("distinguishedName"),255)
		objParentRS("SOADmsExchMailboxRetentionPeriod") = rs1.fields("msExchMailboxRetentionPeriod")
		objParentRS.update	
		rs1.movenext
	wend
	wscript.echo "finished 1st AD query Mailbox Stores"
	rs.movenext
wend
Set objchild = objParentRS("MOWMI").Value
strWinMgmts ="winmgmts:{impersonationLevel=impersonate}!//"& servername &"/root/MicrosoftExchangeV2"
Set objWMIExchange = GetObject(strWinMgmts)
Set listExchange_MailboxSizes = objWMIExchange.ExecQuery("Select * FROM Exchange_Mailbox Where DateDiscoveredAbsentInDS IS NOT Null",,48)
For each objExchange_Mailboxs in listExchange_MailboxSizes
	objchild.addnew 
	objchild("WMILegacyDN") = objExchange_Mailboxs.LegacyDN
	objchild("WMIMailboxDisplayName") = objExchange_Mailboxs.MailboxDisplayName
	objchild("WMISize") = objExchange_Mailboxs.Size
	objchild("WMIDateDiscoveredAbsentInDS") = objExchange_Mailboxs.DateDiscoveredAbsentInDS
	objchild("WMIStorename") = objExchange_Mailboxs.storename
	objchild.update
Next
wscript.echo "finished Exchange WMI query"
wscript.echo 
objParentRS.MoveFirst
Do While Not objParentRS.EOF
    Set objChildRS = objParentRS("MOWMI").Value 
    MSdisplayname = objParentRS("SOADDisplayName")
    wscript.echo   
    wscript.echo "Mailbox Store : "  &  MSdisplayname
    if objParentRS("SOADmsExchMailboxRetentionPeriod") <> 0 then
	retrate = objParentRS("SOADmsExchMailboxRetentionPeriod")\24\60\60
    else
	retrate = 0 
    end if
    wscript.echo "Current Retention Rate : " & retrate & " days"
    Wscript.echo "Number of Deleted Mailboxes not yet purged : " & objChildRS.recordcount
    if  objChildRS.recordcount <> 0 then 
	wscript.echo "Disconnect Mailboxes"
	wscript.echo
    end if
    Do While Not objChildRS.EOF
	mbsize = objChildRS("WMISize")
	wscript.echo "Mailbox : " & objChildRS.fields("WMIMailboxDisplayName")
	wscript.echo "Size of Mailbox : " & formatnumber(mbsize/1024,2) & " MB"
	deldate = dateadd("h",toffset,cdate(DateSerial(Left(objChildRS.fields("WMIDateDiscoveredAbsentInDS"), 4), Mid(objChildRS.fields("WMIDateDiscoveredAbsentInDS"), 5, 2), Mid(objChildRS.fields("WMIDateDiscoveredAbsentInDS"), 7, 2)) & " " & timeserial(Mid(objChildRS.fields("WMIDateDiscoveredAbsentInDS"), 9, 2),Mid(objChildRS.fields("WMIDateDiscoveredAbsentInDS"), 11, 2),Mid(objChildRS.fields("WMIDateDiscoveredAbsentInDS"),13, 2))))
	wscript.echo "Date Deletetion was Detected : " & deldate
	wscript.echo "Date when Mailbox will be purged : " & dateadd("d",retrate,deldate)
	wscript.echo "Number of Days to Mailbox purge : " & datediff("d",deldate,dateadd("d",retrate,deldate))
	wscript.echo 
	wfile.writeline(replace(objChildRS.fields("WMIMailboxDisplayName"),",","") & "," & replace(MSdisplayname,",","") & "," & replace(formatnumber(mbsize/1024,2),",","")  & "," & deldate  & "," & dateadd("d",retrate,deldate) & "," & datediff("d",deldate,dateadd("d",retrate,deldate)) )
	objChildRS.MoveNext
    Loop
    objParentRS.MoveNext
Loop
wfile.close
Wscript.echo
Wscript.echo "CSV file created"