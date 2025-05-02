servername = "servername"
sg1dn = "CN=First Storage Group,CN=InformationStore,........"
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
set conn1 = createobject("ADODB.Connection")
strConnString = "Data Provider=NONE; Provider=MSDataShape"
conn1.Open strConnString		
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\mbove50report.csv",2,true)
wfile.writeline("User,Mailbox Size(MB),StorageGroupName,MailStore")
set objParentRS = createobject("adodb.recordset")
set objChildRS = createobject("adodb.recordset")
strSQL = "SHAPE APPEND" & _
           "  NEW adVarChar(255) AS SOADDisplayName, " & _
           "  NEW adVarChar(255) AS SOADDistName, " & _
           "  NEW adVarChar(255) AS SOADDMailStore, " & _
           "  NEW adVarChar(255) AS SOADDStorageGroup, " & _
           "  NEW adVarChar(255) AS SOADDLegExchangeDN, " & _
           " ((SHAPE APPEND  " & _
 	   "      NEW adVarChar(255) AS WMILegacyDN, " & _
	   "      NEW adVarChar(255) AS WMISize, " & _
	   "      NEW adVarChar(255) AS WMIStorename) " & _
           "       RELATE SOADDLegExchangeDN TO WMILegacyDN) AS MOWMI " 
objParentRS.LockType = 3
objParentRS.Open strSQL, conn1
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
reportonsg(sg1dn)

Set objChildRS = objParentRS("MOWMI").Value
strWinMgmts ="winmgmts:{impersonationLevel=impersonate}!//"& servername &"/root/MicrosoftExchangeV2"
Set objWMIExchange = GetObject(strWinMgmts)
Set listExchange_MailboxSizes = objWMIExchange.ExecQuery("Select * FROM Exchange_Mailbox where size > 51200",,48)
For each objExchange_MailboxSize in listExchange_MailboxSizes
		objChildRS.addnew 
		objChildRS("WMILegacyDN") = objExchange_MailboxSize.LegacyDN
		objChildRS("WMISize") = objExchange_MailboxSize.size
		objChildRS("WMIStorename") = objExchange_MailboxSize.storename
		objChildRS.update
Next
wscript.echo "finished Exchange WMI query"
objParentRS.MoveFirst
Do While Not objParentRS.EOF
    Set objChildRS = objParentRS("MOWMI").Value 
    Do While Not objChildRS.EOF
	wscript.echo objParentRS("SOADDisplayName").value & "	" & formatnumber(objChildRS("WMISize")/1024,2) & "	" & objParentRS("SOADDStorageGroup").value & "	" & objParentRS("SOADDMailStore").value
	wfile.writeline(objParentRS("SOADDisplayName").value & "," & replace(formatnumber(objChildRS("WMISize")/1024,2),",","") & "," &  objParentRS("SOADDStorageGroup").value & "," &  objParentRS("SOADDMailStore").value)  
        objChildRS.MoveNext
    Loop
    objParentRS.MoveNext
Loop
wfile.close
Wscript.echo "CSV file created "


sub reportonSG(sgdn)
set sgroup = getobject("LDAP://" & sg1dn)
svcQuery = "<LDAP://" & sg1dn & ">;(objectCategory=msExchPrivateMDB);name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof
	wscript.echo rs.fields("distinguishedName")
	GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(homeMDB=" & rs.fields("distinguishedName") & ")) )))))"
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,displayname,legacyExchangeDN,homemdb;subtree"
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not rs1.eof
		objParentRS.addnew 
		objParentRS("SOADDisplayName") = rs1.fields("displayname")
		objParentRS("SOADDistName") = left(rs1.fields("distinguishedName"),255)
		objParentRS("SOADDMailStore") = rs.fields("name")
		objParentRS("SOADDStorageGroup") = sgroup.cn
		objParentRS("SOADDLegExchangeDN") = rs1.fields("legacyExchangeDN")
		objParentRS.update	
		rs1.movenext
	wend
	rs.movenext
wend	

wscript.echo "finished AD Query"


end sub

