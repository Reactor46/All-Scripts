servername = wscript.arguments(0)
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
set conn1 = createobject("ADODB.Connection")
strConnString = "Data Provider=NONE; Provider=MSDataShape"
conn1.Open strConnString		
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\exportquotas.csv",2,true)
wfile.writeline("Username,EmailAddress,StorageQuoata(MB),OverLimitQuota(MB),HardLimitQuota(MB),Mailbox Size(MB),Percent Used")
set objParentRS = createobject("adodb.recordset")
set objChildRS = createobject("adodb.recordset")
strSQL = "SHAPE APPEND" & _
           "  NEW adVarChar(255) AS SOADDisplayName, " & _
           "  NEW adVarChar(255) AS SOADDistName, " & _
           "  NEW adVarChar(255) AS SOmDBStorageQuota, " & _
           "  NEW adVarChar(255) AS SOmDBOverQuotaLimit, " & _
           "  NEW adVarChar(255) AS SOmDBOverHardQuotaLimit, " & _
           " ((SHAPE APPEND  " & _
           "      NEW adVarChar(255) AS MOADDisplayName, " & _
           "      NEW adVarChar(255) AS MOADLegacyDN, " & _
           "      NEW adVarChar(255) AS MOADHomeMDB, " & _
	   "      NEW adVarChar(255) AS MOADEmail, " & _
           "      NEW adBoolean AS MOmDBUseDefaults, " & _
           "      NEW adVarChar(255) AS MOmDBStorageQuota, " & _
           "      NEW adVarChar(255) AS MOmDBOverQuotaLimit, " & _
           "      NEW adVarChar(255) AS MOmDBOverHardQuotaLimit, " & _
           " ((SHAPE APPEND  " & _
	   "      NEW adVarChar(255) AS WMILegacyDN, " & _
	   "      NEW adVarChar(255) AS WMISize, " & _
	   "      NEW adVarChar(255) AS WMIStorename) " & _
	   "   RELATE MOADLegacyDN TO WMILegacyDN) AS MOWMI" & _
	   ")" & _
           "      RELATE SOADDistName TO MOADHomeMDB) AS rsSOMO " 
objParentRS.LockType = 3
objParentRS.Open strSQL, conn1
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,distinguishedName,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof	
	sgQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchPrivateMDB)(msExchOwningServer=" & rs.fields("distinguishedName") & "));cn,name,displayname,distinguishedName,mDBStorageQuota,mDBOverQuotaLimit,mDBOverHardQuotaLimit,legacyExchangeDN;subtree"
	Com.CommandText = sgQuery
	Set Rs1 = Com.Execute
	while not rs1.eof
		objParentRS.addnew 
		objParentRS("SOADDisplayName") = rs1.fields("displayname")
		objParentRS("SOADDistName") = left(rs1.fields("distinguishedName"),255)
		objParentRS("SOmDBStorageQuota") = rs1.fields("mDBStorageQuota")
		objParentRS("SOmDBOverQuotaLimit") = rs1.fields("mDBOverQuotaLimit")
		objParentRS("SOmDBOverHardQuotaLimit") = rs1.fields("mDBOverHardQuotaLimit")
		objParentRS.update	
		rs1.movenext
	wend
	wscript.echo "finished 1st AD query storage groups"
	GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(msExchHomeServerName=" & rs.fields("legacyExchangeDN") & ")) )))))"
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,msExchMailboxGuid,mail,displayname,mDBStorageQuota,mDBOverQuotaLimit,mDBOverHardQuotaLimit,mDBUseDefaults,msExchHomeServerName,legacyExchangeDN,homemdb;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs2 = Com.Execute
	objChildRS.LockType = 3
	Set objChildRS = objParentRS("rsSOMO").Value
	while not rs2.eof
		objChildRS.addnew 
		objChildRS("MOADDisplayName") = rs2.fields("displayname")
		objChildRS("MOADLegacyDN") = rs2.fields("legacyExchangeDN")
		objChildRS("MOADHomeMDB") = left(rs2.fields("homemdb"),255)
		objChildRS("MOADEmail") = rs2.fields("mail")
		objChildRS("MOmDBUseDefaults") = rs2.fields("mDBUseDefaults")
		objChildRS("MOmDBStorageQuota") = rs2.fields("mDBStorageQuota")
		objChildRS("MOmDBOverQuotaLimit") = rs2.fields("mDBOverQuotaLimit")
		objChildRS("MOmDBOverHardQuotaLimit") = rs2.fields("mDBOverHardQuotaLimit")
		objChildRS.update
		rs2.movenext
	wend
	wscript.echo "finished 2nd AD query mailbox"
	rs.movenext
wend
Set objgrandchild = objChildRS("MOWMI").Value
strWinMgmts ="winmgmts:{impersonationLevel=impersonate}!//"& servername &"/root/MicrosoftExchangeV2"
Set objWMIExchange = GetObject(strWinMgmts)
Set listExchange_MailboxSizes = objWMIExchange.ExecQuery("Select * FROM Exchange_Mailbox",,48)
For each objExchange_MailboxSize in listExchange_MailboxSizes
		objgrandchild.addnew 
		objgrandchild("WMILegacyDN") = objExchange_MailboxSize.LegacyDN
		objgrandchild("WMISize") = objExchange_MailboxSize.size
		objgrandchild("WMIStorename") = objExchange_MailboxSize.storename
		objgrandchild.update
Next
wscript.echo "finished 3nd Exchange WMI query"
objParentRS.MoveFirst
Do While Not objParentRS.EOF
    Set objChildRS = objParentRS("rsSOMO").Value 
    strQuota = objParentRS("SOmDBStorageQuota")  
    if isnull(objParentRS("SOmDBStorageQuota")) then strQuota = 0
    strQuotaLimit = objParentRS("SOmDBOverQuotaLimit")  
    if isnull(objParentRS("SOmDBOverQuotaLimit")) then strQuotaLimit = 0
    strHardQuotaLimit = objParentRS("SOmDBOverHardQuotaLimit")
    if isnull(objParentRS("SOmDBOverHardQuotaLimit")) then strHardQuotaLimit = 0
    Do While Not objChildRS.EOF
	strUsername = replace(objChildRS("MOADDisplayName"),",","") 
	strEmailAddress = objChildRS("MOADEmail")
	strUsedefaults = objChildRS("MOmDBUseDefaults")
	strusrQuota = objChildRS("MOmDBStorageQuota") 
	if isnull(objChildRS("MOmDBStorageQuota")) then strusrQuota = 0
   	strusrQuotaLimit =  objChildRS("MOmDBOverQuotaLimit")  
	if isnull(objChildRS("MOmDBOverQuotaLimit")) then strusrQuotaLimit = 0
    	strusrHardQuotaLimit = objChildRS("MOmDBOverHardQuotaLimit")
	if isnull(objChildRS("MOmDBOverHardQuotaLimit")) then strusrHardQuotaLimit = 0
    	Set objgrandchild = objChildRS("MOWMI").Value
	Do While Not objgrandchild.EOF
		mbsize = objgrandchild("WMISize")
		if strUsedefaults = false then
			if strusrQuota <> 0 then
				wfile.writeline  strUsername & "," & strEmailAddress & "," & replace(formatnumber(strusrQuota/1024,2),",","")  & "," & replace(formatnumber(strusrQuotaLimit/1024,2),",","") & "," & replace(formatnumber(strusrHardQuotaLimit/1024,2),",","") & "," & replace(formatnumber(mbsize/1024,2),",","") & "," & formatnumber((mbsize/strusrQuota)*100,2) & "%"
			else
				wfile.writeline  strUsername & "," & strEmailAddress & "," & replace(formatnumber(strusrQuota/1024,2),",","")  & "," & replace(formatnumber(strusrQuotaLimit/1024,2),",","") & "," & replace(formatnumber(strusrHardQuotaLimit/1024,2),",","") & "," & replace(formatnumber(mbsize/1024,2),",","") & "," & 0 & "%"
			end if
		else
			if strQuota <> 0 then
				wfile.writeline  strUsername & "," & strEmailAddress & "," & replace(formatnumber(strQuota/1024,2),",","")  & "," & replace(formatnumber(strQuotaLimit/1024,2),",","") & "," & replace(formatnumber(strHardQuotaLimit/1024,2),",","") & "," & replace(formatnumber(mbsize/1024,2),",","") & "," & formatnumber((mbsize/strQuota)*100,2) & "%"
			else
				wfile.writeline  strUsername & "," & strEmailAddress & "," & replace(formatnumber(strQuota/1024,2),",","")  & "," & replace(formatnumber(strQuotaLimit/1024,2),",","") & "," & replace(formatnumber(strHardQuotaLimit/1024,2),",","") & "," & replace(formatnumber(mbsize/1024,2),",","") & "," & 0 & "%"
			end if
		end if
		objgrandchild.MoveNext
	Loop
	objChildRS.MoveNext
    Loop
    objParentRS.MoveNext
Loop
wfile.close
Wscript.echo "CSV file created"