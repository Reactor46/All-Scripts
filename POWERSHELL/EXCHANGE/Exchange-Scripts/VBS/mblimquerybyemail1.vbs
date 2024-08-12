CUserID = wscript.arguments(0)
Set objDNS = CreateObject("ADSystemInfo")	
DomainName = LCase(objDNS.DomainDNSName)
Set oRoot = GetObject("LDAP://" & DomainName & "/rootDSE")
strDefaultNamingContext = oRoot.get("defaultNamingContext")
GALQueryFilter = "(&(&(&(& (mailnickname=*) (| (&(objectCategory=person)(objectClass=user)(!(homeMDB=*))(!(msExchHomeServerName=*)))(&(objectCategory=person)(objectClass=user)(|(homeMDB=*)(msExchHomeServerName=*))) )))(objectCategory=user)(mail=" & CUserID & ")))"
strQuery = "<LDAP://" & DomainName & "/" & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,displayname,mDBStorageQuota,mDBOverQuotaLimit,mDBOverHardQuotaLimit,mDBUseDefaults,msExchHomeServerName,homemdb;subtree"
Set oConn = CreateObject("ADODB.Connection") 'Create an ADO Connection
oConn.Provider = "ADsDSOOBJECT"              ' ADSI OLE-DB provider
oConn.Open "ADs Provider"

Set oComm = CreateObject("ADODB.Command") ' Create an ADO Command
oComm.ActiveConnection = oConn
oComm.Properties("Page Size") = 1000
oComm.CommandText = strQuery

Set rs = oComm.Execute
while not rs.eof
	if rs.fields("mDBUseDefaults").value = true then
		set mbstore = getobject("LDAP://" & rs.fields("homemdb"))
		if mbstore.mDBStorageQuota = "" then 	
			strquota =  "No Quota"
		else
			strquota = formatnumber(mbstore.mDBStorageQuota/1024,2)
		end if
		if mbstore.mDBOverQuotaLimit = "" then 
			stroverquota =  "No Quota"
		else
			stroverquota = formatnumber(mbstore.mDBOverQuotaLimit/1024,2)
		end if
		if mbstore.mDBOverHardQuotaLimit = "" then 
			strHardLimit =  "No Quota"
		else
			strHardLimit = formatnumber(mbstore.mDBOverHardQuotaLimit/1024,2)
		end if
		mbsize = getmbsize(right(rs.fields("msExchHomeServerName"),len(rs.fields("msExchHomeServerName"))-(instr(rs.fields("msExchHomeServerName"),"cn=Servers/cn=")+13)),rs.fields("displayname").value)
		mbsize = formatnumber(mbsize/1024,2)
		if strquota <> "No Quota" then
			wscript.echo "Username:	" & rs.fields("displayname").value 
			wscript.echo "StorageQuota:	" & strquota & " MB"
			wscript.echo "OverQuotaLimit:	" & stroverquota & " MB"
			wscript.echo "QuotaHardLimit:	" & strHardLimit & " MB"
			wscript.echo "Mailbox Size:	" & mbsize & " MB"
			if mbsize = 0 then 
				wscript.echo "Percent Used:	0 %"
			else
				wscript.echo "Percent Used:	" & formatnumber((mbsize/strquota)*100,2) & " %"
			end if
		else
			wscript.echo "Username:	" & rs.fields("displayname").value
			Wscript.echo "StorageQuota:	No Quotas Configured"
			wscript.echo "Mailbox Size:	" & mbsize & " MB"
		end if
	else
		if isnull(rs.fields("mDBStorageQuota").value) then 	
			strquota =  "No Quota"
		else
			strquota = formatnumber(rs.fields("mDBStorageQuota").value/1024,2)
		end if
		if isnull(rs.fields("mDBOverQuotaLimit").value) then 
			stroverquota =  "No Quota"
		else
			stroverquota = formatnumber(rs.fields("mDBOverQuotaLimit").value/1024,2)
		end if
		if isnull(rs.fields("mDBOverHardQuotaLimit").value) then 
			strHardLimit =  "No Quota"
		else
			strHardLimit = formatnumber(rs.fields("mDBOverHardQuotaLimit").value/1024,2)
		end if
		mbsize = getmbsize(right(rs.fields("msExchHomeServerName"),len(rs.fields("msExchHomeServerName"))-(instr(rs.fields("msExchHomeServerName"),"cn=Servers/cn=")+13)),rs.fields("displayname").value)
		mbsize = formatnumber(mbsize/1024,2)
		if strquota <> "No Quota" then
			wscript.echo "Username:	" & rs.fields("displayname").value 
			wscript.echo "StorageQuota:	" & strquota & " MB"
			wscript.echo "OverQuotaLimit:	" & stroverquota & " MB"
			wscript.echo "QuotaHardLimit:	" & strHardLimit & " MB"
			wscript.echo "Mailbox Size:	" & mbsize & " MB"
			if mbsize = 0 then 
				wscript.echo "Percent Used:	0 %"
			else
				wscript.echo "Percent Used:	" & formatnumber((mbsize/strQuota)*100,2) & " %"
			end if
		else
			wscript.echo "Username:	" & rs.fields("displayname").value
			Wscript.echo "StorageQuota:	No Quotas Configured"
			wscript.echo "Mailbox Size:	" & mbsize & " MB"
		end if
	end if
	rs.movenext
wend

function getmbsize(strservername,strmailboxdisplay)
rem On Error Resume Next
Dim cComputerName
Const cWMINameSpace = "root/MicrosoftExchangeV2"
Const cWMIInstance = "Exchange_Mailbox"
cComputerName = strservername
Dim strWinMgmts		' Connection string for WMI
Dim objWMIExchange	' Exchange Namespace WMI object
Dim listExchange_Mailboxs	' ExchangeLogons collection
Dim objExchange_Mailbox		' A single ExchangeLogon WMI
strWinMgmts ="winmgmts:{impersonationLevel=impersonate}!//"& cComputerName&"/" & cWMINameSpace
Set objWMIExchange =  GetObject(strWinMgmts)
Set listExchange_MailboxSizes = objWMIExchange.ExecQuery("Select * FROM Exchange_Mailbox where MailboxDisplayName = '" & strmailboxdisplay & "'")
For each objExchange_MailboxSize in listExchange_MailboxSizes
	getmbsize = objExchange_MailboxSize.size
Next
end function