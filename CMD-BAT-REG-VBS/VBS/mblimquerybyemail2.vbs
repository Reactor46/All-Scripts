CUserID = wscript.arguments(0)
Set objDNS = CreateObject("ADSystemInfo")	
DomainName = LCase(objDNS.DomainDNSName)
Set oRoot = GetObject("LDAP://" & DomainName & "/rootDSE")
strDefaultNamingContext = oRoot.get("defaultNamingContext")
GALQueryFilter = "(&(&(&(& (mailnickname=*) (| (&(objectCategory=person)(objectClass=user)(!(homeMDB=*))(!(msExchHomeServerName=*)))(&(objectCategory=person)(objectClass=user)(|(homeMDB=*)(msExchHomeServerName=*))) )))(objectCategory=user)(mail=" & CUserID & ")))"
strQuery = "<LDAP://" & DomainName & "/" & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,msExchMailboxGuid,displayname,mDBStorageQuota,mDBOverQuotaLimit,mDBOverHardQuotaLimit,mDBUseDefaults,msExchHomeServerName,legacyExchangeDN,homemdb;subtree"
Set oConn = CreateObject("ADODB.Connection") 'Create an ADO Connection
oConn.Provider = "ADsDSOOBJECT"              ' ADSI OLE-DB provider
oConn.Open "ADs Provider"

Set oComm = CreateObject("ADODB.Command") ' Create an ADO Command
oComm.ActiveConnection = oConn
oComm.Properties("Page Size") = 1000
oComm.CommandText = strQuery

Set rs = oComm.Execute
while not rs.eof
	sname =  right(rs.fields("msExchHomeServerName"),len(rs.fields("msExchHomeServerName"))-(instr(rs.fields("msExchHomeServerName"),"cn=Servers/cn=")+13))
	guidary = rs.fields("msExchMailboxGuid").value
	mdbarry = split(rs.fields("homemdb"),",", -1, 1)
	mstorename = replace(lcase(mdbarry(0)),"cn=","")
	sgroupname = replace(lcase(mdbarry(1)),"cn=","")
	strwmicon = "winMgmts:!\\" & sname & "\root\MicrosoftExchangeV2:Exchange_Mailbox.LegacyDN='" & rs.fields("legacyExchangeDN")
 	strwmicon = strwmicon & "',MailboxGUID='" & Octenttohex(guidary) & "',ServerName='" & sname 
	strwmicon = strwmicon & "',StorageGroupName='" & sgroupname & "',StoreName='" & mstorename & "'"
	Set mbox = GetObject(strwmicon)
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
		mbsize = mbox.size
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
		mbsize = mbox.size
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

Function Octenttohex(OctenArry)  
  ReDim aOut(UBound(OctenArry)) 
  For i = 1 to UBound(OctenArry) + 1 
    if len(hex(ascb(midb(OctenArry,i,1)))) = 1 then 
    	aOut(i-1) = "0" & hex(ascb(midb(OctenArry,i,1)))
    else
	aOut(i-1) = hex(ascb(midb(OctenArry,i,1)))
    end if
  Next 
  ' Tranpose to AD format
  Octenttohex1 = "{" & aOut(3) & aOut(2) & aOut(1) & aOut(0) & "-" & aOut(5) & aOut(4) & "-"  & aOut(7) & aOut(6) & "-" & aOut(8) 
  Octenttohex = Octenttohex1 & aOut(9) & "-" & aOut(10) & aOut(11) & aOut(12) & aOut(13) & aOut(14) & aOut(15) & "}"
End Function 