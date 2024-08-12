<SCRIPT LANGUAGE="VBScript">

Sub ExStoreEvents_OnSave(pEventInfo, bstrURLItem, lFlags)
Const EVT_NEW_ITEM = 1
Const EVT_IS_DELIVERED = 8

If (lFlags And EVT_IS_DELIVERED) Or (lFlags And EVT_NEW_ITEM) Then

set msgobj1 = createobject("CDO.Message")
msgobj1.datasource.open bstrURLItem,,3
strComputerName = search_user(replace(replace(msgobj1.fields("urn:schemas:httpmail:fromemail"),">",""),"<",""),2)
if len(strComputerName) <> 0 then
strmessid = replace(replace(msgobj1.fields("urn:schemas:mailheader:message-id"),">",""),"<","")  'Message ID to search for
Const cWMINameSpace = "root/MicrosoftExchangeV2"
Const cWMIInstance = "Exchange_MessageTrackingEntry"
objdlladdress = ""
strWinMgmts = "winmgmts:{impersonationLevel=impersonate}!//" & strComputerName & "/" & cWMINameSpace
Set objWMIExchange =  GetObject(strWinMgmts)
Set listExchange_MessageTrackingEntries = objWMIExchange.ExecQuery("Select * FROM Exchange_MessageTrackingEntry where entrytype = '1024' and MessageID = '" _
& strmessid & "'",,48)
For each objExchange_MessageTrackingEntry in listExchange_MessageTrackingEntries
 for i = 0 to objExchange_MessageTrackingEntry.RecipientCount
	objtempaddr = search_user(objExchange_MessageTrackingEntry.RecipientAddress(i),1)
	if len(objtempaddr) <> 0 then
		if len(objdlladdress) <> 0 then
			objdlladdress = objdlladdress & ";" & objtempaddr
		else
			objdlladdress = objtempaddr
		end if
	end if
	objtempaddr = ""
 next
Next
if len(objdlladdress) <> 0 then msgobj1.fields("urn:schemas:mailheader:to") = objdlladdress
msgobj1.fields("http://schemas.microsoft.com/exchange/outlookmessageclass") = "IPM.Note"
msgobj1.fields.update
msgobj1.datasource.save
set objWMIExchange = nothing
set listExchange_MessageTrackingEntries = nothing
end if
set msgobj1 = nothing
end if
end sub

function search_user(senaddr,searchtype)
Set objDNS = CreateObject("ADSystemInfo")	
DomainName = LCase(objDNS.DomainDNSName)
Set oRoot = GetObject("LDAP://" & DomainName & "/rootDSE")
strDefaultNamingContext = oRoot.get("defaultNamingContext")
if instr(senaddr,"/CN=") then
	strQuery = "select distinguishedName,mail,displayname,objectCategory from 'LDAP://" & strDefaultNamingContext & "' where legacyExchangeDN = '" & senaddr & "'"
else
	strQuery = "select distinguishedName,mail,displayname,objectCategory from 'LDAP://" & strDefaultNamingContext & "' where mail = '" & senaddr & "'"
end if

Set oConn = CreateObject("ADODB.Connection") 'Create an ADO Connection
oConn.Provider = "ADsDSOOBJECT"              ' ADSI OLE-DB provider
oConn.Open "ADs Provider"

Set oComm = CreateObject("ADODB.Command") ' Create an ADO Command
oComm.ActiveConnection = oConn
oComm.Properties("Page Size") = 1000
oComm.CommandText = strQuery
Set rs = oComm.Execute
while not rs.eof
	if searchtype = 1 then
		if instr(lcase(rs.fields("objectCategory")),"cn=group,") or instr(lcase(rs.fields("objectCategory")),"cn=ms-exch-dynamic-distribution-list,")  then
			dispname = rs.fields("displayname")
			emailad = rs.fields("mail")
			search_user = chr(34) & dispname & chr(34) & "<" & emailad & ">"
		end if
	else
		strsendingusername = "LDAP://" & rs.Fields("distinguishedName") 
		set objsenduser = getobject(strsendingusername)
		inplinearray = Split(objsenduser.msExchHomeServerName, "=", -1, 1)
		strservername = inplinearray(ubound(inplinearray))
		search_user = strservername
	end if
	rs.movenext
wend


end function

</SCRIPT>
