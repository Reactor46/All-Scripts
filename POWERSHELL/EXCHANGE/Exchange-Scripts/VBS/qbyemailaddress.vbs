CUserID = wscript.arguments(0)
Set objDNS = CreateObject("ADSystemInfo")	
DomainName = LCase(objDNS.DomainDNSName)
Set oRoot = GetObject("LDAP://" & DomainName & "/rootDSE")
strDefaultNamingContext = oRoot.get("defaultNamingContext")
GALQueryFilter = "(&(&(&(& (mailnickname=*) (| (&(objectCategory=person)(objectClass=user)(!(homeMDB=*))(!(msExchHomeServerName=*)))(&(objectCategory=person)(objectClass=user)(|(homeMDB=*)(msExchHomeServerName=*))) )))(objectCategory=user)(mail=" & CUserID & ")))"
strQuery = "<LDAP://" & DomainName & "/" & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,displayname,TelephoneNumber,ipPhone,homephone,pager,facsimiletelephonenumber,mobile,postalcode,GivenName,sn,title;subtree"
Set oConn = CreateObject("ADODB.Connection") 'Create an ADO Connection
oConn.Provider = "ADsDSOOBJECT"              ' ADSI OLE-DB provider
oConn.Open "ADs Provider"

Set oComm = CreateObject("ADODB.Command") ' Create an ADO Command
oComm.ActiveConnection = oConn
oComm.Properties("Page Size") = 1000
oComm.CommandText = strQuery

Set rs = oComm.Execute
while not rs.eof
	wscript.echo rs.fields("displayname")
	wscript.echo rs.fields("TelephoneNumber")
	wscript.echo rs.fields("mobile")
	wscript.echo rs.fields("facsimiletelephonenumber")
	wscript.echo
	rs.movenext
wend