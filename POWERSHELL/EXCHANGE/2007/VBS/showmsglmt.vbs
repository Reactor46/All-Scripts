set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
gsQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchMessageDeliveryConfig);name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = gsQuery
Set Rs = Com.Execute
Wscript.echo "Global Settings - Message Delivery Properties"
Wscript.echo 
While Not Rs.EOF
	strconfcont = "LDAP://" & rs.fields("distinguishedName")
	set ccConfig = getobject(strconfcont)
	wscript.echo "Sending Message Size Limit: " & ccConfig.submissionContLength  & " KB"
	wscript.echo "Recieving Message Size Limit: " & ccConfig.delivContLength &  " KB"
	wscript.echo "Recipient Limits: " & ccConfig.msExchRecipLimit 
	rs.movenext
wend
Wscript.echo 
Wscript.echo "Connector Settings"
wscript.echo 
vsQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=msExchRoutingSMTPConnector);name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = vsQuery
Set Rs = Com.Execute
While Not Rs.EOF
	strconnect = "LDAP://" & rs.fields("distinguishedName")
	set cnCconnect = getobject(strconnect)
	wscript.echo "Connector Name:" & cnCconnect.cn
	wscript.echo "Max Message Size Limit:" & cnCconnect.delivContLength & " KB"
	wscript.echo
	rs.movenext
wend
Wscript.echo 
Wscript.echo "SMTP Virtual Server Settings"
wscript.echo 
vsQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=protocolCfgSMTPServer);name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = vsQuery
Set Rs = Com.Execute
While Not Rs.EOF
	strstmsrv = "LDAP://" & rs.fields("distinguishedName")
	set svsSmtpserver = getobject(strstmsrv)
	wscript.echo "ServerName:" & mid(svsSmtpserver.distinguishedName,instr(svsSmtpserver.distinguishedName,"CN=Protocols,")+16,instr(svsSmtpserver.distinguishedName,",CN=Servers")-(instr(svsSmtpserver.distinguishedName,"CN=Protocols,")+16))
	wscript.echo "Virtual Server Name:" & svsSmtpserver.adminDisplayName
	wscript.echo "Max Message Size Limit:" & svsSmtpserver.msExchSmtpMaxMessageSize/1024 & " KB"
	wscript.echo "Recipient Limits:" & svsSmtpserver.msExchSmtpMaxRecipients
	wscript.echo 
	rs.movenext
wend
Wscript.echo 
Wscript.echo "Users with Sending Limits"
wscript.echo 
srquery = "<LDAP://" & strDefNamingContext & ">;(&(&(objectCategory=Person)(objectclass=user)(submissionContLength=*)));name,displayname,distinguishedName,submissionContLength;subtree"
Com.ActiveConnection = Conn
Com.CommandText = srquery
Set Rs = Com.Execute
While Not Rs.EOF
	wscript.echo "User:" & rs.fields("displayname")
	wscript.echo "Sending Message Size Limit:" & rs.fields("submissionContLength") & " KB"
	wscript.echo
	rs.movenext
wend
Wscript.echo 
Wscript.echo "Users with Receiving Limits"
wscript.echo 
srquery = "<LDAP://" & strDefNamingContext & ">;(&(&(objectCategory=Person)(objectclass=user)(delivContLength=*)));name,displayname,distinguishedName,delivContLength;subtree"
Com.ActiveConnection = Conn
Com.CommandText = srquery
Set Rs = Com.Execute
While Not Rs.EOF
	wscript.echo "User:" & rs.fields("displayname")
	wscript.echo "Receiving Message Size Limit:" & rs.fields("delivContLength") & " KB"
	wscript.echo
	rs.movenext
wend