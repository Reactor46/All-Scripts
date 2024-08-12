set WshShell = CreateObject("WScript.Shell")

servername = wscript.arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\MeetingDelgatesForwards.csv",2,true)
wfile.writeline("Mailbox,ForwadingAddress,Status")
wfile.close
set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefaultNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
svcQuery = "<LDAP://" & strNameingContext & ">;(&(objectCategory=msExchExchangeServer)(cn=" & Servername & "));cn,name,legacyExchangeDN;subtree"
Com.ActiveConnection = Conn
Com.CommandText = svcQuery
Set Rs = Com.Execute
while not rs.eof	
	GALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)(objectClass=user)(msExchHomeServerName=" & rs.fields("legacyExchangeDN") & ")) )))))"
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,mailnickname,mail;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		cmdexe = "cscript.exe c:\temp\crul1.vbs " & servername & " " & rs1.fields("mail")  
		ef =  WshShell.run(cmdexe,0,true)
		wscript.echo rs1.fields("mail")
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
set fso = nothing
set conn = nothing
set com = nothing
wscript.echo "Done"