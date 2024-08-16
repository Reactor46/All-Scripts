set WshShell = CreateObject("WScript.Shell")
servername = wscript.arguments(0)
PR_HAS_RULES = &H663A000B
PR_URL_NAME = &H6707001E
PR_CREATOR = &H3FF8001E
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\mbxforwardingRules.csv",2,true)
wfile.writeline("Mailbox,RuleType,Condition,ForwdingAddress")
wfile.close
set wfile = nothing
set fso = nothing
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
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,mailnickname;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		 cmdexe = "c:\windows\system32\cscript.exe c:\temp\rdombxrules2.vbs " & rs1.fields("mailnickname") & " " & servername     
		 wscript.echo cmdexe           
		 ef =  WshShell.run(cmdexe,1,true)
		wscript.echo rs1.fields("mailnickname")
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
set objSession = Nothing

set conn = nothing
set com = nothing
wscript.echo "Done"








