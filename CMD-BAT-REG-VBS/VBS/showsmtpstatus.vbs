set conn = createobject("ADODB.Connection")
set com = createobject("ADODB.Command")
Set iAdRootDSE = GetObject("LDAP://RootDSE")
strNameingContext = iAdRootDSE.Get("configurationNamingContext")
strDefNamingContext = iAdRootDSE.Get("defaultNamingContext")
Conn.Provider = "ADsDSOObject"
Conn.Open "ADs Provider"
Com.ActiveConnection = Conn
Com.ActiveConnection = Conn
Wscript.echo 
Wscript.echo "SMTP Virtual Servers Status"
vsQuery = "<LDAP://" & strNameingContext & ">;(objectCategory=protocolCfgSMTPServer);name,distinguishedName;subtree"
Com.ActiveConnection = Conn
Com.CommandText = vsQuery
Set Rs = Com.Execute
While Not Rs.EOF
	strstmsrv = "LDAP://" & rs.fields("distinguishedName")
	set svsSmtpserver = getobject(strstmsrv)
	wscript.echo
	wscript.echo "ServerName:" & mid(svsSmtpserver.distinguishedName,instr(svsSmtpserver.distinguishedName,"CN=Protocols,")+16,instr(svsSmtpserver.distinguishedName,",CN=Servers")-(instr(svsSmtpserver.distinguishedName,"CN=Protocols,")+16))
	call getSTMPstatus(mid(svsSmtpserver.distinguishedName,instr(svsSmtpserver.distinguishedName,"CN=Protocols,")+16,instr(svsSmtpserver.distinguishedName,",CN=Servers")-(instr(svsSmtpserver.distinguishedName,"CN=Protocols,")+16)),svsSmtpserver.adminDisplayName)
	rs.movenext
wend

sub getSTMPstatus(servername,vsname)
Set SMTPVSS = GetObject("IIS://" & Servername & "/SMTPSVC")
for each SMTPVS in SMTPVSS
  if SMTPVS.KeyType = "IIsSmtpServer" then
	if SMTPVS.ServerComment = vsname then
   		wscript.echo "SMTP Server : " & SMTPVS.ServerComment
		select case SMTPVS.ServerState
 			 case 1 Wscript.echo "Current State: Starting"
 			 case 2 Wscript.echo "Current State: Started"
 		 	 case 3 Wscript.echo "Current State: Stopping"
 			 case 4 Wscript.echo "Current State: Stopped"
  			 case 5 Wscript.echo "Current State: Pausing"
 			 case 6 Wscript.echo "Current State: Paused"
 			 case 7 Wscript.echo "Current State: Continuing"
 			 case else Wscript.echo "unknown"
		end select
 	end if
  end if
next

end sub
