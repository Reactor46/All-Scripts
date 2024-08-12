lfcount = 0
lfsize = 0
lfoldatenum = ""
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\SMTPVSSettings.csv",2,true)
wfile.writeline("Servername,Virtual Server Name,Logging Type,Log File Dir,Number of LogFiles,Space Used(MB),Oldest Log File")
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
Wscript.echo "SMTP Virtual Servers Logfile Setting and Resources"
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
wfile.close

sub getSTMPstatus(servername,vsname)
Set SMTPVSS = GetObject("IIS://" & Servername & "/SMTPSVC")
for each SMTPVS in SMTPVSS
  if SMTPVS.KeyType = "IIsSmtpServer" then
	if SMTPVS.ServerComment = vsname then
   		wscript.echo "SMTP Server : " & SMTPVS.ServerComment
		if SMTPVS.logtype = 0 then
			wscript.echo "Logging not enabled"
			wfile.writeline(servername & "," & SMTPVS.ServerComment & "," & "Logging not enabled,,,,")
		else
			select case SMTPVS.logpluginclsid
				case "{FF160663-DE82-11CF-BC0A-00AA006111E0}" Wscript.echo "Logging Type : W3C Extended Log File Format"
					lfiles = 1
					ltype = "W3C Extended Log File Format"
				case "{FF16065B-DE82-11CF-BC0A-00AA006111E0}" Wscript.echo "Logging Type : ODBC Logging"
					wfile.writeline(servername & "," & SMTPVS.ServerComment & "," & "ODBC Logging,,,,")
				case "{FF16065F-DE82-11CF-BC0A-00AA006111E0}" Wscript.echo "Logging Type : NCSA Log File Format"
					lfiles = 1
					ltype = "NCSA Log File Format"
				case "{FF160657-DE82-11CF-BC0A-00AA006111E0}" Wscript.echo "Logging Type : Microsoft IIS Log File Format"
					lfiles = 1
					ltype = "Microsoft IIS Log File Format"
			end select
			if lfiles = 1 then
				if isnull(SMTPVS.logfiledirectory) then 
						wscript.echo "Log File Directory : " & SMTPVSS.logfiledirectory & "SMTPSVC" & SMTPVS.name
						lfiledir = SMTPVSS.logfiledirectory & "SMTPSVC" & SMTPVS.name
						logfileunc = "\" & mid(SMTPVSS.logfiledirectory,3,len(SMTPVSS.logfiledirectory)-2) & "SMTPSVC" & SMTPVS.name
				else
					if mid(SMTPVS.logfiledirectory,len(SMTPVS.logfiledirectory),1) = "\" then
						lfiledir = SMTPVS.logfiledirectory & "SMTPSVC" & SMTPVS.name
						wscript.echo "Log File Directory : " & SMTPVS.logfiledirectory & "SMTPSVC" & SMTPVS.name
						logfileunc = "\" & mid(SMTPVS.logfiledirectory,3,len(SMTPVS.logfiledirectory)-2) & "SMTPSVC" & SMTPVS.name
					else
						lfiledir = SMTPVS.logfiledirectory & "\" & "SMTPSVC" & SMTPVS.name
						wscript.echo "Log File Directory : " & SMTPVS.logfiledirectory & "\" & "SMTPSVC" & SMTPVS.name
						logfileunc = "\" & mid(SMTPVS.logfiledirectory,3,len(SMTPVS.logfiledirectory)-2) & "\" & "SMTPSVC" & SMTPVS.name
					end if
				end if
				drive = left(SMTPVS.logfiledirectory,1) & ":"
				lfilepath = replace(logfileunc,"\","\\")
				lfilepath = right(lfilepath,(len(lfilepath)-2)) & "\\"
				Set objWMIService = GetObject("winmgmts:" _
 				& "{impersonationLevel=impersonate}!\\" & servername & "\root\cimv2")
				Set lfiles = objWMIService.ExecQuery _
				 ("select * from CIM_DataFile where path = """ & lfilepath & """ and extension = ""log"" and drive = """ & drive & """")
				for each lfile in lfiles
						lfcount = lfcount + 1
						lfsize = lfsize + lfile.filesize
						if lfcount = 1 then lfolddatenum = cdate(DateSerial(Left(lfile.lastmodified, 4), Mid(lfile.lastmodified, 5, 2), Mid(lfile.lastmodified, 7, 2)) & " " & timeserial(Mid(lfile.lastmodified, 9, 2),Mid(lfile.lastmodified, 11, 2),Mid(lfile.lastmodified,13, 2)))
						if lfolddatenum > lfile.LastModified then
							lfolddatenum = cdate(DateSerial(Left(lfile.lastmodified, 4), Mid(lfile.lastmodified, 5, 2), Mid(lfile.lastmodified, 7, 2)) & " " & timeserial(Mid(lfile.lastmodified, 9, 2),Mid(lfile.lastmodified, 11, 2),Mid(lfile.lastmodified,13, 2)))
						end if
				next
				wscript.echo "Number of Log files in Directory : " & lfcount
				wscript.echo "Disk Space being used : " & formatnumber(lfsize/1048576,2,0,0,0) & " MB"
				wscript.echo "Oldest Log file in this Directory : " & lfolddatenum
				wfile.writeline(servername & "," & SMTPVS.ServerComment & "," & ltype & "," & lfiledir & "," & lfcount & "," & formatnumber(lfsize/1048576,2,0,0,0) & "," & lfolddatenum)
				lfcount = 0
				lfsize = 0
				lfolddatenum = ""
				ltype = ""
				lfiledir = ""
				
			end if
			lfiles = 0
		end if
 	end if
  end if
next

end sub
