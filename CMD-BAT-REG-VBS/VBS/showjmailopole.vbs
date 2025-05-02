servername = wscript.arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
set wfile = fso.opentextfile("c:\Junkemailsettings.csv",2,true)
wfile.writeline("Mailbox,OWAJunkEmailState,Outlook Filter Setting,Outlook Delete Junk Email Setting")
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
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,mail;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		call ProcMailbox(servername,rs1.fields("mail"))
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
wfile.close
set fso = nothing
set conn = nothing
set com = nothing
wscript.echo "Done"

Sub ProcMailbox(snServername,mnMailboxname)
wscript.echo "Processing : " & mnMailboxname
ffnd = 0
SourceURL = "http://" & snServername & "/exchange/" & mnMailboxname & "/inbox/"
SSql = "SELECT ""DAV:href"", ""DAV:displayname"", ""http://schemas.microsoft.com/mapi/proptag/0x65E90003"", ""http://schemas.microsoft.com/mapi/proptag/0x61010003"", "
SSql = SSql & """http://schemas.microsoft.com/mapi/proptag/0x61020003"" "
SSql = SSql & "FROM scope('shallow traversal of """ & SourceURL & """') " 
SSql = SSql & " Where ""DAV:isfolder"" = false AND ""DAV:ishidden"" = true AND ""http://schemas.microsoft.com/exchange/outlookmessageclass"" = 'IPM.ExtendedRule.Message' " 
SSql = SSql & "AND ""http://schemas.microsoft.com/mapi/proptag/0x65EB001E"" = 'JunkEmailRule'"
Set oConn = CreateObject("ADODB.Connection")
oConn.Provider = "Exoledb.DataSource"
oConn.Open SourceURL
Set oRecSet = CreateObject("ADODB.Recordset")
oRecSet.CursorLocation = 3
oRecSet.Open SSql, oConn.ConnectionString
if err.number <> 0 then wfile.writeline(user & "," & "Error Connection to Mailbox")
While oRecSet.EOF <> True
	ffnd = 1
	filterlist = oRecSet.fields("http://schemas.microsoft.com/mapi/proptag/0x61010003")
	if filterlist = "" then 
		flist = "No Automatic filtering"
	else
		select case Filterlist
			case 6 flist = "Low"
			case 3 flist = "High"
			case -2147483648 flist = "Safe Lists only"
		end select 
	end if
	delist = oRecSet.fields("http://schemas.microsoft.com/mapi/proptag/0x61020003")
	if delist = "" then
		delset = "Disabled"
	else
		select case delist
			case 1 delset = "Enabled"
			case 0 delset = "Disabled"
		end select
	end if
	Rule_State = oRecSet.fields("http://schemas.microsoft.com/mapi/proptag/0x65E90003")
	select case Rule_State
		case 48 wfile.writeline(mnMailboxname & "," & "OWA Junk Email Setting Off," & flist & "," & delset)
		case 49 wfile.writeline(mnMailboxname & "," & "OWA Junk Email Setting Enabled," & flist & "," & delset)
   	end Select
oRecSet.movenext
	
wend
if ffnd = 0 then wfile.writeline(mnMailboxname & "," & "No Junk Email rules Exists")
end sub