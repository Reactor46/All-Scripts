servername = wscript.arguments(0)
domainname = wscript.arguments(1)
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
SourceURL = "http://" & snServername & "/exadmin/" & domainname & "/mbx/" & mnMailboxname & "/inbox/"
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"", ""http://schemas.microsoft.com/mapi/proptag/x65E90003"", "
strQuery = strQuery & """http://schemas.microsoft.com/mapi/proptag/x61010003"", ""http://schemas.microsoft.com/mapi/proptag/x61020003""" 
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & SourceURL & """') Where ""DAV:ishidden"" = True AND ""DAV:isfolder"" = False AND "
strQuery = strQuery & """http://schemas.microsoft.com/exchange/outlookmessageclass"" = 'IPM.ExtendedRule.Message' "
strQuery = strQuery & "AND ""http://schemas.microsoft.com/mapi/proptag/0x65EB001E"" = 'JunkEmailRule' </D:sql></D:searchrequest>"
set req = createobject("microsoft.xmlhttp")
req.open "SEARCH", SourceURL, false,"", ""
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.send strQuery
wscript.echo "Status: " & req.status
wscript.echo "Status text: " & req.statustext
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:href")
   set ostateList = oResponseDoc.getElementsByTagName("d:x65E90003")
   set ofilterList = oResponseDoc.getElementsByTagName("d:x61010003")
   set odeleteList = oResponseDoc.getElementsByTagName("d:x61020003")
   For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	Filterlist = ofilterList(i).nodetypedvalue
	if filterlist = "" then 
		flist = "No Automatic filtering"
	else
		select case Filterlist
			case 6 flist = "Low"
			case 3 flist = "High"
			case -2147483648 flist = "Safe Lists only"
		end select 
	end if
	delist = odeleteList(i).nodetypedvalue
	if delist = "" then
		delset = "Disabled"
	else
		select case delist
			case 1 delset = "Enabled"
			case 0 delset = "Disabled"
		end select
	end if
	Rule_State = ostateList(i).nodetypedvalue
	select case Rule_State
		case 48 wfile.writeline(mnMailboxname & "," & "OWA Junk Email Setting Off," & flist & "," & delset)
		case 49 wfile.writeline(mnMailboxname & "," & "OWA Junk Email Setting Enabled," & flist & "," & delset)
   	end Select
   Next	
   if oNodeList.length = 0 then
	wfile.writeline(mnMailboxname & "," & "No Junk Email rules Exists")
   End if
Else
End If

end sub