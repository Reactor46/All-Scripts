set Req = createobject("Microsoft.XMLHTTP")
Req.open "GET","http://server/exchange/mailbox/inbox/calandermessage.EML",false
Req.setRequestHeader "Translate","f"
Req.send
attendeearry = split(req.responsetext,"ORGANIZER;",-1,1)
for i = 1 to ubound(attendeearry)
	string1 = vbcrlf & " "
	stparse = replace(attendeearry(i),string1,"")
	attaddress = mid(stparse,(instr(stparse,"MAILTO:")+7),instr(stparse,chr(13)))
	attaddress = mid(attaddress,1,(instr(attaddress,vbcrlf)-1))
next
uidarry = mid(req.responsetext,instr(req.responsetext,"UID:")+3,len(req.responsetext))
string1 = vbcrlf & " "
stparse = replace(uidarry,string1,"")
uidprop = mid(stparse,2,instr(stparse,vbcrlf))
uidprop = replace(uidprop,vbcrlf,"")
CUserID = replace(attaddress," ","")
Set objDNS = CreateObject("ADSystemInfo")	
DomainName = LCase(objDNS.DomainDNSName)
Set oRoot = GetObject("LDAP://" & DomainName & "/rootDSE")
strDefaultNamingContext = oRoot.get("defaultNamingContext")
GALQueryFilter = "(&(&(&(& (mailnickname=*) (| (&(objectCategory=person)(objectClass=user)(!(homeMDB=*))(!(msExchHomeServerName=*)))(&(objectCategory=person)(objectClass=user)(|(homeMDB=*)(msExchHomeServerName=*))) )))(objectCategory=user)(mail=" & CUserID & ")))"
strQuery = "<LDAP://" & DomainName & "/" & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,msExchHomeServerName,msExchHideFromAddressLists;subtree"
Set oConn = CreateObject("ADODB.Connection") 'Create an ADO Connection
oConn.Provider = "ADsDSOOBJECT"              ' ADSI OLE-DB provider
oConn.Open "ADs Provider"

Set oComm = CreateObject("ADODB.Command") ' Create an ADO Command
oComm.ActiveConnection = oConn
oComm.Properties("Page Size") = 1000
oComm.CommandText = strQuery

Set rs = oComm.Execute

server = right(rs.fields("msExchHomeServerName"),len(rs.fields("msExchHomeServerName"))-(instr(rs.fields("msExchHomeServerName"),"cn=Servers/cn=")+13))
mailbox = attaddress
strURL = "http://" & server & "/exchange/" & mailbox & "/calendar/"
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT ""DAV:href"" FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""urn:schemas:calendar:uid"" = '" & uidprop & "'</D:sql></D:searchrequest>"
set req = createobject("microsoft.xmlhttp")
req.open "SEARCH", strURL, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.send strQuery
If req.status >= 500 Then
   wscript.echo "Status: " & req.status
   wscript.echo "Status text: An error occurred on the server."
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:href")
   For i = 0 To (oNodeList.length -2)
	set oNode = oNodeList.nextNode
	proccalmess(oNode.Text) 
   Next
Else
   wscript.echo "Status: " & req.status
   wscript.echo "Status text: " & req.statustext
   wscript.echo "Response text: " & req.responsetext
End If

sub proccalmess(objhref)

Req.open "GET", objhref, false
Req.setRequestHeader "Translate","f"
Req.send
attendeearry = split(req.responsetext,"ATTENDEE;",-1,1)
for i = 1 to ubound(attendeearry)
string1 = vbcrlf & " "
stparse = replace(attendeearry(i),string1,"")
attaddress = mid(stparse,(instr(stparse,"MAILTO:")+7),instr(stparse,chr(13)))
attaddress = mid(attaddress,1,instr(attaddress,vbcrlf))
if instr(stparse,"=RESOURCE") then
	wscript.echo attaddress
end if
next

end sub
