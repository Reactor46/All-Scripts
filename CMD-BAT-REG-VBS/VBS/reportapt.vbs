on error resume next
servername = "SERVERNAME"
public username
public password
username = "USERNAME"
password = "PASSWORD"
public datefrom
public dateto
datefrom = "2007-03-11T00:00:00Z"
dateto = "2007-04-01T00:00:00Z"

set shell = createobject("wscript.shell")
set conn1 = createobject("ADODB.Connection")
Set fso = CreateObject("Scripting.FileSystemObject")
fname = "c:\support\scripts\" & servername & ".csv"
set wfile = fso.opentextfile(fname,2,true)
wfile.writeline("Meeting,Organizer")

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
	GALQueryFilter =  "(&(&(&(&(mailnickname=*)(!msExchHideFromAddressLists=TRUE)(|(&(objectCategory=person)(objectClass=user)(msExchHomeServerName=" &rs.fields("legacyExchangeDN") & ")) )))))"
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,mail,displayname,mailnickname;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		wscript.echo "User: " & rs1.fields("displayname")
		user = rs1.fields("mail")
		call QueryAttendees(servername,user)		
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
set conn = nothing
set com = nothing
wscript.echo "Done"


Public Sub QueryAttendees(server,mailbox)

On Error Resume Next

strURL = "http://" & server & "/exchange/" & mailbox & "/calendar/"
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT DISTINCT ""DAV:href"" FROM scope('shallow traversal of """ & strURL & """') "
strQuery = strQuery & " Where ""DAV:isfolder"" = false AND ""DAV:ishidden"" = false AND ""urn:schemas:calendar:alldayevent"" = false "
strQuery = strQuery & "AND ""DAV:contentclass"" = 'urn:content-classes:appointment' "
strQuery = strQuery & "AND ""urn:schemas:calendar:dtend"" &gt; CAST(""" & datefrom & """ as 'dateTime.tz') "
strQuery = strQuery & "AND ""urn:schemas:calendar:dtstart"" &lt; CAST(""" & dateto & """ as 'dateTime.tz') "
strQuery = strQuery & "</D:sql></D:searchrequest>"


wscript.echo strQuery
set req = createobject("microsoft.xmlhttp")
req.open "SEARCH", strURL, false, username, password

  If Err.Number <> 0 Then
      WScript.Echo "Error Opening Search"
      WScript.Echo Err.Number & ": " & Err.Description
  End If

req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.setRequestHeader "Depth", "1,noroot"
req.send strQuery

  If Err.Number <> 0 Then
      WScript.Echo "Error Sending Query"
      WScript.Echo Err.Number & ": " & Err.Description
   End If

wscript.echo req.status
wscript.echo "response" & req.responseXML

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

End Sub

public sub proccalmess(objhref)

set req = createobject("microsoft.xmlhttp")
wscript.echo objhref
wfile.write(objhref & ",")
On Error Resume Next
Req.open "GET", objhref, false,username, password
  If Err.Number <> 0 Then
      WScript.Echo "Error Opening GET"
      WScript.Echo Err.Number & ": " & Err.Description
   End If

Req.setRequestHeader "Translate","f"
Req.send

attendeearry = split(req.responsetext,"ORGANIZER;",-1,1)
for i = 1 to ubound(attendeearry)
string1 = vbcrlf & " "
stparse = replace(attendeearry(i),string1,"")
attaddress = mid(stparse,(instr(stparse,"MAILTO:")+7),instr(stparse,chr(13)))
attaddress = mid(attaddress,1,instr(attaddress,vbcrlf))
wscript.echo attaddress
wfile.writeline(attaddress)
next

end sub