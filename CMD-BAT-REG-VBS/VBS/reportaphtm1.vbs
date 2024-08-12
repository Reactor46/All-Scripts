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
trReportBody = ""

set shell = createobject("wscript.shell")
set conn1 = createobject("ADODB.Connection")


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
set com = Nothing
rpReport = rpReport & "<table border=""1"" width=""100%"">" & vbcrlf
rpReport = rpReport & "  <tr>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""15%""><b><font color=""#FFFFFF"">Start Time</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""15%""><b><font color=""#FFFFFF"">End time</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""30%""><b><font color=""#FFFFFF"">Subject</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""10%""><b><font color=""#FFFFFF"">Location</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""10%""><b><font color=""#FFFFFF"">Organizer</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""10%""><b><font color=""#FFFFFF"">Free/Busy</font></b></td>" & vbcrlf
rpReport = rpReport & "<td align=""center"" bgcolor=""#000080"" width=""10%""><b><font color=""#FFFFFF"">New Clients</font></b></td>" & vbcrlf
rpReport = rpReport & "</tr>" & vbcrlf
rpReport = rpReport & trReportBody 
rpReport = rpReport & "</table>" & vbcrlf
Set fso = CreateObject("Scripting.FileSystemObject")

set wfile = fso.opentextfile("c:\temp\" & servername  & ".htm",2,true) 
wfile.write rpReport
wfile.close
set wfile = nothing
set fso = Nothing
wscript.echo "Done"


Public Sub QueryAttendees(server,mailbox)

On Error Resume Next

strURL = "http://" & server & "/exchange/" & mailbox & "/calendar/"
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:""  xmlns:b=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"">"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"",  ""urn:schemas:httpmail:subject"", "
strQuery = strQuery & """DAV:creationdate"", "
strQuery = strQuery & """http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"" As EntryID, "
strQuery = strQuery & """urn:schemas:httpmail:fromname"",  ""urn:schemas:calendar:dtstart"", ""urn:schemas:calendar:dtend"", "
strQuery = strQuery & """http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8205"" As BusyStatus,"
strQuery = strQuery & """http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x825E"" As NewClients,"
strQuery = strQuery & " ""urn:schemas:calendar:location"", ""http://schemas.microsoft.com/mapi/apptstateflags"" FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:contentclass"" = 'urn:content-classes:appointment' AND "
strQuery = strQuery & " NOT ""urn:schemas:calendar:instancetype"" = 1 AND " 
strQuery = strQuery & """urn:schemas:calendar:dtstart"" &lt;= CAST(""" & dateto & """ as 'dateTime') AND "
strQuery = strQuery & """urn:schemas:calendar:dtend"" &gt;= CAST(""" & datefrom & """ as 'dateTime')</D:sql></D:searchrequest>"


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
   set oDisplayNameNodes = oResponseDoc.getElementsByTagName("a:displayname")
   set oHrefNodes = oResponseDoc.getElementsByTagName("a:href")
   set oSubject = oResponseDoc.getElementsByTagName("d:subject")
   set oEndTime = oResponseDoc.getElementsByTagName("e:dtend")
   Set oStartTime = oResponseDoc.getElementsByTagName("e:dtstart")
   Set oLocation = oResponseDoc.getElementsByTagName("e:location")
   Set oAppstate = oResponseDoc.getElementsByTagName("f:apptstateflags")
   Set oFromname = oResponseDoc.getElementsByTagName("d:fromname")
   Set oEntryID = oResponseDoc.getElementsByTagName("EntryID")
   Set oBusyStatus = oResponseDoc.getElementsByTagName("BusyStatus")
   Set oNewClients = oResponseDoc.getElementsByTagName("NewClients")
    For i = 0 To (oDisplayNameNodes.length -1)
		set oNode = oDisplayNameNodes.nextNode
		set oNode1 = oHrefNodes.nextNode
		set oNode2 = oSubject.nextNode
		set oNode3 = oEndTime.nextNode
		Set oNode4 = oStarttime.nextNode
		Set oNode5 = oLocation.nextNode
		Set oNode6 = oAppstate.nextNode
		Set oNode7 = oFromname.nextNode
		Set oNode8 = oEntryID.nextNode
		Set oNode9 = oBusyStatus.nextNode
		Set oNode10 = oNewClients.nextNode
		if oNode10.text = "" then
			ncNewclients = "False"
		else
			ncNewclients = "True"
		end if
		wscript.echo Octenttohex(oNode8.nodeTypedValue)
		soOrgnizer = ""
		soOrgnizer = oNode7.text
		sdStartDate = dateadd("h",toffset,DateSerial(Mid(oNode4.text,1,4),Mid(oNode4.text,6,2),Mid(oNode4.text,9,2)) & " " & Mid(oNode4.text,12,8))
		edEndDate = dateadd("h",toffset,DateSerial(Mid(oNode3.text,1,4),Mid(oNode3.text,6,2),Mid(oNode3.text,9,2)) & " " & Mid(oNode3.text,12,8))
		wscript.echo soOrgnizer
		wscript.echo 
		trReportBody = trReportBody  & "<tr>" & vbcrlf
		trReportBody  = trReportBody  & "<td align=""center"" width=""15%"">" & sdStartDate &  "&nbsp;</td>" & vbcrlf
		trReportBody  = trReportBody  & "<td align=""center"" width=""15%"">" & edEndDate & "&nbsp;</td>" & vbcrlf
		trReportBody  = trReportBody  & "<td align=""center"" width=""30%""><a href=""outlook:" & Octenttohex(oNode8.nodeTypedValue) & """>"  & oNode2.text & "</a>&nbsp;</td>" & vbcrlf
		trReportBody = trReportBody & "<td align=""center"" width=""10%"">" & oNode5.text & "&nbsp;</td>" & vbcrlf
		trReportBody = trReportBody & "<td align=""center"" width=""10%"">" & soOrgnizer & "&nbsp;</td>" & vbcrlf
		trReportBody = trReportBody & "<td align=""center"" width=""10%"">" & GetBusyStatusText(oNode9.text) & "&nbsp;</td>" & vbcrlf
		trReportBody = trReportBody & "<td align=""center"" width=""10%"">" & ncNewclients & "&nbsp;</td>" & vbcrlf
		trReportBody  = trReportBody  & "</tr>" & vbcrlf

   Next
Else
   wscript.echo "Status: " & req.status
   wscript.echo "Status text: " & req.statustext
   wscript.echo "Response text: " & req.responsetext
End If

End Sub

Function Octenttohex(OctenArry) 
ReDim aOut(UBound(OctenArry)) 
For i = 1 to UBound(OctenArry) + 1 
	if len(hex(ascb(midb(OctenArry,i,1)))) = 1 then 
		aOut(i-1) = "0" & hex(ascb(midb(OctenArry,i,1)))
	else
		aOut(i-1) = hex(ascb(midb(OctenArry,i,1)))
	end if
Next 
Octenttohex = join(aOUt,"")
End Function 

Function GetBusyStatusText(bsBusyStatusProp)

select case bsBusyStatusProp
	    case 0 	GetBusyStatusText = "Free"
	    case 1 	GetBusyStatusText = "Tentative"
	    case 2 	GetBusyStatusText = "Busy"
	    case 3 	GetBusyStatusText = "Out of Office"
	    Case Else GetBusyStatusText = "Unknown"
end Select

End Function