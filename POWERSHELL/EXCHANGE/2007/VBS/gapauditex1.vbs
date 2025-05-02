on error resume next
servername = wscript.arguments(0)
domainname = "yourdomain.com"

set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())
set conn1 = createobject("ADODB.Connection")
Set fso = CreateObject("Scripting.FileSystemObject")
fname = "c:\" & wscript.arguments(0) & ".csv"
set wfile = fso.opentextfile(fname,2,true)
wfile.writeline("User,Subject,UTC timestart,UTC timeend, InstanceType, OutlookTimezone,CDOTimezoneEnumID")
public datefrom
public dateto
datefrom = "2006-03-26T10:00:00Z"
dateto = "2006-04-02T10:00:00Z"
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
	strQuery = "<LDAP://"  & strDefaultNamingContext & ">;" & GALQueryFilter & ";distinguishedName,mail,displayname,mailnickname;subtree"
	com.Properties("Page Size") = 100
	Com.CommandText = strQuery
	Set Rs1 = Com.Execute
	while not Rs1.eof
		wscript.echo "User: " & rs1.fields("displayname")
		user = rs1.fields("mail")
		sConnString = "file://./backofficestorage/" & domainname
		sConnString = sConnString & "/mbx/" & user & "/calendar"
		WScript.Echo sConnString
		call QueryCalendarFolder(sConnString,user)		
		rs1.movenext
	wend
	rs.movenext
wend
rs.close
set conn = nothing
set com = nothing
wscript.echo "Done"


Public Sub QueryCalendarFolder(sConnString,user)
SSql = "SELECT ""DAV:href"", ""urn:schemas:calendar:timezoneid"", ""urn:schemas:httpmail:subject"", "
SSql = SSql & """http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8234"", "
SSql = SSql & """http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x822E"", "
SSql = SSql & """urn:schemas:calendar:instancetype"", ""urn:schemas:calendar:dtstart"" , ""urn:schemas:calendar:dtend"" "
SSql = SSql & "FROM scope('shallow traversal of """ & sConnString & """') " 
SSql = SSql & " Where ""DAV:isfolder"" = false AND ""DAV:ishidden"" = false "           
SSql = SSql & "AND ""urn:schemas:calendar:dtend"" > CAST(""" & datefrom & """ as 'dateTime') " _
& "AND ""urn:schemas:calendar:dtstart"" < CAST(""" & dateto & """ as 'dateTime')"  
   Set oConn = CreateObject("ADODB.Connection")
   oConn.Provider = "Exoledb.DataSource"
   oConn.Open sConnString
   Set oRecSet = CreateObject("ADODB.Recordset")
   oRecSet.CursorLocation = 3
   oRecSet.Open sSQL, oConn.ConnectionString
   if err.number <> 0 then wfile.writeline(user & "," & "Error Connection to Mailbox")
   While oRecSet.EOF <> True
      Wscript.echo User
      wscript.echo oRecSet.fields("DAV:Href").value
      wscript.echo oRecSet.fields("urn:schemas:httpmail:subject").value
      wscript.echo oRecSet.fields("http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8234").value
      wscript.echo oRecSet.fields("urn:schemas:calendar:timezoneid").value
      wscript.echo oRecSet.fields("urn:schemas:calendar:instancetype").value
      set aptobj = createobject("CDO.Appointment")
      organiser = ""
      aptobj.datasource.open oRecSet.fields("DAV:Href").value
      for each attendee in aptobj.attendees
      	if attendee.isorganizer = True then organiser = attendee.address
      next
      wscript.echo 
      wscript.echo 
      wfile.writeline(user & "," & oRecSet.fields("urn:schemas:httpmail:subject").value & ","_
     & oRecSet.fields("urn:schemas:calendar:dtstart").value & "," & oRecSet.fields("urn:schemas:calendar:dtend").value _
     & "," & oRecSet.fields("urn:schemas:calendar:instancetype").value  & "," _
     & replace(replace(oRecSet.fields("http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8234").value,vbcrlf,""),",","") _
     & "," & organiser)
      oRecSet.MoveNext
   wend
   oRecSet.Close
   oConn.Close
   Set oRecSet = Nothing
   Set oConn = Nothing
End Sub


