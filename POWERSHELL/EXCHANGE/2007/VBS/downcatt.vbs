email = WScript.Arguments(0)
Set Person = CreateObject("CDO.Person")
strURL = "mailto:" & email
Person.DataSource.Open strURL
Set Mailbox = Person.GetInterface("IMailbox")
WScript.Echo "URL to inbox for " & email & " is: " & Mailbox.contacts
Set Rs = CreateObject("ADODB.Recordset")
set Rec = CreateObject("ADODB.Record")
Set Conn = CreateObject("ADODB.Connection")
Conn.Provider = "ExOLEDB.DataSource"
Rec.Open Mailbox.contacts, ,3
SSql = "SELECT ""DAV:href"", ""DAV:contentclass"" FROM scope('shallow traversal of """ & Mailbox.contacts & """') " 
SSql = SSql & " Where ""DAV:isfolder"" = false AND ""urn:schemas:httpmail:hasattachment"" = true AND ""DAV:ishidden"" = false "  
Rs.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
Rs.CursorType = 3
i = 1
rs.open SSql, rec.ActiveConnection, 3
if Rs.recordcount <> 0 then 
Rs.movefirst
while not rs.eof
	wscript.echo Rs.Fields("DAV:href").Value
	call procmail(Rs.Fields("DAV:href").Value)
	rs.movenext
wend
end if
rs.close


sub procmail(murl)

set msg = createobject("cdo.message") 
msg.datasource.open murl
set objattachments = msg.attachments 
for each objattachment in objattachments 
if objAttachment.ContentMediaType = "message/rfc822" then
	set msg1 = createobject("cdo.message") 
	msg1.datasource.OpenObject objattachment, "ibodypart"
	Randomize   ' Initialize random-number generator.
	rndval = Int((20000000000 * Rnd) + 1)   
	fnFileName = "c:\temp\attachmsg" &  day(now) & month(now) & year(now) & hour(now) & minute(now) & rndval & ".eml"
	set stm = msg1.getstream
	stm.savetofile fnFileName 
else 
	fnFileName = "c:\temp\" & objattachment.filename
	objAttachment.savetofile fnFileName
end if
next 
set msg = nothing 


end sub
