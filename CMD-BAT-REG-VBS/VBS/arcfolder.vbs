mailboxurl = "file://./backofficestorage/yourdomain.com/MBX/mailbox/inbox"
set Rec = CreateObject("ADODB.Record")
set Rs = CreateObject("ADODB.Recordset")
Set Conn = CreateObject("ADODB.Connection")
Conn.Provider = "ExOLEDB.DataSource"
Rec.Open mailboxurl, ,3
SSql = "SELECT ""DAV:href"", ""DAV:contentclass"" FROM scope('shallow traversal of """ & mailboxurl & """') " 
SSql = SSql & " WHERE (""urn:schemas:httpmail:datereceived"" < CAST(""" & isodateit(now()-31) & """ as 'dateTime')) AND ""DAV:isfolder"" = false"                 
SSql = SSql & " AND ""DAV:contentclass"" = 'urn:content-classes:message'"
Rs.CursorLocation = 2 'adUseServer = 2, adUseClient = 3
rs.open SSql, rec.ActiveConnection, 3
while not rs.eof
	rs.delete 1
	rs.movenext
wend
rs.close


function isodateit(datetocon)
	strDateTime = year(datetocon) & "-"
	if (Month(datetocon) < 10) then strDateTime = strDateTime & "0"
	strDateTime = strDateTime & Month(datetocon) & "-"
	if (Day(datetocon) < 10) then strDateTime = strDateTime & "0"
	strDateTime = strDateTime & Day(datetocon) & "T" & formatdatetime(datetocon,4) & ":00Z"
	isodateit = strDateTime
end function