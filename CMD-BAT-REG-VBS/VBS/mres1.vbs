set WshShell = CreateObject("WScript.Shell")
Set obArgs = WScript.Arguments
tmailbox = obArgs.Item(0)
Set Rs = CreateObject("ADODB.Recordset")
Set fso = CreateObject("Scripting.FileSystemObject")
Set msgobj = CreateObject("CDO.Message")
tyear = year(now()-60)
tmonth = month(now()-60)
if tmonth < 10 then tmonth = 0 & tmonth
stday = day(now()-60)
if stday < 10 then stday = 0 & stday
sttime = formatdatetime(now1,4)
qdatest = tyear & "-" & tmonth & "-" & stday & "T"
qdatest1 = qdatest & sttime & ":" & "00Z"
set Rec = CreateObject("ADODB.Record")
set Rec1 = CreateObject("ADODB.Record")
Set Conn = CreateObject("ADODB.Connection")
mailboxurl = "file://./backofficestorage/yourdomain.com.au/MBX/" & tmailbox & "/"
Conn.Provider = "ExOLEDB.DataSource"
Rec.Open mailboxurl, ,3
mailboxurl = "file://./backofficestorage/yourdomain.com.au/MBX/" & tmailbox & "/inbox/"
SSql = "SELECT ""DAV:href"", ""DAV:contentclass"" FROM scope('shallow traversal of """ & mailboxurl & """') " 
SSql = SSql & " WHERE (""urn:schemas:httpmail:datereceived"" < CAST(""" & qdatest1 & """ as 'dateTime')) AND ""DAV:isfolder"" = false AND ""urn:schemas:httpmail:read"" = false"                 
Rs.CursorLocation = 2 'adUseServer = 2, adUseClient = 3
Rs.CursorType = 3
rs.open SSql, rec.ActiveConnection, 3
if Rs.recordcount <> 0 then 
Rs.movefirst
while not rs.eof
	rs.delete 1
	rs.movenext
wend
end if
rs.close
wscript.echo "done"

