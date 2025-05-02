mailbox = "mailbox"
Set Rs = CreateObject("ADODB.Recordset")
Set msgobj = CreateObject("CDO.Message")
set Rec = CreateObject("ADODB.Record")
set Rec1 = CreateObject("ADODB.Record")
Set Conn = CreateObject("ADODB.Connection")
mailboxurl = "file://./backofficestorage/yourdomain.com/MBX/" & mailbox & "/"
Conn.Provider = "ExOLEDB.DataSource"
Rec.Open mailboxurl, ,3
mailboxurl = "file://./backofficestorage/yourdomain.com/MBX/" & mailbox & "/inbox/"
SSql = "SELECT ""DAV:href"", ""DAV:uid"", ""DAV:contentclass"" FROM scope('shallow traversal of """ & mailboxurl & """') " 
SSql = SSql & " WHERE ""http://schemas.microsoft.com/mapi/proptag/0x65EB001E"" = 'JunkEmailRule' and ""http://schemas.microsoft.com/exchange/outlookmessageclass"" = 'IPM.ExtendedRule.Message' "             
Rs.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
Rs.CursorType = 3
rs.open SSql, rec.ActiveConnection, 3
if Rs.recordcount <> 0 then 
Rs.movefirst
while not rs.eof
	wscript.echo Rs.Fields("DAV:href").Value
	rec1.open Rs.Fields("DAV:href").Value,,3
	wscript.echo rec1.fields("http://schemas.microsoft.com/mapi/proptag/0x61010003").Value
	wscript.echo rec1.fields("http://schemas.microsoft.com/mapi/proptag/0x61020003").Value
	rec1.fields("http://schemas.microsoft.com/mapi/proptag/0x61010003").Value = 3
	rec1.fields("http://schemas.microsoft.com/mapi/proptag/0x61020003").Value = 0
	rec1.fields.update
	rec1.close
	rs.movenext
wend
end if
rs.close

