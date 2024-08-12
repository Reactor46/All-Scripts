on error resume next
Set Rs = CreateObject("ADODB.Recordset")
set Rec = CreateObject("ADODB.Record")
set contobj = createobject("CDO.Person")
Set Conn = CreateObject("ADODB.Connection")
contactfolderurl = "file://./backofficestorage/domain.com/mbx/sourcemailbox/contacts/"
Conn.Provider = "ExOLEDB.DataSource"
Rec.Open contactfolderurl, ,3
SSql = "Select ""DAV:href"", ""http://schemas.microsoft.com/exchange/permanenturl"" "
SSql = SSql & " FROM scope('shallow traversal of """ & contactfolderurl & """') " 
SSql = SSql & "WHERE ""DAV:contentclass"" = 'urn:content-classes:person' and ""DAV:ishidden"" = false"
Rs.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
Rs.CursorType = 3
rs.open SSql, rec.ActiveConnection, 3
if Rs.recordcount <> 0 then 
Rs.movefirst
while not rs.eof
	set contobj1 = createobject("CDO.Person")
	wscript.echo  rs.fields("DAV:href")
	ourl = rs.fields("http://schemas.microsoft.com/exchange/permanenturl")
	contobj.datasource.open ourl,,3
	set stm = contobj.getvcardstream()
	set stm1 = contobj1.getvcardstream()
	stm1.writetext = stm.readtext
	stm1.flush
	contobj1.fields("urn:schemas:contacts:fileas") = contobj.fields("urn:schemas:contacts:fileas")
	contobj1.fields("http://schemas.microsoft.com/mapi//id/{00062004-0000-0000-C000-000000000046}/0x8080") = contobj.fields("http://schemas.microsoft.com/mapi//id/{00062004-0000-0000-C000-000000000046}/0x8080")	
	contobj1.fields("http://schemas.microsoft.com/mapi//id/{00062004-0000-0000-C000-000000000046}/0x818A") = contobj.fields("http://schemas.microsoft.com/mapi//id/{00062004-0000-0000-C000-000000000046}/0x818A")
	contobj1.fields("http://schemas.microsoft.com/mapi//id/{00062004-0000-0000-C000-000000000046}/0x818B") = contobj.fields("http://schemas.microsoft.com/mapi//id/{00062004-0000-0000-C000-000000000046}/0x818B")
        contobj1.fields("http://schemas.microsoft.com/mapi//id/{00062004-0000-0000-C000-000000000046}/0x818C") = contobj.fields("http://schemas.microsoft.com/mapi//id/{00062004-0000-0000-C000-000000000046}/0x818C")
        contobj1.fields("http://schemas.microsoft.com/mapi//id/{00062004-0000-0000-C000-000000000046}/0x818D") = contobj.fields("http://schemas.microsoft.com/mapi//id/{00062004-0000-0000-C000-000000000046}/0x818D")
	contobj1.fields("http://schemas.microsoft.com/mapi/proptag/0x3A15001E") = contobj.fields("http://schemas.microsoft.com/mapi/proptag/0x3A15001E")
        contobj1.fields("http://schemas.microsoft.com/mapi/proptag/0x3A2A001E") = contobj.fields("http://schemas.microsoft.com/mapi/proptag/0x3A2A001E")
        contobj1.fields("http://schemas.microsoft.com/mapi/proptag/0x3A2B001E") = contobj.fields("http://schemas.microsoft.com/mapi/proptag/0x3A2B001E")
	contobj1.fields.update
	contobj1.datasource.savetocontainer "file://./backofficestorage/domain.com/mbx/targetmailbox/contacts/"
	wscript.echo err.description
	err.clear 
	set stm = nothing
	set stm1 = nothing
	set contobj1 = nothing
	rs.movenext
wend
end if