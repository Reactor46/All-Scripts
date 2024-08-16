calurl = "file://./backofficestorage/youdomain.com/public folders/calenderfoldername"
call updateappointment(calurl,0)
wscript.echo
wscript.echo "Reccuring Appointments"
wscript.echo
call updateappointment(calurl,1)

sub updateappointment(CalendarURL,instancetype)

set rec = createobject("ADODB.Record")
set rec1 = createobject("ADODB.Record")
set rs = createobject("ADODB.Recordset")
Rec.Open CalendarURL
Set Rs.ActiveConnection = Rec.ActiveConnection
Rs.Source = "SELECT ""DAV:href"", " & _
" ""urn:schemas:calendar:location"", " & _
" ""urn:schemas:calendar:instancetype"", " & _
" ""urn:schemas:calendar:dtstart"", " & _
" ""urn:schemas:calendar:dtend"", " & _
" ""http://schemas.microsoft.com/mapi/proptag/0x3FF8001E"" " & _
"FROM scope('shallow traversal of """ & CalendarURL & """') " & _
"WHERE (""urn:schemas:calendar:dtstart"" >= CAST(""2005-12-01T08:00:00Z"" as 'dateTime')) " & _
"AND (""urn:schemas:calendar:dtend"" <= CAST(""2006-12-01T08:00:00Z"" as 'dateTime'))" & _
" AND (""urn:schemas:calendar:instancetype"" = " & instancetype  & ")"
Rs.Open
If Not (Rs.EOF) Then
Rs.MoveFirst
Do Until Rs.EOF
	ItemURL = Rs.Fields("DAV:Href").Value
	wscript.echo ItemURL
	creator = " Created by " &  Rs.Fields("http://schemas.microsoft.com/mapi/proptag/0x3FF8001E").Value 
	wscript.echo creator
	if instr(rs.fields("urn:schemas:calendar:location"),"Created by") then
		wscript.echo "Creator Exists"
	else
		rec1.open Rs.Fields("DAV:Href").Value,,3
		rec1.fields("urn:schemas:calendar:location") = rec1.fields("urn:schemas:calendar:location") & creator 
		rec1.fields.update 
		rec1.close
		wscript.echo "Added Creator"
	end if 
	rs.movenext
loop
end if

end sub

