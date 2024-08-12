set req = Createobject("microsoft.xmlhttp")

CalendarURL = "https://servername/exchange/mailbox/calendar"
AppointmentURL = CalendarURL & "/appointment.eml"

If RecurringAppointment = false Then 
	xmlstr = "<?xml version=""1.0""?><a:propertyupdate xmlns:a=""DAV:""><a:remove><a:prop>" _
	& "<d:reminderoffset xmlns:d=""urn:schemas:calendar:"" /></a:prop></a:remove>" _
	& "<a:target><a:href>"  & AppointmentURL & "</a:href></a:target></a:propertyupdate>"
else
	xmlstr = "<?xml version=""1.0""?><a:propertyupdate xmlns:a=""DAV:""><a:set><a:prop>" _
	& "<d:remindernexttime xmlns:d=""urn:schemas:calendar:"">4501-01-01T00:00:00.000Z</d:remindernexttime>" _
	& "</a:prop></a:set><a:target><a:href>" & AppointmentURL & "</a:href></a:target></a:propertyupdate>"
End if

req.open "BPROPPATCH", CalendarURL, False
req.setRequestHeader "Content-Type", "text/xml;"
req.setRequestHeader "Translate", "f"
req.setRequestHeader "Content-Length:", Len(xmlstr)
req.send(xmlstr)
If (req.Status >= 200 And req.Status < 300) Then
 	Wscript.echo "Reminder Acknowledged" 
else
	Wscript.echo "Request Failed. Results = " & req.Status & ": " &  req.statusText
End If 
