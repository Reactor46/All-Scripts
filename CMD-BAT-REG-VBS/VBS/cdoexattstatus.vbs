set apptobj = createobject("CDO.Appointment")
apptobj.datasource.open "file://./backofficestorage/domain.com/MBX/mailbox/calendar/appoitment.EML"
for each attend in apptobj.Attendees
 wscript.echo attend.address
 wscript.echo attend.role
 wscript.echo attend.status
next
