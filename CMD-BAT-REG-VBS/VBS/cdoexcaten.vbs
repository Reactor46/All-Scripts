Set iCalMsg = CreateObject("CDO.CalendarMessage")
iCalMsg.datasource.open "http://server/exchange/mailbox/Inbox/calmessage.EML"
For Each iCalPart In iCalMsg.CalendarParts
      Set iAppt = iCalPart.GetUpdatedItem
	cuid1 = iAppt.fields("urn:schemas:calendar:uid")	
	for each attend in iAppt.Attendees
		if attend.IsOrganizer <> 0 then
			Set Person = CreateObject("CDO.Person")
			strURL = attend.address
			Person.DataSource.Open strURL
			Set Mailbox = Person.GetInterface("IMailbox")
			set iAppt1 = iCalPart.GetAssociatedItem(Mailbox.calendar)
			for each attend1 in iAppt1.Attendees
				 wscript.echo attend1.address
 				 wscript.echo attend1.role
 				 wscript.echo attend1.status
				 wscript.echo attend1.type
			next
		end if
	next
Next 


