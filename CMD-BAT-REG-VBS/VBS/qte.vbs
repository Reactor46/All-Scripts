strComputerName = "servername"
Searchaddress = "address@domain.com"

toffset = datediff("h",convertUTC(now(), 57, 0),now())
dtListFrom = DateAdd("h",datediff("h",now(),convertUTC(now(), 57, 0)),now)
dtListFrom = DateAdd("n",-15,dtListFrom)
strStartDateTime = year(dtListFrom)
if (Month(dtListFrom) < 10) then strStartDateTime = strStartDateTime & "0"
strStartDateTime = strStartDateTime & Month(dtListFrom)
if (Day(dtListFrom) < 10) then strStartDateTime = strStartDateTime & "0"
strStartDateTime = strStartDateTime & Day(dtListFrom)
if (Hour(dtListFrom) < 10) then strStartDateTime = strStartDateTime & "0"
strStartDateTime = strStartDateTime & Hour(dtListFrom)
if (Minute(dtListFrom) < 10) then strStartDateTime = strStartDateTime & "0"
strStartDateTime = strStartDateTime & Minute(dtListFrom)
if (Second(dtListFrom) < 10) then strStartDateTime = strStartDateTime & "0"
strStartDateTime = strStartDateTime & Second(dtListFrom) & ".000000+000"
Const cWMINameSpace = "root/MicrosoftExchangeV2"
Const cWMIInstance = "Exchange_MessageTrackingEntry"



strWinMgmts = "winmgmts:{impersonationLevel=impersonate}!//" & strComputerName & "/" & cWMINameSpace
Set objWMIExchange =  GetObject(strWinMgmts)
Set listExchange_MessageTrackingEntries = objWMIExchange.ExecQuery("Select * FROM Exchange_MessageTrackingEntry where entrytype = '1020' and OriginationTime >= '" & strStartDateTime & "' or entrytype = '1028' and OriginationTime > '" & strStartDateTime & "'")
For each objExchange_MessageTrackingEntry in listExchange_MessageTrackingEntries
	for i = 1 to objExchange_MessageTrackingEntry.RecipientCount
		if objExchange_MessageTrackingEntry.RecipientAddress((i-1)) = Searchaddress then
			wscript.echo objExchange_MessageTrackingEntry.senderaddress
		end if
	next
Next




function condate(date2con)
dtcon = date2con
if month(dtcon) < 10 then 
	if day(dtcon) < 10 then
		qdat = year(dtcon) & "0" & month(dtcon) & "0" & day(dtcon)
	else
		qdat = year(dtcon) & "0" & month(dtcon) & day(dtcon)
	end if 
else
	if day(dtcon) < 10 then
		qdat = year(dtcon) & month(dtcon) & "0" & day(dtcon)
	else
		qdat = year(dtcon) & month(dtcon) & day(dtcon)
	end if 
end if
condate = qdat 
end function 

function convertUTC(dtconv, tzfr, tzTo)
	Set tapptobj = CreateObject("CDO.Appointment")
	Set tapptconf = CreateObject("CDO.Configuration")
	tapptobj.Configuration = tapptconf
	tapptconf.Fields("urn:schemas:calendar:timezoneid").Value = tzfr
	tapptconf.Fields.Update
	tapptobj.StartTime = dtconv
	tapptconf.Fields("urn:schemas:calendar:timezoneid").Value = tzTo
	tapptconf.Fields.Update
	convertutc = tapptobj.StartTime
end function