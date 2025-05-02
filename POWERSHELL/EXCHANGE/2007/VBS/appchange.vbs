<SCRIPT LANGUAGE="VBScript">

Sub ExStoreEvents_OnSave(pEventInfo, bstrURLItem, lFlags)

Const EVT_NEW_ITEM = 1
Const EVT_IS_DELIVERED = 8

If (lFlags And EVT_IS_DELIVERED) Or (lFlags And EVT_NEW_ITEM) Then

chgappt = 0
LocalSearchdomain = "@yourdomain.com"
set apptobj = createobject("CDO.Appointment")
apptobj.datasource.open bstrURLItem,,3
cval = apptobj.fields("http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8214")
if apptobj.Attendees.count > 1 then
for each attend in apptobj.Attendees
  if instr(lcase(attend.address),LocalSearchdomain) = 0 then
	chgappt = 1
  end if
next
if chgappt = 1 then
	if cavl <> 2 then 
		apptobj.fields("http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8214") = clng(2)
		apptobj.fields.update
		apptobj.datasource.save
	end if
else
	if cavl <> 3 then 
		apptobj.fields("http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x8214") = clng(3)
		apptobj.fields.update
		apptobj.datasource.save	
	end if
end if
end if
apptobj = nothing

end if

End Sub

</SCRIPT>