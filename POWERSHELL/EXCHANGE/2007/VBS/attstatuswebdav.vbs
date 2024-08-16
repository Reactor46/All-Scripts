set Req = createobject("Microsoft.XMLHTTP")
Req.open "GET", "http://server/exchange/mailbox/Calendar/appoinment.EML", false
Req.setRequestHeader "Translate","f"
Req.send
attendeearry = split(req.responsetext,"ATTENDEE;",-1,1)
for i = 1 to ubound(attendeearry)
string1 = vbcrlf & " "
stparse = replace(attendeearry(i),string1,"")
attaddress = mid(stparse,(instr(stparse,"MAILTO:")+7),instr(stparse,chr(13)))
attaddress = mid(attaddress,1,instr(attaddress,vbcrlf))
attrole = mid(stparse,(instr(stparse,"ROLE=")+5),instr((instr(stparse,"ROLE=")+5),stparse,";")-(instr(stparse,"ROLE=")+5))
attstatus = mid(stparse,(instr(stparse,"PARTSTAT=")+9),instr((instr(stparse,"PARTSTAT=")+9),stparse,";")-(instr(stparse,"PARTSTAT=")+9))
wscript.echo attaddress
wscript.echo attrole
wscript.echo attstatus
next
