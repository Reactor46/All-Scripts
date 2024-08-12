Set iCalMsg = CreateObject("CDO.Message")
iCalMsg.datasource.open "http://server/exchange/mailbox/inbox/calendarmessage.eml"
recplist = iCalMsg.fields("http://schemas.microsoft.com/mapi/proptag/0x8167001E")
recparray = split(recplist,";",-1,1) 
for i = lbound(recparray) to ubound(recparray) 
    if instr(iCalMsg.fields("http://schemas.microsoft.com/mapi/proptag/0x0E04001E"),recparray(I)) then 
    else 
    if instr(iCalMsg.fields("http://schemas.microsoft.com/mapi/proptag/0x0E03001E"),recparray(I)) then 
    else 
        wscript.echo recparray(I)
    end if
    end if 
next