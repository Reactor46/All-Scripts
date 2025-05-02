xmlstr = ""
xmlstr = xmlstr & "Cmd=options" & vbLf
xmlstr = xmlstr & "junkemailstate=1" & vbLf
xmlstr = xmlstr & "cmd=savejunkemailrule" & vbLf
Set ObjxmlHttp = CreateObject("Microsoft.XMLHTTP")
ObjxmlHttp.Open "POST", "http://server/exchange/mailbox/", False, "domain\user", "password"
ObjxmlHttp.setRequestHeader "Accept-Language:","en-us"
ObjxmlHttp.setRequestHeader "Content-type:","application/x-www-UTF8-encoded"
ObjxmlHttp.setRequestHeader "Content-Length:", Len(xmlstr)
ObjxmlHttp.Send xmlstr
Wscript.echo ObjxmlHttp.responseText
