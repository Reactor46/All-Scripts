stmsgbody ="Hello Fred <BR<BR>"
stmsgbody = stmsgbody & "What to get some coffee Mate, Meet me downstairs in ten<BR><BR>"
stmsgbody = stmsgbody & "Cheers<BR>"
stmsgbody = stmsgbody & "Barney<BR>"
szXml = ""
szXml = szXml & "Cmd=send" & vbLf
szXml = szXml & "MsgTo=address@domain.com" & vbLf
szXml = szXml & "MsgCc=" & vbLf
szXml = szXml & "MsgBcc=" & vbLf
szXml = szXml & "urn:schemas:httpmail:importance=1" & vbLf
szXml = szXml & "http://schemas.microsoft.com/exchange/readreceiptrequested=1" & vbLf
szXml = szXml & "http://schemas.microsoft.com/exchange/deliveryreportrequested=1" & vbLf
szXml = szXml & "http://schemas.microsoft.com/exchange/sensitivity-long=" & vbLf
szXml = szXml & "urn:schemas:httpmail:subject=Coffee ???" & vbLf
szXml = szXml & "urn:schemas:httpmail:htmldescription=<!DOCTYPE HTML PUBLIC " _
& """-//W3C//DTD HTML 4.0 Transitional//EN""><HTML DIR=ltr><HEAD><META HTTP-EQUIV" _
& "=""Content-Type"" CONTENT=""text/html; charset=utf-8""></HEAD><BODY><DIV>" _
& "<FONT face='Arial' color=#000000 size=2>" & stmsgbody & "</font>" _
& "</DIV></BODY></HTML>" & vbLf
Set ObjxmlHttp = CreateObject("Microsoft.XMLHTTP")
ObjxmlHttp.Open "POST", "https://server/exchange/mailbox/Drafts", False, "",""
ObjxmlHttp.setRequestHeader "Accept-Language:", "en-us"
ObjxmlHttp.setRequestHeader "Content-type:", "application/x-www-UTF8-encoded"
ObjxmlHttp.setRequestHeader "Content-Length:", Len(xmlstr)
ObjxmlHttp.Send szXml
Wscript.echo ObjxmlHttp.responseText 