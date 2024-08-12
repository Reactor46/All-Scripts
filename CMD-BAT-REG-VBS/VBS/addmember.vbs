Set ObjxmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
xmlstr = ""
xmlstr = xmlstr & "Cmd=addmember" & vbLf
xmlstr = xmlstr & "msgclass=IPM.DistList" & vbLf
xmlstr = xmlstr & "member=user@domain.com" & vbLf
ObjxmlHttp.Open "POST", "http://server/public/test2/Dlname.eml", false, "domain\user", "pass"
ObjxmlHttp.setRequestHeader "Accept-Language", "en-us"
ObjxmlHttp.setRequestHeader "Content-type", "application/x-www-UTF8-encoded"
ObjxmlHttp.setRequestHeader "Content-Length", Len(xmlstr)
ObjxmlHttp.Send xmlstr
