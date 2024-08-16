usertodel = "user@domain.com"
Set ObjxmlHttp = CreateObject("Microsoft.XMLHTTP")
ObjxmlHttp.Open "GET","http://server/public/test2/dlname.EML?Cmd=viewmembers", False, "domain\user", "pass"
ObjxmlHttp.Send
set oResponseDoc = ObjxmlHttp.responseXML
set oNodeList = oResponseDoc.getElementsByTagName("memberid")
set oNodeList1 = oResponseDoc.getElementsByTagName("email")
For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	set oNode1 = oNodeList1.nextNode
	if oNode1.text = usertodel then 
		delmember(oNode.text)
		wscript.echo "Member Deleted " & oNode1.text
	end if
next

function delmember(utodel)
Set ObjxmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
xmlstr = ""
xmlstr = xmlstr & "Cmd=deletemember" & vbLf
xmlstr = xmlstr & "msgclass=IPM.DistList" & vbLf
xmlstr = xmlstr & "memberid=" & utodel
ObjxmlHttp.Open "POST", "http://server/public/test2/dlname.EML", false, "domain\user", "pass"
ObjxmlHttp.setRequestHeader "Accept-Language", "en-us"
ObjxmlHttp.setRequestHeader "Content-type", "application/x-www-UTF8-encoded"
ObjxmlHttp.setRequestHeader "Content-Length", Len(xmlstr)
ObjxmlHttp.Send xmlstr
end function
