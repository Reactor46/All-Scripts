Set ObjxmlHttp = CreateObject("Microsoft.XMLHTTP")
ObjxmlHttp.Open "GET","http://server/public/test2/dlname.EML?Cmd=viewmembers", False, "domain\user", "pass"
ObjxmlHttp.Send
set oResponseDoc = ObjxmlHttp.responseXML
set oNodeList = oResponseDoc.getElementsByTagName("memberid")
set oNodeList1 = oResponseDoc.getElementsByTagName("email")
For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	set oNode1 = oNodeList1.nextNode
	wscript.echo oNode.text & " " & oNode1.text
next
