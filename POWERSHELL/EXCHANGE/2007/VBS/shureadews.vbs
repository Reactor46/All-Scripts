set shell = createobject("wscript.shell")
strValueName = "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
minTimeOffset = shell.regread(strValueName)
toffset = datediff("h",DateAdd("n", minTimeOffset, now()),now())

servername = "ws03r2eeexchlcs"

smSoapMessage  = "<?xml version='1.0' encoding='utf-8'?>" _
& "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " _
& " xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""" _
& " xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"" >" _
& "<soap:Body>" _
& "<FindItem xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages"" " _
& "xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"" Traversal=""Shallow""> " _
& "<ItemShape>" _
& "<t:BaseShape>AllProperties</t:BaseShape>" _
& "</ItemShape>" _
& "<Restriction>" _
& "<t:IsEqualTo>" _
& "<t:FieldURI FieldURI=""message:IsRead""/>" _ 
& "<t:FieldURIOrConstant>" _
& "<t:Constant Value=""0""/>" _
& "</t:FieldURIOrConstant>" _
& "</t:IsEqualTo>" _
& "</Restriction>" _
& "<ParentFolderIds>" _
& "<t:DistinguishedFolderId Id=""inbox""/>" _
& "</ParentFolderIds>" _
& "</FindItem>" _
& "</soap:Body></soap:Envelope>"

set req = createobject("microsoft.xmlhttp")
req.Open "post", "http://" & servername & "/ews/Exchange.asmx", False,"CONTOSO\Administrator", "Evaluation1" 
req.setRequestHeader "Content-Type", "text/xml"
req.setRequestHeader "translate", "F"
req.send smSoapMessage
wscript.echo "Request Status : " & req.status
Wscript.echo "Recieved	From	Subject	Size"
Set oXMLDoc = req.responseXML
Set oXMLUnreadNodes = oXMLDoc.getElementsByTagName("t:Subject")
Set oXMLFromNodes = oXMLDoc.getElementsByTagName("t:From")
Set oXMLTSentNodes = oXMLDoc.getElementsByTagName("t:DateTimeSent")
Set oXMLTSizeNodes = oXMLDoc.getElementsByTagName("t:Size")
For i = 0 To (oXMLUnreadNodes.length -1)
	set oNode = oXMLUnreadNodes.nextNode
	set oNode1 = oXMLFromNodes.nextNode
	set oNode2 = oXMLTSentNodes.nextNode
	set oNode3 = oXMLTSizeNodes.nextNode
	for each emnode in oNode1.childNodes
		From = emnode.text
	next
	wscript.echo dateadd("h",toffset,formatdatetime(replace(replace(oNode2.text,"T"," "),"Z",""))) & "	" & From & "	" & oNode.text & "	" & oNode3.text
next
