servername = "ws03r2eeexchlcs"

smSoapMessage  = "<?xml version='1.0' encoding='utf-8'?>" _
& "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " _
& " xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""" _
& " xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types"" >" _
& "<soap:Body>" _
& "<GetFolder xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages"" " _
& " xmlns:t=""http://schemas.microsoft.com/exchange/services/2006/types""> " _
& "<FolderShape> " _
& "<t:BaseShape>Default</t:BaseShape> " _
& "</FolderShape> " _
& "<FolderIds> " _
& "<t:DistinguishedFolderId Id=""inbox""/> " _
& "</FolderIds> " _
& "</GetFolder> " _
& "</soap:Body></soap:Envelope>"

set req = createobject("microsoft.xmlhttp")
req.Open "post", "http://" & servername & "/ews/Exchange.asmx", False,"CONTOSO\Administrator", "Evaluation1" 
req.setRequestHeader "Content-Type", "text/xml"
req.setRequestHeader "translate", "F"
req.send smSoapMessage
wscript.echo "Request Status : " & req.status
wscript.echo 
Set oXMLDoc = req.responseXML
Set oXMLUnreadNodes = oXMLDoc.getElementsByTagName("t:UnreadCount")
For i = 0 To (oXMLUnreadNodes.length -1)
	set oNode = oXMLUnreadNodes.nextNode
	wscript.echo "Number of Unread Email : " & oNode.text
next
