emailaddress = "user@domain.com"
stSubjet = "hello World"
ebEmailBody = "1234567910 Blast Off"
servername = "servername"
unUsername = "domain\username"
pwPassWord = "password"

set msMessageObject = createobject("CDO.Message")
msMessageObject.Subject = stSubjet 
msMessageObject.TextBody = ebEmailBody
msMessageObject.AddAttachment "c:\file.ext"
set stm = msMessageObject.getstream
stm.Type = 1
set convobj = CreateObject("Msxml2.DOMDocument.4.0")
Set oRoot = convobj.createElement("test")
oRoot.dataType = "bin.base64"
oRoot.nodeTypedValue = stm.Read


smSoapMessage  = "<?xml version='1.0' encoding='utf-8'?>" _
& "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " _
& " xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">" _
& "<soap:Body><CreateItem MessageDisposition=""SendAndSaveCopy"" " _
& " xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""> " _
& "<SavedItemFolderId><DistinguishedFolderId Id=""sentitems"" xmlns=""http://schemas.microsoft.com/exchange/services/2006/types""/>" _
& "</SavedItemFolderId><Items><Message xmlns=""http://schemas.microsoft.com/exchange/services/2006/types"">" _
& "<MimeContent>" & oRoot.text &  "</MimeContent><ToRecipients>" _
& "<Mailbox><EmailAddress>" & EmailAddress & "</EmailAddress></Mailbox></ToRecipients></Message></Items>" _
& "</CreateItem></soap:Body></soap:Envelope>"
set req = createobject("microsoft.xmlhttp")

req.Open "post", "https://" & servername & "/ews/Exchange.asmx", False,unUsername, pwPassWord 
req.setRequestHeader "Content-Type", "text/xml"
req.setRequestHeader "translate", "F"
req.send smSoapMessage
wscript.echo req.status
wscript.echo 
wscript.echo req.responsetext

