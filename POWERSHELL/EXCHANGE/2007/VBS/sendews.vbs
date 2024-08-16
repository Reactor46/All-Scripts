emailaddress = "Administrator@contoso.com"
stSubjet = "hello World"
ebEmailBody = "1234567910 Blast Off"
servername = "ws03r2eeexchlcs"

smSoapMessage  = "<?xml version='1.0' encoding='utf-8'?>" _
& "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " _
& " xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">" _
& "<soap:Body><CreateItem MessageDisposition=""SendAndSaveCopy"" " _
& " xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""> " _
& "<SavedItemFolderId><DistinguishedFolderId Id=""sentitems"" xmlns=""http://schemas.microsoft.com/exchange/services/2006/types""/>" _
& "</SavedItemFolderId><Items><Message xmlns=""http://schemas.microsoft.com/exchange/services/2006/types"">" _
& "<ItemClass>IPM.Note</ItemClass><Subject>" & stSubjet & "</Subject><Body BodyType=""Text"">" & ebEmailBody &  "</Body><ToRecipients>" _
& "<Mailbox><EmailAddress>" & EmailAddress & "</EmailAddress></Mailbox></ToRecipients></Message></Items>" _
& "</CreateItem></soap:Body></soap:Envelope>"
set req = createobject("microsoft.xmlhttp")

req.Open "post", "http://" & servername & "/ews/Exchange.asmx", False,"CONTOSO\Administrator", "Evaluation1" 
req.setRequestHeader "Content-Type", "text/xml"
req.setRequestHeader "translate", "F"
req.send smSoapMessage
wscript.echo req.status
wscript.echo 
wscript.echo req.responsetext