$EmailAddress = "user@domain.com"
$servername = "servername"
$ebEmailBody = "Hello World"
$stSubjet = "Test Email subject"
$cdUsrCredentials = new-object System.Net.NetworkCredential("username", "password", "domain");
$msMessage = New-Object -comobject  CDO.Message
$msMessage.Subject = "Test Email subject"
$msMessage.TextBody = $ebEmailBody
$msMessage.AddAttachment("c:\file.ext",$null,$null);
$stStream = $msMessage.getStream()
$stStream.Type = 1
$binaryData1 = new-object byte[] $stStream.Size
$binaryData1 = $stStream.Read($stStream.Size)


$smSoapMessage  = "<?xml version='1.0' encoding='utf-8'?>" `
+ "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " `
+ " xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">" `
+ "<soap:Body><CreateItem MessageDisposition=""SendAndSaveCopy"" " `
+ " xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""> " `
+ "<SavedItemFolderId><DistinguishedFolderId Id=""sentitems"" xmlns=""http://schemas.microsoft.com/exchange/services/2006/types""/>" `
+ "</SavedItemFolderId><Items><Message xmlns=""http://schemas.microsoft.com/exchange/services/2006/types"">" `
+ "<MimeContent>" + [System.Convert]::ToBase64String($binaryData1) +  "</MimeContent><ToRecipients>" `
+ "<Mailbox><EmailAddress>" + $EmailAddress + "</EmailAddress></Mailbox></ToRecipients></Message></Items>" `
+ "</CreateItem></soap:Body></soap:Envelope>"


$strRootURI = "https://" + $servername + "/ews/Exchange.asmx"
$WDRequest = [System.Net.WebRequest]::Create($strRootURI)
$WDRequest.ContentType = "text/xml"
$WDRequest.Headers.Add("Translate", "F")
$WDRequest.Method = "Post"
$WDRequest.Credentials = $cdUsrCredentials
$bytes = [System.Text.Encoding]::UTF8.GetBytes($smSoapMessage)
$WDRequest.ContentLength = $bytes.Length
$RequestStream = $WDRequest.GetRequestStream()
$RequestStream.Write($bytes, 0, $bytes.Length)
$RequestStream.Close()
$WDResponse = $WDRequest.GetResponse()
$ResponseStream = $WDResponse.GetResponseStream()
"Message Sent"
