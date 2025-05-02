$ebEmailBody = "Hello World"
$EmailAddress = "Administrator@contoso.com"
$servername = "ws03r2eeexchlcs"
$stSubjet = "Test Email subject"
$cdUsrCredentials = new-object System.Net.NetworkCredential("Administrator", "Evaluation1", "CONTOSO");

$smSoapMessage  = "<?xml version='1.0' encoding='utf-8'?>" `
+ "<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" " `
+ " xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">" `
+ "<soap:Body><CreateItem MessageDisposition=""SendAndSaveCopy"" " `
+ " xmlns=""http://schemas.microsoft.com/exchange/services/2006/messages""> " `
+ "<SavedItemFolderId><DistinguishedFolderId Id=""sentitems"" xmlns=""http://schemas.microsoft.com/exchange/services/2006/types""/>" `
+ "</SavedItemFolderId><Items><Message xmlns=""http://schemas.microsoft.com/exchange/services/2006/types"">" `
+ "<ItemClass>IPM.Note</ItemClass><Subject>" + $stSubjet + "</Subject><Body BodyType=""Text"">" + $ebEmailBody +  "</Body><ToRecipients>" `
+ "<Mailbox><EmailAddress>" + $EmailAddress + "</EmailAddress></Mailbox></ToRecipients></Message></Items>" `
+ "</CreateItem></soap:Body></soap:Envelope>"


$servername = "ws03r2eeexchlcs"
$strRootURI = "http://" + $servername + "/ews/Exchange.asmx"
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
