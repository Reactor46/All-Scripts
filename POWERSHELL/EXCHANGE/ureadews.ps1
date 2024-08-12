$servername = "ws03r2eeexchlcs"
$cdUsrCredentials = new-object System.Net.NetworkCredential("Administrator", "Evaluation1", "CONTOSO")
$smSoapMessage  = "<?xml version='1.0' encoding='utf-8'?>" `
+ "<soap:Envelope xmlns:soap=`"http://schemas.xmlsoap.org/soap/envelope/`" " `
+ " xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns:xsd=`"http://www.w3.org/2001/XMLSchema`"" `
+ " xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`" >" `
+ "<soap:Body>" `
+ "<GetFolder xmlns=`"http://schemas.microsoft.com/exchange/services/2006/messages`" " `
+ " xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`"> " `
+ "<FolderShape> " `
+ "<t:BaseShape>Default</t:BaseShape> " `
+ "</FolderShape> " `
+ "<FolderIds> " `
+ "<t:DistinguishedFolderId Id=`"inbox`"/> " `
+ "</FolderIds> " `
+ "</GetFolder> " `
+ "</soap:Body></soap:Envelope>"

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
$ResponseXmlDoc = new-object System.Xml.XmlDocument
$ResponseXmlDoc.Load($ResponseStream)
$UreadNameNodes = @($ResponseXmlDoc.GetElementsByTagName("t:UnreadCount"))
"Number of Unread Email : " + $UreadNameNodes[0].'#text'


