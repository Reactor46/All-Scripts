$servername = "ws03r2eeexchlcs"
$cdUsrCredentials = new-object System.Net.NetworkCredential("Administrator", "Evaluation1", "CONTOSO")
$datetimetoquery = get-date



$smSoapMessage  = "<?xml version='1.0' encoding='utf-8'?>" `
+ "<soap:Envelope xmlns:soap=`"http://schemas.xmlsoap.org/soap/envelope/`" " `
+ " xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns:xsd=`"http://www.w3.org/2001/XMLSchema`"" `
+ " xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`" >" `
+ "<soap:Body>" `
+ "<FindItem xmlns=`"http://schemas.microsoft.com/exchange/services/2006/messages`" " `
+ "xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`" Traversal=`"Shallow`"> " `
+ "<ItemShape>" `
+ "<t:BaseShape>AllProperties</t:BaseShape>" `
+ "</ItemShape>" `
+ "<Restriction>" `
+ "<t:And>" `
+ "<t:IsEqualTo>" `
+ "<t:FieldURI FieldURI=`"message:IsRead`"/>"`
+ "<t:FieldURIOrConstant>" `
+ "<t:Constant Value=`"0`"/>" `
+ "</t:FieldURIOrConstant>" `
+ "</t:IsEqualTo>" `
+ "<t:IsGreaterThanOrEqualTo>" `
+ "<t:FieldURI FieldURI=`"item:DateTimeSent`"/>"`
+ "<t:FieldURIOrConstant>" `
+ "<t:Constant Value=`"" + $datetimetoquery.ToUniversalTime().AddDays(-1).ToString("yyyy-MM-ddThh:mm:ssZ")  + "`"/>"`
+ "</t:FieldURIOrConstant>"`
+ "</t:IsGreaterThanOrEqualTo>"`
+ "</t:And>"`
+ "</Restriction>"`
+ "<ParentFolderIds>" `
+ "<t:DistinguishedFolderId Id=`"inbox`"/>" `
+ "</ParentFolderIds>" `
+ "</FindItem>" `
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
$subjectnodes = @($ResponseXmlDoc.getElementsByTagName("t:Subject"))
$FromNodes = @($ResponseXmlDoc.getElementsByTagName("t:Name"))
$SentNodes = @($ResponseXmlDoc.getElementsByTagName("t:DateTimeSent"))
$SizeNodes = @($ResponseXmlDoc.getElementsByTagName("t:Size"))
for($i=0;$i -lt $subjectnodes.Count;$i++){
	$Senttime = [System.Convert]::ToDateTime($SentNodes[$i].'#text'.ToString())
	$Senttime.ToString()  + "	" + $FromNodes[$i].'#text' + "	" + $subjectnodes[$i].'#text' + "	" + $SizeNodes[$i].'#text' 
}


