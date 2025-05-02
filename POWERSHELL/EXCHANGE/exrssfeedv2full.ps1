function GetItem($smSoapMessage){
$bdBodytext = ""
$WDRequest1 = [System.Net.WebRequest]::Create($strRootURI)
$WDRequest1.ContentType = "text/xml"
$WDRequest1.Headers.Add("Translate", "F")
$WDRequest1.Method = "Post"
$WDRequest1.Credentials = $cdUsrCredentials
$bytes1 = [System.Text.Encoding]::UTF8.GetBytes($smSoapMessage)
$WDRequest1.ContentLength = $bytes1.Length
$RequestStream1 = $WDRequest1.GetRequestStream()
[void]$RequestStream1.Write($bytes1, 0, $bytes1.Length)
[void]$RequestStream1.Close()
$WDResponse1 = $WDRequest1.GetResponse()
$ResponseStream1 = $WDResponse1.GetResponseStream()
$ResponseXmlDoc1 = new-object System.Xml.XmlDocument
$ResponseXmlDoc1.Load($ResponseStream1)  
$tbBodyNodes = @($ResponseXmlDoc1.getElementsByTagName("t:Body"))
for($itemNums=0;$itemNums -lt $tbBodyNodes.Count;$itemNums++){
	$bdBodytext = $tbBodyNodes[$itemNums].'#text'.ToString()
}
return $bdBodytext
}


$snServername = "servername"
$unUserName = "user"
$psPassword = "password"
$dnDomainName = "domain"
$mbMailboxToAccess = "user@smtpdomain.com"
$cdUsrCredentials = new-object System.Net.NetworkCredential($unUserName , $psPassword , $dnDomainName)
$xsXmlFileName = "c:\feedname.xml"
[System.Reflection.Assembly]::LoadWithPartialName("System.Web") > $null
$xrXmlWritter = new-object System.Xml.XmlTextWriter($xsXmlFileName,[System.Text.Encoding]::UTF8)
$xrXmlWritter.WriteStartDocument()
$xrXmlWritter.WriteStartElement("rss")
$xrXmlWritter.WriteAttributeString("version", "2.0")
$xrXmlWritter.WriteStartElement("channel")
$xrXmlWritter.WriteElementString("title", "Inbox Feed For " + $mbMailboxToAccess)
$xrXmlWritter.WriteElementString("link", "https://" + $snServerName + "/owa/")
$xrXmlWritter.WriteElementString("description", "Exchange Inbox Feed For" + $mbMailboxToAccess)
$datetimetoquery = get-date
$smSoapMessage  = "<?xml version='1.0' encoding='utf-8'?>" `
+ "<soap:Envelope xmlns:soap=`"http://schemas.xmlsoap.org/soap/envelope/`" " `
+ " xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns:xsd=`"http://www.w3.org/2001/XMLSchema`"" `
+ " xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`" >" `
+ "<soap:Header>" `
+ "<t:ExchangeImpersonation>" `
+ "<t:ConnectingSID>" `
+ "<t:PrimarySmtpAddress>" + $mbMailboxToAccess + "</t:PrimarySmtpAddress>" `
+ "</t:ConnectingSID>" `
+ "</t:ExchangeImpersonation>" `
+ "</soap:Header>" `
+ "<soap:Body>" `
+ "<FindItem xmlns=`"http://schemas.microsoft.com/exchange/services/2006/messages`" " `
+ "xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`" Traversal=`"Shallow`"> " `
+ "<ItemShape>" `
+ "<t:BaseShape>AllProperties</t:BaseShape>" `
+ "<AdditionalProperties xmlns=""http://schemas.microsoft.com/exchange/services/2006/types"">" `
+ "<ExtendedFieldURI PropertyTag=""0x3FD9"" PropertyType=""String"" />" `
+ "<ExtendedFieldURI PropertyTag=""0x10F3"" PropertyType=""String"" />" `
+ "<ExtendedFieldURI PropertyTag=""0x0C1A"" PropertyType=""String"" />" `
+ "</AdditionalProperties>" `
+ "</ItemShape>" `
+ "<Restriction>" `
+ "<t:IsGreaterThanOrEqualTo>" `
+ "<t:FieldURI FieldURI=`"item:DateTimeSent`"/>"`
+ "<t:FieldURIOrConstant>" `
+ "<t:Constant Value=`"" + $datetimetoquery.ToUniversalTime().AddDays(-7).ToString("yyyy-MM-ddThh:mm:ssZ")  + "`"/>"`
+ "</t:FieldURIOrConstant>"`
+ "</t:IsGreaterThanOrEqualTo>"`
+ "</Restriction>"`
+ "<ParentFolderIds>" `
+ "<t:DistinguishedFolderId Id=`"inbox`"/>" `
+ "</ParentFolderIds>" `
+ "</FindItem>" `
+ "</soap:Body></soap:Envelope>"

$strRootURI = "https://" + $snServername + "/ews/Exchange.asmx"
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
$IDNodes = @($ResponseXmlDoc.getElementsByTagName("t:ItemId"))
$dsDescription = @($ResponseXmlDoc.getElementsByTagName("t:Value"))
for($i=0;$i -lt $subjectnodes.Count;$i++){
	$Senttime = [System.Convert]::ToDateTime($SentNodes[$i].'#text'.ToString())
	$Senttime.ToString()  + "	" + $FromNodes[$i].'#text' + "	" + $subjectnodes[$i].'#text' + "	" + $SizeNodes[$i].'#text' 
	$IdNodeID = $IDNodes[$i].GetAttributeNode("Id")	
	$ckChangeKey = $IDNodes[$i].GetAttributeNode("ChangeKey")
	$smSoapMessage  = "<?xml version='1.0' encoding='utf-8'?>" `
	+ "<soap:Envelope xmlns:soap=`"http://schemas.xmlsoap.org/soap/envelope/`" " `
	+ " xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" xmlns:xsd=`"http://www.w3.org/2001/XMLSchema`"" `
	+ " xmlns:t=`"http://schemas.microsoft.com/exchange/services/2006/types`" >" `
	+ "<soap:Header>" `
	+ "<t:ExchangeImpersonation>" `
	+ "<t:ConnectingSID>" `
	+ "<t:PrimarySmtpAddress>" + $mbMailboxToAccess + "</t:PrimarySmtpAddress>" `
	+ "</t:ConnectingSID>" `
	+ "</t:ExchangeImpersonation>" `
	+ "</soap:Header>" `
	+ "<soap:Body><GetItem xmlns=`"http://schemas.microsoft.com/exchange/services/2006/messages`"><ItemShape>" `
	+ "<BaseShape xmlns=`"http://schemas.microsoft.com/exchange/services/2006/types`">Default</BaseShape></ItemShape>" `
	+ "<ItemIds><ItemId Id=`"" + $IdNodeID.'#text' + "`"" `
	+ " xmlns=`"http://schemas.microsoft.com/exchange/services/2006/types`" /></ItemIds></GetItem></soap:Body>" `
	+ "</soap:Envelope>"
	$xrXmlWritter.WriteStartElement("item")
        $xrXmlWritter.WriteElementString("title", $subjectnodes[$i].'#text')
        $xrXmlWritter.WriteElementString("link", "https://" + $snServername + "/owa/?ae=Item&t=IPM.Note&id=Rg" + [System.Web.HttpUtility]::UrlEncode($IdNodeID.'#text').Substring(58).Replace("%3d","J"))
        $xrXmlWritter.WriteElementString("author", $FromNodes[$i].'#text')
        $xrXmlWritter.WriteStartElement("description")
        $xrXmlWritter.WriteRaw("<![CDATA[")
	$bdBodytext = GetItem($smSoapMessage)
        $xrXmlWritter.WriteRaw($bdBodytext)
        $xrXmlWritter.WriteRaw("]]>")
        $xrXmlWritter.WriteEndElement()
        $xrXmlWritter.WriteElementString("pubDate", $Senttime.ToString("r"))
        $xrXmlWritter.WriteElementString("guid", $IdNodeID.'#text')
        $xrXmlWritter.WriteEndElement()

	}
$xrXmlWritter.WriteEndElement()
$xrXmlWritter.WriteEndElement()
$xrXmlWritter.WriteEndDocument()
$xrXmlWritter.Close()
"Done"

