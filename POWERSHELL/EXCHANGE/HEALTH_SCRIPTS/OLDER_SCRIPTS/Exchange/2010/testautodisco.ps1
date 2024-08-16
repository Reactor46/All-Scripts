
function AutoDisco($EmailAddress){
## AutoDisco for EWS
$autodiscoResponse = "<Autodiscover xmlns=`"http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006`">" `
+ "	<Request>"`
+ "		<EMailAddress>" + $EmailAddress + "</EMailAddress>"`
+ "		<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>"`
+ "	</Request>"`
+ "</Autodiscover>"
$strRootURI = $scpval
"using Autodisovery URI : " + $strRootURI 
$WDRequest = [System.Net.WebRequest]::Create($strRootURI)
$WDRequest.ContentType = "text/xml"
$WDRequest.Headers.Add("Translate", "F")
$WDRequest.Method = "Post"
$WDRequest.UseDefaultCredentials = $True
$bytes = [System.Text.Encoding]::UTF8.GetBytes($autodiscoResponse)
$WDRequest.ContentLength = $bytes.Length
$RequestStream = $WDRequest.GetRequestStream()
$RequestStream.Write($bytes, 0, $bytes.Length)
$RequestStream.Close()
$WDResponse = $WDRequest.GetResponse()
$ResponseStream = $WDResponse.GetResponseStream()
$ResponseXmlDoc = new-object System.Xml.XmlDocument
$ResponseXmlDoc.Load($ResponseStream)
$ServerNodes = @($ResponseXmlDoc.getElementsByTagName("Server"))
if ($ServerNodes.length -ne 0){
	write-host $ServerNodes[0].'#text'
}
}


$ScpUrlGuidString = "77378F46-2C66-4aa9-A6A6-3E7A48B19596"
$ScpPtrGuidString = "67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68"

$ComputerSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::getComputerSite()
"ComputerSite : " + $ComputerSite
$root = [ADSI]'LDAP://RootDSE' 
$cfConfigRootpath = "LDAP://" + $root.ConfigurationNamingContext.tostring()
$configRoot = [ADSI]$cfConfigRootpath 
$searcher = new-object System.DirectoryServices.DirectorySearcher($configRoot)
$searcher.Filter = "(&(objectClass=serviceConnectionPoint)(|(keywords=" + $ScpPtrGuidString + ")(keywords=" + $ScpUrlGuidString + ")))"
$searchresult = $searcher.FindOne().GetDirectoryEntry()
[String]$scpval = $searchresult.serviceBindingInformation[0].ToString() 
AutoDisco("user@domain.com")