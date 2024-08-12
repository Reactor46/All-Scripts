## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]
$exportFolder = "c:\temp\"

$ABGUID = "6c118670-2f72-4213-944c-ab1e97d63f9b";

## Load Managed API dll  
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"  
  
## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2  
  
## Create Exchange Service Object  
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
  
## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
  
#Credentials Option 1 using UPN for the windows Account  
$psCred = Get-Credential  
$creds = New-Object System.Net.NetworkCredential($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())  
$service.Credentials = $creds      
  
#Credentials Option 2  
#service.UseDefaultCredentials = $true  
  
## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
  
## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

## end code from http://poshcode.org/624
  
## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
  
#CAS URL Option 1 Autodiscover  
$service.AutodiscoverUrl($MailboxName,{$true})  
"Using CAS Server : " + $Service.url   


   
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 

function getPeopleRequest($offset){

$request = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
<soap:Header>
<RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" />
</soap:Header><soap:Body>
<FindPeople xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"><PersonaShape>
<BaseShape xmlns="http://schemas.microsoft.com/exchange/services/2006/types">Default</BaseShape>
</PersonaShape><IndexedPageItemView MaxEntriesReturned="100" Offset="$offset" BasePoint="Beginning" />
<ParentFolderId>
<AddressListId Id="$ABGUID" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" />
</ParentFolderId></FindPeople></soap:Body></soap:Envelope>
"@
return $request
}

function getPersonaRequest($PersonalId){
$request = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013"/>
  </soap:Header>
  <soap:Body xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
    <GetPersona>
      <PersonaId Id="$PersonalId"/>
    </GetPersona>
  </soap:Body>
</soap:Envelope>
"@
return $request
}

function AutoDiscoverPhotoURL{
       param (
              $EmailAddress="$( throw 'Email is a mandatory Parameter' )",
              $Credentials="$( throw 'Credentials is a mandatory Parameter' )"
              )
       process{
              $version= [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
              $adService= New-Object Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService($version);
              $adService.Credentials = $Credentials
              $adService.EnableScpLookup=$false;
              $adService.RedirectionUrlValidationCallback= {$true}
              $adService.PreAuthenticate=$true;
              $UserSettings= new-object Microsoft.Exchange.WebServices.Autodiscover.UserSettingName[] 1
              $UserSettings[0] = [Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::ExternalPhotosUrl
              $adResponse=$adService.GetUserSettings($EmailAddress, $UserSettings)
              $PhotoURI= $adResponse.Settings[[Microsoft.Exchange.WebServices.Autodiscover.UserSettingName]::ExternalPhotosUrl]
              return $PhotoURI.ToString()
       }
}

$Script:PhotoURL = AutoDiscoverPhotoURL -EmailAddress $MailboxName  -Credentials $creds
Write-host ("Photo URL : " + $Script:PhotoURL) 

function ProcessPersona($PersonalId){


	$personalRequest = getPersonaRequest ($PersonalId)
	$mbMailboxFolderURI = New-Object System.Uri($service.url)  
	$wrWebRequest = [System.Net.WebRequest]::Create($mbMailboxFolderURI)  
	$wrWebRequest.CookieContainer =  New-Object System.Net.CookieContainer   
	$wrWebRequest.KeepAlive = $false;  
	$wrWebRequest.Headers.Set("Pragma", "no-cache");  
	$wrWebRequest.Headers.Set("Translate", "f");  
	$wrWebRequest.Headers.Set("Depth", "0");  
	$wrWebRequest.ContentType = "text/xml";  
	$wrWebRequest.ContentLength = $expRequest.Length;  
	$wrWebRequest.Timeout = 90000;  
	$wrWebRequest.Method = "POST";  
	$wrWebRequest.Credentials = $creds  
	$wrWebRequest.UserAgent = "EWS Script"
	
	$bqByteQuery = [System.Text.Encoding]::ASCII.GetBytes($personalRequest);  
	$wrWebRequest.ContentLength = $bqByteQuery.Length;  
	$rsRequestStream = $wrWebRequest.GetRequestStream();  
	$rsRequestStream.Write($bqByteQuery, 0, $bqByteQuery.Length);  
	$rsRequestStream.Close();  
	$wrWebResponse = $wrWebRequest.GetResponse();  
	$rsResponseStream = $wrWebResponse.GetResponseStream()  
	$sr = new-object System.IO.StreamReader($rsResponseStream);  
	$rdResponseDocument = New-Object System.Xml.XmlDocument  
	$rdResponseDocument.LoadXml($sr.ReadToEnd());  
	$Persona =@($rdResponseDocument.getElementsByTagName("Persona")) 
	$Persona
	$DisplayName = "";
	if($Persona.DisplayName -ne $null){
		$DisplayName = $Persona.DisplayName."#text"
	} 
	$fileName =  $exportFolder + $DisplayName + "-" + [Guid]::NewGuid().ToString() + ".vcf"
	add-content -path $filename "BEGIN:VCARD"
	add-content -path $filename "VERSION:2.1"
	$givenName = ""
	if($Persona.GivenName -ne $null){
		$givenName = $Persona.GivenName."#text"
	}
	$surname = ""
	if($Persona.Surname -ne $null){
		$surname = $Persona.Surname."#text"
	}
	add-content -path $filename ("N:" + $surname + ";" + $givenName)
	add-content -path $filename ("FN:" + $Persona.DisplayName."#text")
	$Department = "";
	if($Persona.Department -ne $null){
		$Department = $Persona.Department."#text"
	}
	if($Persona.EmailAddress -ne $null){
		add-content -path $filename ("EMAIL;PREF;INTERNET:" + $Persona.EmailAddress.EmailAddress)
	}
	$CompanyName = "";
	if($Persona.CompanyName -ne $null){
		$CompanyName = $Persona.CompanyName."#text"
	}
	add-content -path $filename ("ORG:" + $CompanyName + ";" + $Department)	
	if($Persona.Titles -ne $null){
		add-content -path $filename ("TITLE:" + $Persona.Titles.StringAttributedValue.Value)
	}
	if($Persona.MobilePhones -ne $null){
		add-content -path $filename ("TEL;CELL;VOICE:" + $Persona.MobilePhones.PhoneNumberAttributedValue.Value.Number)		
	}
	if($Persona.HomePhones -ne $null){
		add-content -path $filename ("TEL;HOME;VOICE:" + $Persona.HomePhones.PhoneNumberAttributedValue.Value.Number)		
	}
	if($Persona.BusinessPhoneNumbers -ne $null){
		add-content -path $filename ("TEL;WORK;VOICE:" + $Persona.BusinessPhoneNumbers.PhoneNumberAttributedValue.Value.Number)		
	}
	if($Persona.WorkFaxes -ne $null){
		add-content -path $filename ("TEL;WORK;FAX:" + $Persona.WorkFaxes.PhoneNumberAttributedValue.Value.Number)
	}
	if($Persona.BusinessHomePages -ne $null){
		add-content -path $filename ("URL;WORK:" + $Persona.BusinessHomePages.StringAttributedValue.Value)
	}
	if($Persona.BusinessAddresses -ne $null){
		$Country = $Persona.BusinessAddresses.PostalAddressAttributedValue.Value.Country
		$City = $Persona.BusinessAddresses.PostalAddressAttributedValue.Value.City
		$Street = $Persona.BusinessAddresses.PostalAddressAttributedValue.Value.Street
		$State = $Persona.BusinessAddresses.PostalAddressAttributedValue.Value.State
		$PCode = $Persona.BusinessAddresses.PostalAddressAttributedValue.Value.PostalCode
		$addr =  "ADR;WORK;PREF:;" + $Country + ";" + $Street + ";" +$City + ";" + $State + ";" + $PCode + ";" + $Country
		add-content -path $filename $addr
	}
	try{
		$PhotoSize = "HR96x96" 
		$PhotoURL= $Script:PhotoURL + "/GetUserPhoto?email="  + $Persona.EmailAddress.EmailAddress + "&size=" + $PhotoSize;
		$wbClient = new-object System.Net.WebClient
		$wbClient.Credentials = $creds
		$photoBytes = $wbClient.DownloadData($PhotoURL);
		add-content -path $filename "PHOTO;ENCODING=BASE64;TYPE=JPEG:"
		$ImageString = [System.Convert]::ToBase64String($photoBytes,[System.Base64FormattingOptions]::InsertLineBreaks)
		add-content -path $filename $ImageString
		add-content -path $filename "`r`n"	
	}
	catch{

	}
	add-content -path $filename "END:VCARD"

}

$peopleCollection = @()
$offset = 0;

do{
	$mbMailboxFolderURI = New-Object System.Uri($service.url)  
	$wrWebRequest = [System.Net.WebRequest]::Create($mbMailboxFolderURI)  
	$wrWebRequest.CookieContainer =  New-Object System.Net.CookieContainer   
	$wrWebRequest.KeepAlive = $false;  
	$wrWebRequest.Useragent = "EWS Script"
	$wrWebRequest.Headers.Set("Pragma", "no-cache");  
	$wrWebRequest.Headers.Set("Translate", "f");  
	$wrWebRequest.Headers.Set("Depth", "0");  
	$wrWebRequest.ContentType = "text/xml";  
	$wrWebRequest.ContentLength = $expRequest.Length;  
	$wrWebRequest.Timeout = 60000;  
	$wrWebRequest.Method = "POST";  
	$wrWebRequest.Credentials = $creds  

	$fpRequest = getPeopleRequest ($offset)
	$bqByteQuery = [System.Text.Encoding]::ASCII.GetBytes($fpRequest);  
	$wrWebRequest.ContentLength = $bqByteQuery.Length;  
	$rsRequestStream = $wrWebRequest.GetRequestStream();  
	$rsRequestStream.Write($bqByteQuery, 0, $bqByteQuery.Length);  
	$rsRequestStream.Close();  
	$wrWebResponse = $wrWebRequest.GetResponse();  
	$rsResponseStream = $wrWebResponse.GetResponseStream()  
	$sr = new-object System.IO.StreamReader($rsResponseStream);  
	$rdResponseDocument = New-Object System.Xml.XmlDocument  
	$rdResponseDocument.LoadXml($sr.ReadToEnd());  
	$totalCount = @($rdResponseDocument.getElementsByTagName("TotalNumberOfPeopleInView")) 
	$Personas =@($rdResponseDocument.getElementsByTagName("Persona"))  
	Write-Host ("People Count : " + $Personas.Count)
	$offset += $Personas.Count
	foreach($persona in $Personas){
		if($persona.PersonaType -eq "Person"){
			ProcessPersona($persona.PersonaId.Id.ToString()) 
		}
	}
	[Int32]$tc = $totalCount."#text"
	Write-Host ("Offset: " + $offset)
	Write-Host ("Total count: " + $tc)

}while($tc -gt $offset)


