## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]

## Load Managed API dll  
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"  
  
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

# Bind to the SentItems Folder
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$MailboxName)   
$SentItems = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
#Define ExtendedProps
$OriginalServerIP = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(
[Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::InternetHeaders,"x-ms-exchange-organization-originalserveripaddress",
[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)

$OriginalClientIP = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(
[Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::InternetHeaders,"x-ms-exchange-organization-originalclientipaddress",
[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)

$psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
$psPropset.add($OriginalServerIP)
$psPropset.add($OriginalClientIP)

#Setup GEOIP
$geoip = New-Object -ComObject "GeoIPCOMEx.GeoIPEx"
$geoip.set_db_path('c:\mec\')


#Define ItemView to retrive just 1000 Items    
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000) 
$ivItemView.PropertySet = $psPropset
$fiItems = $null
$rptCollection = @()
do{    
    $fiItems = $service.FindItems($SentItems.Id,$ivItemView)    
    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
    foreach($Item in $fiItems.Items){ 
		$rptobj = "" | select DateSent,Subject,ClientIP,LocationCity,LocationCountry,Latitute,Longitute
		"Subject : " + $Item.Subject
		$rptobj.Subject = $Item.Subject
		$rptobj.DateSent = $Item.DateTimeSent
		$clientIpvalue = $null
		if($Item.TryGetProperty($OriginalClientIP,[ref]$clientIpvalue)){
			$rptobj.ClientIP = $clientIpvalue
			"ClientIP : " + $clientIpvalue
			if($geoip.find_by_addr($clientIpvalue)){
				"Location City : " + $geoip.city
				$rptobj.LocationCity = $geoip.city
				"Location Country : " + $geoip.country_name
				$rptobj.LocationCountry = $geoip.country_name
				"Location Latitute : " + $geoip.latitude
				$rptobj.Latitute = $geoip.latitude
				"Location Longitute : " + $geoip.longitude
				$rptobj.Longitute = $geoip.longitude
			}
		}
		$serverIpValue = $null
		if($Item.TryGetProperty($OriginalServerIP,[ref]$serverIpValue)){
			"ServerIP : " + $serverIpValue
		} 
		$rptCollection += $rptobj

    }    
    $ivItemView.Offset += $fiItems.Items.Count    
}while($fiItems.MoreAvailable -eq $true) 

$tableStyle = @"
<style>
BODY{background-color:white;}
TABLE{border-width: 1px;
  border-style: solid;
  border-color: black;
  border-collapse: collapse;
}
TH{border-width: 1px;
  padding: 10px;
  border-style: solid;
  border-color: black;
  background-color:#66CCCC
}
TD{border-width: 1px;
  padding: 2px;
  border-style: solid;
  border-color: black;
  background-color:white
}
</style>
"@
  
$body = @"
<p style="font-size:25px;family:calibri;color:#ff9100">
$TableHeader
</p>
"@

$rptCollection | ConvertTo-HTML -head $tableStyle –body $body | Out-File c:\temp\GeoIpUserSentReport.html


