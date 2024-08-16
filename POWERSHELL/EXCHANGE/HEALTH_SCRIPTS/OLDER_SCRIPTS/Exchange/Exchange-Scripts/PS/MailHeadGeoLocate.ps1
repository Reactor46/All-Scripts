## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]
$rptCollection = @()

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

# Bind to the Inbox Folder
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::JunkEmail,$MailboxName)   
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

#Define Property for the Message Header
$PR_TRANSPORT_MESSAGE_HEADERS = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x007D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);

$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
$psPropset.add($PR_TRANSPORT_MESSAGE_HEADERS)

#Setup GEOIP
$geoip = New-Object -ComObject "GeoIPCOMEx.GeoIPEx"
$geoip.set_db_path('c:\mec\')

#Define ItemView to retrive just 10 Items    
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(50)        
    $fiItems = $service.FindItems($Inbox.Id,$ivItemView)    
    [Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
    foreach($Item in $fiItems.Items){      
		$MailHeaderVal = $null
		if($Item.TryGetProperty($PR_TRANSPORT_MESSAGE_HEADERS,[ref]$MailHeaderVal)){	
			$rptObj = "" | Select DateTimeReceived,From,Subject,Size,TransitCities
			$rptObj.DateTimeReceived = $Item.DateTimeReceived 
			$rptObj.From = $Item.Sender.Address
			if($Item.Subject.Length -gt 50){
				$rptObj.Subject = $Item.Subject.SubString(0,50)
			}
			else{
				$rptObj.Subject = $Item.Subject
			}
			$rptObj.Size = $Item.Size
			$SMTPTrace = $MailHeaderVal.Substring(0,$MailHeaderVal.IndexOf("From:"))
			# RegEx for IP address
			$RegExIP = '\b(([01]?\d?\d|2[0-4]\d|25[0-5])\.){3}([01]?\d?\d|2[0-4]\d|25[0-5])\b'
			$matchedItems = [regex]::matches($SMTPTrace, $RegExIP)
			$lochash = @{}
			foreach($Match in $matchedItems){
				$preVal = $MailHeaderVal.SubString(($Match.Index-1),1)
				if($preVal -eq "[" -bor $preVal -eq "("){
					if($geoip.find_by_addr($Match.Value)){
						if($geoip.country_name -ne "Localhost" -band $geoip.country_name -ne "Local Area Network"){
							if($lochash.ContainsKey(($geoip.city + "-" +  $geoip.country_name)) -eq $false){
								$lochash.add(($geoip.city + "-" +  $geoip.country_name),1)
								$rptObj.TransitCities = $rptObj.TransitCities + $geoip.city + "-" +  $geoip.country_name + ";"
							}
						}
					}
				}
			}
			$rptCollection += $rptObj
		}          
    }
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

$rptCollection | ConvertTo-HTML -head $tableStyle –body $body | Out-File c:\temp\jnkReport.htm
