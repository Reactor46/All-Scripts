## Get the Mailbox to Access from the 1st commandline argument

$MailboxName = $args[0]

## Load Managed API dll  
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"  
  
## Set Exchange Version  
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013
  
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
$ExportCollection = @()
Write-Host "Process Recipient Cache"
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecipientCache,$MailboxName)   
$RecipientCache = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
#Define ItemView to retrive just 1000 Items    
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
$fiItems = $null    
do{    
    $fiItems = $service.FindItems($RecipientCache.Id,$ivItemView)    
    [Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
    foreach($Item in $fiItems.Items){     
		if($Item -is [Microsoft.Exchange.WebServices.Data.Contact]){
			$expObj = "" | select Source,DisplayName,Email1DisplayName,Email1Type,Email1EmailAddress
			$expObj.Source = "RecipientCache"
			$expObj.DisplayName = $Item.DisplayName
			if($Item.EmailAddresses.Contains([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1)){				
				$expObj.Email1DisplayName = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name
				$expObj.Email1Type = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].RoutingType
				$expObj.Email1EmailAddress = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address
			}
			$ExportCollection += $expObj
		}
    }    
    $ivItemView.Offset += $fiItems.Items.Count    
}while($fiItems.MoreAvailable -eq $true) 

Write-Host "Process Suggested Contacts"
$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName) 
$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)
$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"Suggested Contacts")
$findFolderResults = $service.FindFolders($folderid,$SfSearchFilter,$fvFolderView)

if($findFolderResults.Folders.Count -gt 0){
	$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
	#Define ItemView to retrive just 1000 Items    
	$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
	$fiItems = $null    
	do{    
	    $fiItems = $service.FindItems($findFolderResults.Folders[0].Id,$ivItemView)    
	    [Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
	    foreach($Item in $fiItems.Items){     
			if($Item -is [Microsoft.Exchange.WebServices.Data.Contact]){
				$expObj = "" | select Source,DisplayName,Email1DisplayName,Email1Type,Email1EmailAddress
				$expObj.Source = "Suggested Contacts"
				$expObj.DisplayName = $Item.DisplayName
				if($Item.EmailAddresses.Contains([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1)){				
					$expObj.Email1DisplayName = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name
					$expObj.Email1Type = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].RoutingType
					$expObj.Email1EmailAddress = $Item.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Address
				}
				$ExportCollection += $expObj
			}
	    }    
	    $ivItemView.Offset += $fiItems.Items.Count    
	}while($fiItems.MoreAvailable -eq $true) 
}
Write-Host "Process OWA AutocompleteCache"
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName) 
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)     
#Specify the Calendar folder where the FAI Item is  
$UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($service, "OWA.AutocompleteCache", $folderid, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)  
#Get the XML in String Format  
$acXML = [System.Text.Encoding]::UTF8.GetString($UsrConfig.XmlData)  
#Deal with the first character being a Byte Order Mark  
$boMark = $acXML.SubString(0,1)  
#Parse the XML  
[XML]$acXML = $acXML.SubString(1)  
foreach($AcEnt in $acXML.AutoCompleteCache.Entry){
	$expObj = "" | select Source,DisplayName,Email1DisplayName,Email1Type,Email1EmailAddress
	$expObj.Source = "OWA AutocompleteCache"
	$expObj.DisplayName = $AcEnt.displayName
	$expObj.Email1DisplayName= $AcEnt.displayName
	$expObj.Email1Type= "SMTP"
	$expObj.Email1EmailAddress= $AcEnt.smtpAddr
	$ExportCollection +=$expObj	
}
$fnFileName = "c:\temp\" + $MailboxName + "AC-CacheExport.csv"
$ExportCollection | Export-Csv -NoTypeInformation -Path $fnFileName
"Exported to " + $fnFileName