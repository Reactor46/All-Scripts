## Get the Mailbox to Access from the 1st commandline argument
$rptCollection = @()
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

# Bind to the Inbox Folder
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::JunkEmail,$MailboxName)   
$JunkEmail = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

$PidLidSpamOriginalFolder = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Common,0x859C, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
$psPropset.Add($PidLidSpamOriginalFolder)

#Define ItemView to retrive just 1000 Items
$fldIdHash = @{}
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
$ivItemView.PropertySet = $psPropset 
$fiItems = $null    
do{    
    $fiItems = $service.FindItems($JunkEmail.Id,$ivItemView)    
    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
    foreach($Item in $fiItems.Items){ 
		$rptobj = "" | Select DateTimeReceived,Subject,FilteredBy,OriginalFolder
		$propval = $null
		$rptobj.Subject = $Item.Subject
		$rptobj.DateTimeReceived = $Item.DateTimeReceived
		if($Item.TryGetProperty($PidLidSpamOriginalFolder,[ref]$propval)){

			$fldName = ""
			$FolderEntryId = $null
			$FolderEntryId = [System.BitConverter]::ToString($propval).Replace("-","")
			if($fldIdHash.ContainsKey($FolderEntryId) -eq $false){
				    $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId      
				    $aiItem.Mailbox = $MailboxName      
				    $aiItem.UniqueId = $FolderEntryId   
				    $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId;  
					$fldId = $null
				    $fldId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId)  
					$fld = $null
					if($fldId -ne $null){
						$orgID = new-object Microsoft.Exchange.WebServices.Data.FolderId($fldId.UniqueId)   
						$fld = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$orgID)
				        $fldIdHash.Add($FolderEntryId,$fld.DisplayName)
						$fldName = $fld.DisplayName
					}
							
			}
			else{
				$fldName = $fldIdHash[$FolderEntryId]
			}
			$rptobj.FilteredBy = "Client"
			$rptobj.OriginalFolder = $fldName	
		}
		else{
			$rptobj.FilteredBy = "Server"			
		}	
		$rptCollection += $rptobj
    }    
    $ivItemView.Offset += $fiItems.Items.Count    
}while($fiItems.MoreAvailable -eq $true) 
$rptCollection | Export-Csv c:\temp\junkemailClientreport.csv -NoTypeInformation