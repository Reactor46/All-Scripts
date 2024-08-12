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
$queryTime = (Get-Date).AddYears(-1)
$AQSString =  "System.Message.DateReceived:<" + $queryTime.ToString("MM/dd/yyyy")   
# Bind to the Inbox Folder
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)   
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
$fiItems = $null
$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.ItemId" -as "Type")
$archiveItems = [Activator]::CreateInstance($type)  
do{    
    $fiItems = $service.FindItems($Inbox.Id,$AQSString,$ivItemView)    
    #[Void]$service.LoadPropertiesForItems($fiItems,$psPropset)  
    foreach($Item in $fiItems.Items){
		$archiveItems.Add($Item.Id)		
	} 
    
    $ivItemView.Offset += $fiItems.Items.Count    
}while($fiItems.MoreAvailable -eq $true)
$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.ItemId" -as "Type") 
$batchArchive = [Activator]::CreateInstance($type) 
foreach($itItemID in $archiveItems){
	$batchArchive.add($itItemID)
	if($batchArchive.Count -eq 100){
		$service.ArchiveItems($batchArchive,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
		$batchArchive.clear()
	}
}
if($batchArchive.Count -gt 0){
	$service.ArchiveItems($batchArchive,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
	$batchArchive.clear()
}
