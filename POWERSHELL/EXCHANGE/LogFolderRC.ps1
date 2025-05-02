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
  
   
    
    

Get-Mailbox -ResultSize Unlimited | ForEach-Object{   
    $MailboxName = $_.PrimarySMTPAddress.ToString()  
    "Processing Mailbox : " + $MailboxName  
    if($service.url -eq $null){  
        ## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
  
        #CAS URL Option 1 Autodiscover  
        $service.AutodiscoverUrl($MailboxName,{$true})  
        "Using CAS Server : " + $Service.url   
          
        #CAS URL Option 2 Hardcoded  
        #$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
        #$service.Url = $uri    
    }  
  
	$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
	$PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);  

	$PR_ADDITIONAL_REN_ENTRYIDS = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x36D8, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::BinaryArray); 
	$Propset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
	$Propset.add($PR_ADDITIONAL_REN_ENTRYIDS)
	$Propset.add($PR_MESSAGE_SIZE_EXTENDED)
	#Sync Folders

	$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)   
	$RootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid,$Propset)
	$objVal = $null

	function ConvertFolderid($hexId){
		$aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId  
		$aiItem.Mailbox = $MailboxName  
		$aiItem.UniqueId = $hexId
		$aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::HexEntryId;  
		return $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EWSId) 
	}

	if($RootFolder.TryGetProperty($PR_ADDITIONAL_REN_ENTRYIDS,[ref]$objVal)){
		if($objVal[0] -ne $null){
			$rptobj = "" | Select MailboxName,SyncIssuesCount,SyncIssuesSize,ConflictsCount,ConflictsSize,LocalFailuresCount,LocalFailuresSize,ServerFailuresCount,ServerFailuresSize
			$rptobj.MailboxName = $MailboxName
			$cfid = ConvertFolderid([System.BitConverter]::ToString($objVal[0]).Replace("-",""))
			if($cfid.UniqueId -ne $null){
			$ConflictsFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId($cfid.UniqueId.ToString())
			$ConflictsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$ConflictsFolderId,$Propset)
			$ConflictsFolder.DisplayName
			$folderSize = $null
			if($ConflictsFolder.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED,[ref]$folderSize)){
				$rptobj.ConflictsCount = $ConflictsFolder.TotalCount
				$rptobj.ConflictsSize = [Math]::Round($folderSize/1MB) 
				"ItemCount  : " + $ConflictsFolder.TotalCount
				"FolderSize : " + [Math]::Round($folderSize/1MB) + " MB"
			}
			$siId = ConvertFolderid([System.BitConverter]::ToString($objVal[1]).Replace("-",""))
			$SyncIssuesFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId($siId.UniqueId.ToString())
			$SyncIssueFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$SyncIssuesFolderID,$Propset)
			$SyncIssueFolder.DisplayName
			if($SyncIssueFolder.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED,[ref]$folderSize)){
				$rptobj.SyncIssuesCount = $SyncIssueFolder.TotalCount
				$rptobj.SyncIssuesSize = [Math]::Round($folderSize/1MB)
				"ItemCount  : " + $SyncIssueFolder.TotalCount
				"FolderSize : " + [Math]::Round($folderSize/1MB) + " MB"
			}
			$lcId = ConvertFolderid([System.BitConverter]::ToString($objVal[2]).Replace("-",""))
			$localFailureId = new-object Microsoft.Exchange.WebServices.Data.FolderId($lcId.UniqueId.ToString())
			$localFailureFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$localFailureId,$Propset)
			$localFailureFolder.DisplayName
			if($localFailureFolder.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED,[ref]$folderSize)){
				$rptobj.LocalFailuresCount = $localFailureFolder.TotalCount
				$rptobj.LocalFailuresSize = [Math]::Round($folderSize/1MB)
				"ItemCount  : " + $localFailureFolder.TotalCount
				"FolderSize : " + [Math]::Round($folderSize/1MB) + " MB"
			}
			$sfid = ConvertFolderid([System.BitConverter]::ToString($objVal[3]).Replace("-",""))
			$ServerFailureId = new-object Microsoft.Exchange.WebServices.Data.FolderId($sfid.UniqueId.ToString())
			$ServerFailureFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$ServerFailureId,$Propset)
			$ServerFailureFolder.DisplayName
			if($ServerFailureFolder.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED,[ref]$folderSize)){
				$rptobj.ServerFailuresCount = $ServerFailureFolder.TotalCount
				$rptobj.ServerFailuresSize = [Math]::Round($folderSize/1MB)
				"ItemCount  : " + $ServerFailureFolder.TotalCount
				"FolderSize : " + [Math]::Round($folderSize/1MB) + " MB"
			}
			$rptCollection += $rptobj
			}
		}
	}
}
$rptCollection
$rptCollection | Export-Csv -NoTypeInformation -Path c:\temp\SyncFolderReport.csv  

