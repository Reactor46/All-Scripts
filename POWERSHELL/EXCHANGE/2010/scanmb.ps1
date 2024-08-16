## EWS Managed API Connect Module Script written by Glen Scales
## Requires the EWS Managed API and Powershell V2.0 or greator

## Load Managed API dll
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"

## Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1

## Create Exchange Service Object
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials

#Credentials Option 1 using UPN for the windows Account
$creds = New-Object System.Net.NetworkCredential("user@domain.com","password") 
$service.Credentials = $creds    

#Credentials Option 2
#service.UseDefaultCredentials = $true

## Choose to ignore any SSL Warning issues caused by Self Signed Certificates

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use

$MailboxName = "user@domain.com"
#CAS URL Option 1 Autodiscover
$service.AutodiscoverUrl($MailboxName,{$true})
"Using CAS Server : " + $Service.url 

#Define Query Time

$queryTime = [system.DateTime]::Now.AddDays(-1)

$PR_LOCAL_COMMIT_TIME_MAX = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x670A, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime);

$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName) 
$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan($PR_LOCAL_COMMIT_TIME_MAX,$queryTime)

##Option add in the Deletions Folder
$delfolderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsDeletions,$MailboxName) 
$DelFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$delfolderid)

$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

$findFolderResults = $tfTargetFolder.FindFolders($SfSearchFilter,$fvFolderView)
$findFolderResults.Folders.Add($DelFolder)

$AQSString =  "System.Message.DateReceived:>" + $queryTime.ToString("MM/dd/yyyy")
$AQSString
foreach($folder in $findFolderResults.Folders){
	if($folder.TotalCount -gt 0){
		"Processing Folder " + $folder.DisplayName
		$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
		$findItemsResults = $null
		do{
			$findItemsResults = $folder.FindItems($AQSString,$ivItemView)
			foreach($itItem in $findItemsResults.Items){
				$itItem.Subject
			}
			$ivItemView.offset += $findItemsResults.Items.Count
		}while($findItemsResults.MoreAvailable -eq $true)
	}
}

