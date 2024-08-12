## EWS Managed API Connect Script
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
  
#CAS URL Option 1 Autodiscover  
$service.AutodiscoverUrl("email@domain.com",{$true})  
"Using CAS Server : " + $Service.url   
   
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, "email@domain.com")  


#Define the FolderSize Extended Property
$PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$Propset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$Propset.add($PR_MESSAGE_SIZE_EXTENDED)

$PFRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
$NonIPMPfRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$PFRoot.ParentFolderId)
$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"NON_IPM_SUBTREE")
$folders = $NonIPMPfRoot.Findfolders($sfSearchFilter,$fvFolderView)
foreach($folder in $folders.Folders){
	#$folder 
	$sfSearchFilter1 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"OFFLINE ADDRESS BOOK")
	$fvFolderView1 =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
	$fvFolderView1.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;
	$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
	$OABFolder = $folder.Findfolders($sfSearchFilter1,$fvFolderView1).Folders[0]
	$OABFolders = $OABFolder.Findfolders($fvFolderView1)
	foreach($OABSubFolder in $OABFolders.Folders){
		if($OABSubFolder.ChildFolderCount -gt 0){
			$OABSubFolder.DisplayName
			$fvFolderView2 =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
			$fvFolderView2.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;
			$fvFolderView2.PropertySet = $Propset
			$SubFolders = $OABSubFolder.Findfolders($fvFolderView2)
				foreach($SubFolder in $SubFolders.Folders){
				$rptObj = "" | select  RootFolderName,SubFolderName,FolderItemCount,FolderSize,NewestItemLastModified
				$rptObj.RootFolderName = $OABSubFolder.DisplayName
				$rptObj.SubFolderName = $SubFolder.DisplayName
				$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
				$FindItems = $SubFolder.FindItems($ivItemView)
				$rptObj.FolderItemCount = $FindItems.Items.Count
				if($FindItems.Items.Count -gt 0){
					$rptObj.NewestItemLastModified = $FindItems.Items[0].LastModifiedTime.ToString()
				}
				$folderSize = $null
				if($SubFolder.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED, [ref]$folderSize)){
					$rptObj.FolderSize = [MATH]::Round($folderSize/1024,0)
				}
					
				$rptCollection += $rptObj
			}

		}
		

	}
}
$rptCollection | Export-Csv -NoTypeInformation c:\temp\OabStatReports.csv