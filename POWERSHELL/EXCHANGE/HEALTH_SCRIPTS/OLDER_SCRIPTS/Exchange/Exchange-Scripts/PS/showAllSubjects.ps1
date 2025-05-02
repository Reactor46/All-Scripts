$RptObjColl = @()
$MailboxName = "user@domain.com"

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
$service.AutodiscoverUrl($MailboxName,{$true})  
"Using CAS Server : " + $Service.url   
   
#CAS URL Option 2 Hardcoded  
  
#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
#$service.Url = $uri    
  
## Optional section for Exchange Impersonation  
  
#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 

# Bind to the Archive Root folder  
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)   
$MsgRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
#Deep Transval will ensure all folders in the search path are returned 
 
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
#The Search filter will exclude any Search Folders 
$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
$PR_NORMALIZED_SUBJECT = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E1D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);   
$psPropertySet.add($PR_NORMALIZED_SUBJECT)
$PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
$sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")  
$fiResult = $null  
#The Do loop will handle any paging that is required if there are more the 1000 folders in a mailbox 
$rptHash = @{}
$AQSString = "kind:email" 
do { 
	$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
	$ivItemView.PropertySet = $psPropertySet
    $fiResult = $Service.FindFolders($folderid,$sfSearchFilter,$fvFolderView)  
    foreach($ffFolder in $fiResult.Folders){  
    "Processing Folder : " + $ffFolder.displayName 
	if($ffFolder.UnreadCount -gt 0){
		$fiResults = $null
		$updateColl = @()
		do{  
		    $fiResults = $ffFolder.findItems($AQSString,$ivItemView)
		    foreach($Item in $fiResults.Items){  
				$subject = $null
				if($Item.TryGetProperty($PR_NORMALIZED_SUBJECT,[ref]$subject)){
					if($subject -ne $null){
						"Processing Messsage : " + $subject
						if($rptHash.Contains($subject) -eq $false){
							$rptHash.add($subject,1);
						}
						else{
							$rptHash[$subject] +=1
						}
					}
				}
			}  
		    $ivItemView.Offset += $fiResults.Items.Count  
		}while($fiResults.MoreAvailable -eq $true)
	}         
    } 
    $fvFolderView.Offset += $fiResult.Folders.Count
}while($fiResult.MoreAvailable -eq $true) 

$rptHash.GetEnumerator() | Sort-Object value -Descending | ForEach-Object{
	$rptobj = "" | select Subject, NumberofMessages
	$rptobj.Subject = $_.Key
	$rptobj.NumberofMessages = $_.Value
	$RptObjColl += $rptobj
}
$RptObjColl | Export-Csv -NoTypeInformation -path c:\temp\subjectReport.csv
$RptObjColl