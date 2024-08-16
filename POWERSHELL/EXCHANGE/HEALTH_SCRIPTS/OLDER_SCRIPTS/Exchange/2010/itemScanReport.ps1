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
$rptCollection = @()

#Define ItemType
$ItemType = "IPM.Note.Exchange.ActiveSync.MailboxLog"

$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName) 

$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

$sfItemSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,$ItemType)

$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  

$pfPropSet =  new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
$fvFolderView.PropertySet = $pfPropSet
$fvFolderView.PropertySet.Add($PR_Folder_Path)

#Define Function to convert String to FolderPath  
function ConvertToString($ipInputString){  
    $Val1Text = ""  
    for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){  
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))  
            $clInt++  
    }  
    return $Val1Text  
}  

$findFolderResults = $null
do{
	$findFolderResults = $tfTargetFolder.FindFolders($fvFolderView)
	foreach($folder in $findFolderResults.Folders){
			"Processing Folder " + $folder.DisplayName
			$foldpathval = $null  
	        #Try to get the FolderPath Value and then covert it to a usable String   
	        if ($folder.TryGetProperty($PR_Folder_Path,[ref] $foldpathval))  
	        {  
	            $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)  
	            $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }  
	            $hexString = $hexArr -join ''  
	            $hexString = $hexString.Replace("FEFF", "5C00")  
	            $fpath = ConvertToString($hexString)  
	        }  
			$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
			$rptObj = "" | select FolderPath,NumberofItems,SizeofItems
			$rptObj.NumberofItems = 0
			$rptObj.FolderPath = $fpath
			$findItemsResults = $null
			do{
				$findItemsResults = $folder.FindItems($sfItemSearchFilter,$ivItemView)
				foreach($itItem in $findItemsResults.Items){
					$rptObj.NumberofItems += 1
					$rptObj.SizeofItems += [INT32]$itItem.Size
				}
				$ivItemView.offset += $findItemsResults.Items.Count
			}while($findItemsResults.MoreAvailable -eq $true)
			if($rptObj.NumberofItems -gt 0){
				$rptCollection += $rptObj
			}
		}
	$fvFolderView.offset += $findFolderResults.Folders.Count
}while($findFolderResults.MoreAvailable -eq $true)
$rptCollection
$rptCollection | Export-Csv -NoTypeInformation c:\mbItemTypeReport.csv

