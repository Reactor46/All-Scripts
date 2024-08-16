$AqsString = "System.Message.DateReceived:01/01/2011..01/31/2011"
$MailboxName = "domain.com"

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$service.Credentials = New-Object System.Net.NetworkCredential("user@domain.com","passwod")

$service.AutodiscoverUrl($MailboxName,{$true})
$PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);

"Checking : " + $MailboxName 
$folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)
$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;
$sf1 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"2")
$sf2 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"allitems")
$sfSearchFilterCol = new-object  Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
$sfSearchFilterCol.Add($sf1)
$sfSearchFilterCol.Add($sf2)
$fiResult = $Service.FindFolders($folderidcnt,$sfSearchFilterCol,$fvFolderView)
$fiItems = $null
$ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
if($fiResult.Folders.Count -gt 0){
	$fiResult.Folders[0].DisplayName
	do{
		$fiItems = $fiResult.Folders[0].findItems($AqsString,$ItemView)
		$ItemView.offset += $fiItems.Items.Count
		foreach($Item in $fiItems.Items){
			$Item.Subject
		}
	}while($fiItems.MoreAvailable -eq $true)
}