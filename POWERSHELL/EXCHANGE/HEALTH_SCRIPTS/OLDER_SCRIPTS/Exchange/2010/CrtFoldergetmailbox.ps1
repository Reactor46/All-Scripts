function CreateFolder($MailboxName) {
	"Mailbox Name : " + $MailboxName
	$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
	$ibInboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
	$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)
	$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$newFolderName)
	$findFolderResults = $service.FindFolders($ibInboxFolder.Id,$SfSearchFilter,$fvFolderView)
	if ($findFolderResults.TotalCount -eq 0){
		"Doesn't Exist"
		$NewFolder = new-object Microsoft.Exchange.WebServices.Data.Folder($service)
		$NewFolder.DisplayName = $newFolderName
		# $NewFolder.Save($ibInboxFolder.Id.UniqueId)
 		"Folder Created"
	}
	else{
		"Folder Already Exist - Do Nothing"
	}
}


$newFolderName = "mynewfolder123"

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)
$frun = $true


Get-mailbox | foreach-object {
	$WindowsEmailAddress = $_.WindowsEmailAddress.ToString()
	if ($frun -eq $true) {
		$frun = $false
		$service.AutodiscoverUrl($WindowsEmailAddress)
	}
	CreateFolder($WindowsEmailAddress)
}
	




