$MailboxName = "mailbox@domain.com"

$credentials = New-Object System.Net.NetworkCredential("user@domain.com","Password#")


$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010)

$Permission = [Microsoft.Exchange.WebServices.Data.FolderPermissionLevel]::Editor

$service.Credentials = $credentials
$service.AutodiscoverUrl($MailboxName,{$true})
$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName);


function addFolderPerm($folder){
	$existingperm = $null
	foreach($fperm in $folder.Permissions){
		if($fperm.UserId.StandardUser -eq [Microsoft.Exchange.WebServices.Data.StandardUser]::Default){
				$existingperm = $fperm
		}
	}
	if($existingperm -ne $null){
		$folder.Permissions.Remove($existingperm)
	} 
	$newfp = new-object Microsoft.Exchange.WebServices.Data.FolderPermission([Microsoft.Exchange.WebServices.Data.StandardUser]::Default,$Permission)
	$folder.Permissions.Add($newfp)
	$folder.Update()
}

"Checking : " + $MailboxName 
$folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)
$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderidcnt)
"Set Calendar Rights"
addFolderPerm($Calendar)
$sf1 = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"Freebusy Data")

$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;

$folderidRoot = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)
$fiResult = $Service.FindFolders($folderidRoot,$sf1,$fvFolderView)
if($fiResult.Folders.Count -eq 1){
	"Set FreeBusy Rights"
	$Freebusyfld = $fiResult.Folders[0]
	$Freebusyfld.Load()
	addFolderPerm($Freebusyfld)
}





