
$MailboxName = "user@domain.com"
$ReportingCollection = @()

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)

$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

$rfRootFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)
$rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$rfRootFolderID)
$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10000);
$PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$Propset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$Propset.add($PR_MESSAGE_SIZE_EXTENDED)
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$fvFolderView.PropertySet = $Propset
$ffResponse = $rfRootFolder.FindFolders($fvFolderView);

foreach ($ffFolder in $ffResponse.Folders){
	$fldObject = "" | select FolderName,FolderSize
	$folderSize = $null
	$ptProptest2 = $ffFolder.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED, [ref]$folderSize)
	$fldObject.FolderName = $ffFolder.DisplayName
	$fldObject.FolderSize = [INT]$folderSize
	$ReportingCollection += $fldObject	 
}
$ReportingCollection	

