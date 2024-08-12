
$usertoSet = "user@domain.com"

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$wcmultiline = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings,"http://schemas.microsoft.com/exchange/wcmultiline", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean);

$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

$rfRootFolderID = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$usertoSet)

$rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$rfRootFolderID);
$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000);
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$ffResponse = $rfRootFolder.FindFolders($fvFolderView);
foreach ($ffFolder in $ffResponse.Folders)            {
	"Folder Name" + $ffFolder.DisplayName.ToString()
	$ffFolder.ExtendedProperties.Add($wcmultiline,"0")
	$ffFolder.update()

}

