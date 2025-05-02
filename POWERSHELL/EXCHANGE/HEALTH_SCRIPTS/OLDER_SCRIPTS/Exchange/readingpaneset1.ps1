$MailboxName = "user@domain.com"

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($mailboxname)

$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)
$Root = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$OWAConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($service, "OWA.UserOptions", $Root.ParentFolderId, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)
if($OWAConfig.Dictionary.ContainsKey("previewmarkasread")){
       $OWAConfig.Dictionary["previewmarkasread"] = 2 }
else{
       $OWAConfig.Dictionary.Add("previewmarkasread", 2)
}
$OWAConfig.Update()     
"Done"   
