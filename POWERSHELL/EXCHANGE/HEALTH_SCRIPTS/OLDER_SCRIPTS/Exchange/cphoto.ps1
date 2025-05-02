$MailboxName = "user@mbmailbox.com"
$emEmailAdddrestoFind = "fred@fred.com"
$cntPhotoFile = "c:\ck1.jpg"

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($mailboxname)


$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)
$ContactsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$Iview = new-object Microsoft.Exchange.WebServices.Data.ItemView(1000)
$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ContactSchema]::EmailAddress1,$emEmailAdddrestoFind)
$frContactResults = $ContactsFolder.FindItems($SfSearchFilter,$Iview)
foreach ($cnContacts in $frContactResults.Items){
	$cnContacts.Subject
	$atattach = $cnContacts.Attachments.AddFileAttachment($cntPhotoFile)
        $atattach.IsContactPhoto = $true
	$cnContacts.update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
}