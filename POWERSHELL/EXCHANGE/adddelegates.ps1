$delegatetoAdd = "delegate@youdomain.com"

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

$mbMailbox = new-object Microsoft.Exchange.WebServices.Data.Mailbox($aceuser.mail.ToString())
$dgUser = new-object Microsoft.Exchange.WebServices.Data.DelegateUser($delegatetoAdd)
$dgUser.ViewPrivateItems = $false
$dgUser.ReceiveCopiesOfMeetingMessages = $false
$dgUser.Permissions.CalendarFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::Editor
$dgUser.Permissions.InboxFolderPermissionLevel = [Microsoft.Exchange.WebServices.Data.DelegateFolderPermissionLevel]::Reviewer
$dgArray = new-object Microsoft.Exchange.WebServices.Data.DelegateUser[] 1
$dgArray[0] = $dgUser
$service.AddDelegates($mbMailbox, [Microsoft.Exchange.WebServices.Data.MeetingRequestsDeliveryScope]::DelegatesAndMe, $dgArray);