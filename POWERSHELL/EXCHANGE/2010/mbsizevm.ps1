$MailboxName = "user@domain.com"
##$folderItemType = "IPF.Note"
$rptCollection = @()


$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)


$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind
$service.AutodiscoverUrl($aceuser.mail.ToString(),{$true})

$TotalItemCount = 0
$TotalItemSize = 0

"Checking : " + $MailboxName 
$folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)
$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(10000)
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;
$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
$PR_DELETED_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26267,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
$PR_DELETED_MSG_COUNT = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26176,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
$psPropertySet.Add($PR_MESSAGE_SIZE_EXTENDED);
$psPropertySet.Add($PR_DELETED_MESSAGE_SIZE_EXTENDED);
$psPropertySet.Add($PR_DELETED_MSG_COUNT);
$fvFolderView.PropertySet = $psPropertySet;
##$sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, $folderItemType)
##$fiResult = $Service.FindFolders($folderidcnt,$sfSearchFilter,$fvFolderView)
$fiResult = $Service.FindFolders($folderidcnt,$fvFolderView)
foreach($ffFolder in $fiResult.Folders){
    $TotalItemCount =  $TotalItemCount + $ffFolder.TotalCount;
    $FolderSize = $null;
    if ($ffFolder.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED,[ref] $FolderSize))
    {
	$TotalItemSize = $TotalItemSize + [Int64]$FolderSize
    }
    $DeletedItemFolderSize = $null;
    if ($ffFolder.TryGetProperty($PR_DELETED_MESSAGE_SIZE_EXTENDED, [ref] $DeletedItemFolderSize))
    {
	 $TotalDeletedItemSize = $TotalDeletedItemSize + [Int64]$DeletedItemFolderSize         
    }
    $DeletedMsgCount = $null;
    if ($ffFolder.TryGetProperty($PR_DELETED_MSG_COUNT, [ref] $DeletedMsgCount))
   {
	 $TotalDeletedItemCount = $TotalDeletedItemCount + [Int32]$DeletedMsgCount;
   }
}
$rptobj = "" | select DisplayName,LegacyDN,TotalItemSize,TotalItemCount,TotalDeletedItemSize,TotalDeletedItemCount 
$rptobj.DisplayName = $_.DisplayName
$rptobj.LegacyDN = $_.LegacyExchangeDN
$rptobj.TotalItemCount = $TotalItemCount
$rptobj.TotalItemSize  = $TotalItemSize
$rptobj.TotalDeletedItemSize = $TotalDeletedItemSize
$rptobj.TotalDeletedItemCount = $TotalDeletedItemCount
$rptCollection += $rptobj

$rptCollection   
