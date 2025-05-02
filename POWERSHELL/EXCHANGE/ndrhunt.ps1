$MailboxName = $args[0]

function RepeartSearch{
	$frFolderResult = $InboxFolder.FindItems($Sfir,$view)
	$Itemids = [activator]::createinstance(([type]'System.Collections.Generic.List`1').makegenerictype([Microsoft.Exchange.WebServices.Data.ItemId]))
	foreach($ndr in $frFolderResult.items){
		$Itemids.add($ndr.Id)
	}
	if ($frFolderResult.Items.Count -ne 0){
		$Result = $service.DeleteItems($Itemids,[Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete,[Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone,[Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]::AllOccurrences)
		[INT]$Rcount = 0
		foreach ($res in $Result){
			if ($res.Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success){
				$Rcount++
			}
		}
		$Rcount.ToString() + " Items Deleted"
		if ($frFolderResult.Items.Count -eq 1000) {
			RepeartSearch
		}
	}
	else{
		"No Items to delete"
	}

}

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
$InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$Sfir = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, "REPORT.IPM.Note.NDR")
$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(1000)
$frFolderResult = $InboxFolder.FindItems($Sfir,$view)

$Itemids = [activator]::createinstance(([type]'System.Collections.Generic.List`1').makegenerictype([Microsoft.Exchange.WebServices.Data.ItemId]))
foreach($ndr in $frFolderResult.items){
	$Itemids.add($ndr.Id)
}
if ($frFolderResult.Items.Count -ne 0){
	$Result = $service.DeleteItems($Itemids,[Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete,[Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone,[Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]::AllOccurrences)
	[INT]$Rcount = 0
	foreach ($res in $Result){
		if ($res.Result -eq [Microsoft.Exchange.WebServices.Data.ServiceResult]::Success){
			$Rcount++
		}
	}
	$Rcount.ToString() + " Items Deleted"
	if ($frFolderResult.Items.Count -eq 1000) {
		RepeartSearch
	}
}
else{
	"No Items to delete"
}
