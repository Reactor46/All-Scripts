$MailboxName = "mailbox@domain.com"

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)


$subjecttoSearch  = "I hate Spam"
$AQSQuery = "Received:this week AND subject:`"" + $subjecttoSearch + "`""
$MailDate = [system.DateTime]::Now.AddDays(-7) 
$AQSQuery
$deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete
$aptcancelMode = [Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendToNone
$taskmode =  [Microsoft.Exchange.WebServices.Data.AffectedTaskOccurrence]::AllOccurrences
 
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

$rfRootFolderID = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)
$rfRootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$rfRootFolderID)
$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10000);
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$fvFolderView.PropertySet = $Propset
$ffResponse = $rfRootFolder.FindFolders($fvFolderView);

foreach ($ffFolder in $ffResponse.Folders){
	"Checking " + $ffFolder.DisplayName
	$ivview = new-object Microsoft.Exchange.WebServices.Data.ItemView(20000)
	$frFolderResult = $ffFolder.FindItems($AQSQuery,$ivview)
	$Itembatch = [activator]::createinstance(([type]'System.Collections.Generic.List`1').makegenerictype([Microsoft.Exchange.WebServices.Data.ItemId]))
	foreach ($miMailItems in $frFolderResult.Items){
        #Doublic Check Exact Subect match
        if ($miMailItems.Subject -eq $subjecttoSearch){
		   "****** Found" + $miMailItems.Subject
		   $Itembatch.add($miMailItems.Id)
        }
	}
	"Number of Items found in folder : " + $Itembatch.Count 
	if($Itembatch.Count -ne 0){
		"Deleting " + $Itembatch.Count + " Items"
		$DelResponse = $service.DeleteItems($Itembatch,$deleteMode,$aptcancelMode,$taskmode)
		foreach ($dr in $DelResponse) {
          		$dr.Result
		}
        }

}


	

