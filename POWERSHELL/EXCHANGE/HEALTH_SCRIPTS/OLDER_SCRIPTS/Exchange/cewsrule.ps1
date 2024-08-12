$mailbox = "user@domain.com"
$fldname = "\\Inbox\Sales"

Function FindTargetFolder($FolderPath){
	$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$mailbox)
	$pfArray = $FolderPath.Split("\")
	for ($lint = 2; $lint -lt $pfArray.Length; $lint++) {
		$pfArray[$lint]
		$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)
		$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$pfArray[$lint])
                $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView)
		if ($findFolderResults.TotalCount -gt 0){
			foreach($folder in $findFolderResults.Folders){
				$tfTargetFolder = $folder				
			}
		}
		else{
			"Error Folder Not Found"
			$tfTargetFolder = $null
			break
		}	
	}
	$Global:findFolder = $tfTargetFolder
}

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$service.TraceEnabled = $false

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind
$service.autodiscoverurl($mailbox,{$true})

$irInboxRules = service.GetInboxRules

$nrNewRule = New-Object Microsoft.Exchange.WebServices.Data.Rule
$nrNewRule.DisplayName = "Move all Mails with Sales Reports in Subject"
$nrNewRule.Conditions.ContainsSubjectStrings.Add("Sales Reports");

FindTargetFolder($fldname)

$nrNewRule.Actions.MoveToFolder = $Global:findFolder.Id.UniqueId

$RuleOperation = New-Object Microsoft.Exchange.WebServices.Data.createRuleOperation[] 1
$RuleOperation[0] = $nrNewRule

$service.UpdateInboxRules($RuleOperation,$irInboxRules.OutlookRuleBlobExists)

