
$Subject = "Daily Export"
$ProcessedFolderPath = "/Folder1/Folder2/etc"
$downloadDirectory = "c:\temp\"

Function FindTargetFolder($FolderPath){
	$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
	$pfArray = $FolderPath.Split("/")
	for ($lint = 1; $lint -lt $pfArray.Length; $lint++) {
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

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

FindTargetFolder($ProcessedFolderPath)


$PublicFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$Global:findFolder.Id.UniqueId)
$Sfir = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, $false)
$Sfsub = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, $Subject)
$Sfha = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfCollection.add($Sfir)
$sfCollection.add($Sfsub)
$sfCollection.add($Sfha)
$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(2000)
$frFolderResult = $PublicFolder.FindItems($sfCollection,$view)
foreach ($miMailItems in $frFolderResult.Items){
	$miMailItems.Subject
	$miMailItems.Load()
	foreach($attach in $miMailItems.Attachments){
	$attach.Load()
		$fiFile = new-object System.IO.FileStream(($downloadDirectory + “\” + $attach.Name.ToString()), [System.IO.FileMode]::Create)
		$fiFile.Write($attach.Content, 0, $attach.Content.Length)
		$fiFile.Close()
		write-host "Downloaded Attachment : " + (($downloadDirectory + “\” + $attach.Name.ToString()))
	}
	$miMailItems.isread = $true
	$miMailItems.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
	[VOID]$miMailItems.Move($Global:findFolder.Id.UniqueId)
}
