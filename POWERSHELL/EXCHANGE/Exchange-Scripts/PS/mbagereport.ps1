$Range = "01/01/1990..01/01/2008"
$MailboxName = "user@domain.com"

$AQSString = "System.Message.DateReceived:" + $Range
$rptCollection = @()


function ConvertToString($ipInputString){
	$Val1Text = ""
	for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){
			$Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))
			$clInt++
	}
	return $Val1Text
}


$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.1\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010)
#$service.Credentials = New-Object System.Net.NetworkCredential("username","password")

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind
$service.AutodiscoverUrl($MailboxName,{$true})
$PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);

"Checking : " + $MailboxName 
$folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)
$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;
$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$PR_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
$PR_DELETED_MESSAGE_SIZE_EXTENDED = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26267,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long);
$PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);

$psPropertySet.Add($PR_MESSAGE_SIZE_EXTENDED);
$psPropertySet.Add($PR_Folder_Path);
$fvFolderView.PropertySet = $psPropertySet;
$sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")
$fiResult = $Service.FindFolders($folderidcnt,$sfSearchFilter,$fvFolderView)
foreach($ffFolder in $fiResult.Folders){
    "Processing : " + $ffFolder.displayName
    $TotalItemCount =  $TotalItemCount + $ffFolder.TotalCount;
    $FolderSize = $null;
	$FolderSizeValue = 0
    if ($ffFolder.TryGetProperty($PR_MESSAGE_SIZE_EXTENDED,[ref] $FolderSize))
    {
      	$FolderSizeValue = [Int64]$FolderSize
    }
	$foldpathval = $null
	if ($ffFolder.TryGetProperty($PR_Folder_Path,[ref] $foldpathval))
	{
	
	}
	$binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)
	$hexArr = $binarry | ForEach-Object { $_.ToString("X2") }
    $hexString = $hexArr -join ''
	$hexString = $hexString.Replace("FEFF", "5C00")
	$fpath = ConvertToString($hexString)
    $fiFindItems = $null
    $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
	$psPropertySet1 = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
	$psPropertySet1.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Size)
	$ItemView.PropertySet
	$itemCollection = @()
	do{
		$fiFindItems = $ffFolder.findItems($AQSString,$ItemView)
		$ItemView.offset += $fiFindItems.Items.Count
		foreach($Item in $fiFindItems.Items){
			$rptobject = "" | select Size
			$rptobject.Size = $Item.Size
			$itemCollection +=$rptobject
		}
    }while($fiFindItems.MoreAvailable -eq $true)
	$outObj =  $itemCollection | Measure-Object Size -Sum -Average -Min -Max
	if($outObj -ne $null){
		Add-Member -InputObject $outObj -MemberType NoteProperty -Name Mailbox -Value $MailboxName
		Add-Member -InputObject $outObj NoteProperty -Name Folder -Value $ffFolder.DisplayName
		Add-Member -InputObject $outObj NoteProperty -Name FolderPath -Value $fpath
		Add-Member -InputObject $outObj NoteProperty -Name TotalFolderSize -Value $FolderSizeValue
		Add-Member -InputObject $outObj NoteProperty -Name DateRange -Value $Range
		$rptCollection += $outObj
	}
}
$rptCollection | select Mailbox,Folder,FolderPath,@{label="FolderSize(MB)";expression={[math]::Round($_.TotalFolderSize/1MB,2)}},@{label="RangeSize(MB)";expression={[math]::Round($_.Sum/1MB,2)}},@{label="RangeCount";expression={$_.Count}},@{label="PercentOfSize";expression={'{0:P0}' -f ($_.Sum/$_.TotalFolderSize)}} | export-csv c:\temp\MbAgeReport.csv -NoTypeInformation
