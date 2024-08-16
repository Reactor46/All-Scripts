[void][Reflection.Assembly]::LoadFile("C:\EWSUtil.dll")

$mbMailboxEmail = "quantinemailbox@domain.com"
$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$false, $null,$null,$null,$null)


$drDuration = new-object EWSUtil.EWS.Duration
$drDuration.StartTime = [DateTime]::UtcNow.AddDays(-365)
$drDuration.EndTime = [DateTime]::UtcNow.AddDays(-31)

$dTypeFld = new-object EWSUtil.EWS.DistinguishedFolderIdType
$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::inbox

$mbMailbox = new-object EWSUtil.EWS.EmailAddressType
$mbMailbox.EmailAddress = $mbMailboxEmail
$dTypeFld.Mailbox = $mbMailbox


$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 1
$fldarry[0] = $dTypeFld
$msgList = $ewc.FindItems($fldarry, $drDuration, $null, "")
$batchsize = 100
$bcount = 0
if ($msgList.Count -ne 0){
	$itarry = new-object EWSUtil.EWS.ItemIdType[] $batchsize
	for($ic=0;$ic -lt $msgList.Count;$ic++){
		if ($bcount -ne $batchsize){
			$itarry[$bcount] = $msgList[$ic].ItemId
			$bcount++
		}
		else{
			$ewc.deleteItems($itarry,[EWSUtil.EWS.DisposalType]::SoftDelete)
			$itarry = $null
			$itarry = new-object EWSUtil.EWS.ItemIdType[] $batchsize
			$bcount = 0
			$itarry[$bcount] = $msgList[$ic].ItemId
			$bcount++
		}
          
	} 
	$ewc.deleteItems($itarry,[EWSUtil.EWS.DisposalType]::SoftDelete)
}
