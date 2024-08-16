$servername = "server"
[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")
$ewc = new-object EWSUtil.EWSConnection("user@domamain.com",$false, "username", "password", "domain", "https://" + $servername + "/EWS/Exchange.asmx")
$drDuration = new-object EWSUtil.EWS.Duration
[EWSUtil.EWS.DistinguishedFolderIdType] $dType = new-object EWSUtil.EWS.DistinguishedFolderIdType
$dType.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::inbox
$qfolder = new-object EWSUtil.QueryFolder($ewc,$dType, $null, $true,$false)
$wfeed = new-object EWSUtil.WriteFeed($ewc.esb,$ewc.emEmailAddress, "c:\inboxfeed3.xml", "Inbox-Feed", $servername, $qfolder.fiFolderItems)

