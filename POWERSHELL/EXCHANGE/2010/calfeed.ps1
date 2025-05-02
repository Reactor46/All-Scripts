$servername = "server"
[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")
$ewc = new-object EWSUtil.EWSConnection("user@domamain.com",$false, "username", "password", "domain", "https://" + $servername + "/EWS/Exchange.asmx")
$drDuration = new-object EWSUtil.EWS.Duration
$drDuration.StartTime = [DateTime]::UtcNow
$drDuration.EndTime = [DateTime]::UtcNow.AddDays(-30)
[EWSUtil.EWS.DistinguishedFolderIdType] $dType = new-object EWSUtil.EWS.DistinguishedFolderIdType
$dType.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::calendar
$qfolder = new-object EWSUtil.QueryFolder($ewc,$dType, $drDuration,$false,$true)
$wfeed = new-object EWSUtil.WriteFeed($ewc.esb,$ewc.emEmailAddress, "c:\calendarfeed1.xml", "Calendar-Feed", $servername, $qfolder.fiFolderItems)

