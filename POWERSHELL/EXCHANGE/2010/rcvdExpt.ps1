[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")

function getMessage($recpAddress){
$ascii = new-object System.Text.ASCIIEncoding
$ewc = new-object EWSUtil.EWSConnection($recpAddress,$true, "", "", "",$null)
$dType = new-object EWSUtil.EWS.DistinguishedFolderIdType
$dType.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::inbox
$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 1
$fldarry[0] = $dType
$randNumber = New-Object system.random
$prop = new-object EWSUtil.EWS.PathToUnindexedFieldType
$prop.FieldURI = [EWSUtil.EWS.UnindexedFieldURIType]::messageInternetMessageId
$messagelist = $ewc.FindMessage($fldarry,$prop,$_.MessageID,$false)
foreach ($message in $messagelist){
	$baByteArray = [Convert]::FromBase64String($message.MimeContent.Value)
	$emlMessage = $ascii.GetString($baByteArray)
	$emlfile = new-object IO.StreamWriter(("c:\export\message\" +  $message.Subject.Replace("#","").Replace(":","") + $mc+ ".eml"),$true)
	$emlfile.WriteLine($emlMessage)
	$emlfile.Close()
	"Exported Message " + $message.Subject
	$mc = $mc +1
	if ($message.hasattachments){
		"Exported Message to " + ("c:\message\" +  $message.Subject + $mc + ".eml")
			foreach($attach in $message.Attachments){
      	                  $ewc.DownloadAttachment(("c:\export\Attachments\"  + $randNumber.next(1,1000) + $attach.Name.ToString()),$attach.AttachmentId);
      	              	  "Downloaded Attachment : " +  $attach.Name.ToString()
			}
		}
}
}


$servername = "servername"
$DomainHash = @{ }

get-accepteddomain | ForEach-Object{
	if ($_.DomainType -eq "Authoritative"){
		$DomainHash.add($_.DomainName.SmtpDomain.ToString().ToLower(),1)
	}

}

$dtQueryDT = [DateTime]::UtcNow.AddHours(-2)
$dtQueryDTf =  [DateTime]::UtcNow
Get-MessageTrackingLog -Server $servername -ResultSize Unlimited -Start $dtQueryDT -End $dtQueryDTf -EventId "RECEIVE" | where {$_.TotalBytes -gt 5242880} | ForEach-Object{ 
	foreach($recp in $_.recipients){
	if ($recp -ne ""){
		$recparray = $recp.split("@")
		if ($DomainHash.ContainsKey($recparray[1])){
			$vuser = get-user $recp
			getMessage($vuser.WindowsEmailAddress)
		}
	}
}}