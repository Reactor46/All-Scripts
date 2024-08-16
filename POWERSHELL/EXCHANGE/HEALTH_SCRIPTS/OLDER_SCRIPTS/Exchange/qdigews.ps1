[void][Reflection.Assembly]::LoadFile("C:\EWSUtil.dll")

$mbMailboxEmail = "quantinemailbox@domain.com"
$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$false, $null,$null,$null,$null)

$drDuration = new-object EWSUtil.EWS.Duration
$drDuration.StartTime = [DateTime]::UtcNow.AddDays(-1)
$drDuration.EndTime = [DateTime]::UtcNow

$dTypeFld = new-object EWSUtil.EWS.DistinguishedFolderIdType
$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::inbox

$mbMailbox = new-object EWSUtil.EWS.EmailAddressType
$mbMailbox.EmailAddress = $mbMailboxEmail
$dTypeFld.Mailbox = $mbMailbox

$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 1
$fldarry[0] = $dTypeFld
$NDRList = $ewc.GetNDRs($fldarry, $drDuration)

$rpReport = $rpReport + "<table><tr bgcolor=`"#95aedc`">" +"`r`n"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:15%;`" ><b>Recieved</b></td>" +"`r`n"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:20%;`" ><b>From</b></td>" +"`r`n"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:20%;`" ><b>To</b></td>" +"`r`n"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:40%;`" ><b>Subject</b></td>" +"`r`n"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:5%;`" ><b>Size</b></td>" +"`r`n"
$rpReport = $rpReport + "</tr>" + "`r`n"
foreach ($message in $NDRList){
	if ($message.ExtendedProperty[0].Item.ToString() -ne "SMTP"){
		$fromstring = $message.ExtendedProperty[2].Item.ToString()}
	else{
		$fromstring = $message.ExtendedProperty[1].Item.ToString()
	}
	$Oulookid = $ewc.convertid($message.ItemId,[EWSUtil.EWS.IdFormatType]::HexEntryId)
	if ($fromstring.length -gt 30){$fromstring = $fromstring.Substring(0,30)}
	if ($message.ExtendedProperty[3].Item.ToString().length -gt 30){$Tostring = $message.ExtendedProperty[3].Item.ToString().Substring(0,30)}
						 else{$Tostring = $message.ExtendedProperty[3].Item.ToString()}
	$rpReport = $rpReport + "  <tr>"  + "`r`n"
	$rpReport = $rpReport + "<td>" + $message.DateTimeReceived.ToString() + "</td>"  + "`r`n"
	$rpReport = $rpReport + "<td>" +  $fromstring + "</td>"  + "`r`n"
	$rpReport = $rpReport + "<td>" + $Tostring + "</td>"  + "`r`n"
	$rpReport = $rpReport + "<td><a href=`"outlook:" + $Oulookid + "`">" + $message.ExtendedProperty[4].Item.ToString() + "</td>"  + "`r`n"
	$rpReport = $rpReport + "<td>" +  ($message.Size/1024).ToString(0.00) + "</td>"  + "`r`n"
	$rpReport = $rpReport + "</tr>"  + "`r`n"
}
$rpReport = $rpReport + "</table>"  + "  " 
$mrMailRecp = new-object EWSUtil.EWS.EmailAddressType
$mrMailRecp.EmailAddress = "user@domain.com"
$raRecpArray = new-object EWSUtil.EWS.EmailAddressType[] 1
$raRecpArray[0] = $mrMailRecp
$ewc.SendMessage($raRecpArray,"Quarantine Digest",$rpReport)