[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")

$mbMailboxEmail = "user@domain.com.au"
$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$false, $null,$null,$null,$null)

$drDuration = new-object EWSUtil.EWS.Duration
$drDuration.StartTime = [DateTime]::UtcNow.AddDays(-7)
$drDuration.EndTime = [DateTime]::UtcNow

$dTypeFld = new-object EWSUtil.EWS.DistinguishedFolderIdType
$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::inbox

$mbMailbox = new-object EWSUtil.EWS.EmailAddressType
$mbMailbox.EmailAddress = $mbMailboxEmail
$dTypeFld.Mailbox = $mbMailbox

$SenderName = new-object EWSUtil.EWS.PathToExtendedFieldType
$SenderName.PropertyTag = "0x0C1A"
$SenderName.PropertyType = [EWSUtil.EWS.MapiPropertyTypeType]::String

$beAdditionproperteis = new-object EWSUtil.EWS.BasePathToElementType[] 1
$beAdditionproperteis[0] = $SenderName

$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 1
$fldarry[0] = $dTypeFld
$MSGList = $ewc.FindItems($fldarry, $drDuration, $beAdditionproperteis, "")


$rpReport = $rpReport + "<table><tr bgcolor=`"#95aedc`">" +"`r`n"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:15%;`" ><b>Recieved</b></td>" +"`r`n"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:30%;`" ><b>From</b></td>" +"`r`n"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:50%;`" ><b>Subject</b></td>" +"`r`n"
$rpReport = $rpReport + "<td align=`"center`" style=`"width:5%;`" ><b>Size</b></td>" +"`r`n"
$rpReport = $rpReport + "</tr>" + "`r`n"
foreach ($message in $MSGList){
	$Oulookid = $ewc.convertid($message.ItemId,[EWSUtil.EWS.IdFormatType]::HexEntryId)
	$rpReport = $rpReport + "  <tr>"  + "`r`n"
	$rpReport = $rpReport + "<td>" + $message.DateTimeReceived.ToString() + "</td>"  + "`r`n"
	$rpReport = $rpReport + "<td>" +  $message.ExtendedProperty[0].Item.ToString() + "</td>"  + "`r`n"
	$rpReport = $rpReport + "<td><a href=`"outlook:" + $Oulookid + "`">" + $message.Subject.ToString() + "</td>"  + "`r`n"
	$rpReport = $rpReport + "<td>" +  ($message.Size/1024).ToString(0.00) + "</td>"  + "`r`n"
	$rpReport = $rpReport + "</tr>"  + "`r`n"
}
$rpReport = $rpReport + "</table>"  + "  " 
$mrMailRecp = new-object EWSUtil.EWS.EmailAddressType
$mrMailRecp.EmailAddress = "user@domain.com.au"
$raRecpArray = new-object EWSUtil.EWS.EmailAddressType[] 1
$raRecpArray[0] = $mrMailRecp
$ewc.SendMessage($raRecpArray,"Message Digest",$rpReport)