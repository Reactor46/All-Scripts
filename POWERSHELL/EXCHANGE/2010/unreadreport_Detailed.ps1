[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")
$fname = "c:\temp\unreadreport.csv"

$mbcombCollection = @()

get-mailbox | foreach-object{
	$mbcomb = "" | select DisplayName,EmailAddress,Inbox_Number_Unread,Inbox_Unread_LastRecieved,Sent_Items_LastSent
	$mbcomb.DisplayName = $_.DisplayName.ToString()
	$mbcomb.EmailAddress = $_.WindowsEmailAddress.ToString()
	
	"Mailbox : " + $_.WindowsEmailAddress.ToString()
	$mbMailboxEmail = $_.WindowsEmailAddress
	$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$false, $null,$null,$null,$null)


	$drDuration = new-object EWSUtil.EWS.Duration
	$drDuration.StartTime = [DateTime]::UtcNow.AddDays(-356)
	$drDuration.EndTime = [DateTime]::UtcNow

	$dTypeFld = new-object EWSUtil.EWS.DistinguishedFolderIdType
	$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::inbox
	$dTypeFld2 = new-object EWSUtil.EWS.DistinguishedFolderIdType
	$dTypeFld2.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::sentitems

	$mbMailbox = new-object EWSUtil.EWS.EmailAddressType
	$mbMailbox.EmailAddress = $mbMailboxEmail
	$dTypeFld.Mailbox = $mbMailbox
	$dTypeFld2.Mailbox = $mbMailbox

	$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 1
	$fldarry[0] = $dTypeFld
	$msgList = $ewc.FindUnread($fldarry, $drDuration, $null, "")
	$mbcomb.Inbox_Number_Unread = $msgList.Count
	if ($msgList.Count -ne 0){
	        $mbcomb.Inbox_Unread_LastRecieved = $msgList[0].DateTimeSent.ToLocalTime().ToString()
	}

	$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 1
	$fldarry[0] = $dTypeFld2
	$msgList = $ewc.FindItems($fldarry, $drDuration, $null, "")
	if ($msgList.Count -ne 0){
	     $mbcomb.Sent_Items_LastSent = $msgList[0].DateTimeSent.ToLocalTime().ToString()
	}
	$mbcombCollection += $mbcomb
}

$mbcombCollection | export-csv -noTypeInformation $fname 

