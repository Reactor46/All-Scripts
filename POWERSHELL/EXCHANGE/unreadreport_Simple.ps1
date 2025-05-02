[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")
$fname = "c:\temp\unreadreport.csv"

$mbcombCollection = @()

get-mailbox | foreach-object{
	$mbcomb = "" | select DisplayName,EmailAddress,Unread
	$mbcomb.DisplayName = $_.DisplayName.ToString()
	$mbcomb.EmailAddress = $_.WindowsEmailAddress.ToString()
	$mbMailboxEmail = $_.WindowsEmailAddress.ToString()
	"Mailbox : " + $mbMailboxEmail
	$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$false, $null,$null,$null,$null)
	$dTypeFld = new-object EWSUtil.EWS.DistinguishedFolderIdType
	$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::inbox

	$mbMailbox = new-object EWSUtil.EWS.EmailAddressType
	$mbMailbox.EmailAddress = $mbMailboxEmail
	$dTypeFld.Mailbox = $mbMailbox

	$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 1
	$fldarry[0] = $dTypeFld

	$fldList = $ewc.GetFolder($fldarry)
	[EWSUtil.EWS.FolderType]$pfld = [EWSUtil.EWS.FolderType]$fldList[0];
        $mbcomb.Unread = $pfld.UnreadCount
	$mbcombCollection += $mbcomb
}

$mbcombCollection | export-csv -noTypeInformation $fname 

