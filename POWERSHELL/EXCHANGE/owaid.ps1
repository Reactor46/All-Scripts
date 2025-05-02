[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")
$null = [Reflection.Assembly]::LoadWithPartialName("System.Web")
$pfPublicFolderPath = "\PubContacts"
$mbMailboxEmail = "useremail@domain.com"
$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$false, "", "", "","")
$pfFolder = get-publicFolder -identity $pfPublicFolderPath
$Oulookid = $ewc.convertHexidPublicFolder($pfFolder.Entryid,[EWSUtil.EWS.IdFormatType]::OWAid)
[System.Web.HttpUtility]::UrlDecode($Oulookid)