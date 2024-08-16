$setPPTo = "0"

[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")
$casUrl = "https://servername/ews/exchange.asmx"
$mbMailboxEmail= "user@domain.com"
$ewc = new-object EWSUtil.EWSConnection("user@domain.com",$false, "username", "Password", "domain",$casUrl)

$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 1
$dTypeFld = new-object EWSUtil.EWS.DistinguishedFolderIdType
$dTypeFld.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::msgfolderroot
$mbMailbox = new-object EWSUtil.EWS.EmailAddressType
$mbMailbox.EmailAddress = $mbMailboxEmail
$dTypeFld.Mailbox = $mbMailbox
$fldarry[0] = $dTypeFld

$FolderList = $ewc.GetAllMailboxFolders($fldarry) 
$fldarry1 = new-object EWSUtil.EWS.BaseFolderIdType[] $FolderList.Count
for ($fcint=0;$fcint -lt $FolderList.Count;$fcint++){
	$fldarry1[$fcint] = $FolderList[$fcint].FolderId
}
$psPreviewSetting = new-object EWSUtil.EWS.PathToExtendedFieldType
$psPreviewSetting.DistinguishedPropertySetIdSpecified = $true
$psPreviewSetting.DistinguishedPropertySetId = [EWSUtil.EWS.DistinguishedPropertySetType]::PublicStrings
$psPreviewSetting.PropertyName = "http://schemas.microsoft.com/exchange/preview"
$psPreviewSetting.PropertyType = [EWSUtil.EWS.MapiPropertyTypeType]::Integer
$beparray = new-object EWSUtil.EWS.BasePathToElementType[] 1
$beparray[0] = $psPreviewSetting 
$Folders = $ewc.GetFolder($fldarry1,$beparray)
If ($Folders.Count -ne 0) {
	 ForEach ($Folder in $Folders) {
		$cval = "n/a"
		if ($Folder.extendedProperty -ne $null){
			$cval = $Folder.extendedProperty[0].Item.ToString()
			switch ($Folder.extendedProperty[0].Item.ToString()){
				0 { $Folder.DisplayName + " Preview Pane Set Off"}
				1 { $Folder.DisplayName + " Preview Pane Set Right"}
				2 { $Folder.DisplayName + " Preview Pane Set Bottom"}
		}			
		}
		else{
			$Folder.DisplayName + " Not set will default to On"
		}
		if ($cval -ne $setPPTo){
			$exProp = new-object EWSUtil.EWS.ExtendedPropertyType
			$exProp.ExtendedFieldURI = $psPreviewSetting
			$exProp.Item = $setPPTo
			$ewc.UpdateFolderExtendedProperty($exProp,$Folder)
		}
	}
}