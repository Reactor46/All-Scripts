[void][Reflection.Assembly]::LoadFile("C:\temp\EWSUtil.dll")

$fnFileName = "c:\flint.csv"
$casUrl = "https://casservername/ews/exchange.asmx"
$mbMailboxEmail = "mailbox@domain.com"
$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$false, "username", "password12#", "domain",$casUrl)

$lcLocalContactFolderid = new-object EWSUtil.EWS.DistinguishedFolderIdType
$lcLocalContactFolderid.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::contacts
$mbMailbox = new-object EWSUtil.EWS.EmailAddressType
$mbMailbox.EmailAddress = $mbMailboxEmail
$lcLocalContactFolderid.Mailbox = $mbMailbox
$lcLocalContacts = $ewc.getfolder($lcLocalContactFolderid)
import-csv $fnFileName | foreach-object{
	$newcontact = new-object EWSUtil.EWS.ContactItemType
	$namearray =  $_.Name.split(" ")
	$newcontact.GivenName = $namearray[0]
	$newcontact.Surname = $namearray[1]
	$newcontact.Subject = $_.Name
	$newcontact.FileAs =  $_.Name
	$newcontact.DisplayName = $_.Name
	$newcontact.EmailAddresses =  new-object EWSUtil.EWS.EmailAddressDictionaryEntryType[] 1
        $newcontact.EmailAddresses[0] = new-object EWSUtil.EWS.EmailAddressDictionaryEntryType
        $newcontact.EmailAddresses[0].Key = [EWSUtil.EWS.EmailAddressKeyType]::EmailAddress1
        $newcontact.EmailAddresses[0].Value = $_.EmailAddress

	$dnEmailDisplayName1 = new-object EWSUtil.EWS.PathToExtendedFieldType
	$dnEmailDisplayName1.DistinguishedPropertySetIdSpecified = $true
	$dnEmailDisplayName1.DistinguishedPropertySetId = [EWSUtil.EWS.DistinguishedPropertySetType]::Address
	$dnEmailDisplayName1.PropertyId = 32896
	$dnEmailDisplayName1.PropertyIdSpecified = $true
	$dnEmailDisplayName1.PropertyType = [EWSUtil.EWS.MapiPropertyTypeType]::String

	$daEmailDisplayName1  = new-object EWSUtil.EWS.PathToExtendedFieldType
	$daEmailDisplayName1.DistinguishedPropertySetIdSpecified = $true
	$daEmailDisplayName1.DistinguishedPropertySetId = [EWSUtil.EWS.DistinguishedPropertySetType]::Address
	$daEmailDisplayName1.PropertyId = 32900
	$daEmailDisplayName1.PropertyIdSpecified = $true;
	$daEmailDisplayName1.PropertyType = [EWSUtil.EWS.MapiPropertyTypeType]::String

	$atEmailAddressType1  = new-object EWSUtil.EWS.PathToExtendedFieldType
	$atEmailAddressType1.DistinguishedPropertySetIdSpecified = $true
	$atEmailAddressType1.DistinguishedPropertySetId = [EWSUtil.EWS.DistinguishedPropertySetType]::Address
	$atEmailAddressType1.PropertyId = 32898
	$atEmailAddressType1.PropertyIdSpecified = $true;
	$atEmailAddressType1.PropertyType = [EWSUtil.EWS.MapiPropertyTypeType]::String

	$newcontact.ExtendedProperty = new-object EWSUtil.EWS.ExtendedPropertyType[] 3
        $newcontact.ExtendedProperty[0] = new-object EWSUtil.EWS.ExtendedPropertyType
        $newcontact.ExtendedProperty[0].ExtendedFieldURI = $dnEmailDisplayName1
        $newcontact.ExtendedProperty[0].Item =  $_.Name + " (" + $_.EmailAddress + ")" 

	$newcontact.ExtendedProperty[1] = new-object EWSUtil.EWS.ExtendedPropertyType
        $newcontact.ExtendedProperty[1].ExtendedFieldURI = $daEmailDisplayName1
        $newcontact.ExtendedProperty[1].Item = $_.EmailAddress

	$newcontact.ExtendedProperty[2] = new-object EWSUtil.EWS.ExtendedPropertyType
        $newcontact.ExtendedProperty[2].ExtendedFieldURI = $atEmailAddressType1
        $newcontact.ExtendedProperty[2].Item = "SMTP"

	$_.EmailAddress

	$ewc.AddNewContact($newcontact,$lcLocalContacts[0].FolderId)

}