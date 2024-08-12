[void][Reflection.Assembly]::LoadFile("C:\projects\EWSUtil\EWSOofUtil\bin\Debug\EWSUtil.dll")

$fnFileName = "c:\flint.csv"
$pfPublicFolderPath = "/PubContacts/SyncContacts"
$casUrl = "https://casservername/ews/exchange.asmx"
$mbMailboxEmail = "mailbox@domain.com"
$ewc = new-object EWSUtil.EWSConnection($mbMailboxEmail,$false, "username", "password12#", "domain",$casUrl)

$parentFolder = new-object EWSUtil.EWS.DistinguishedFolderIdType
$parentFolder.Id = [EWSUtil.EWS.DistinguishedFolderIdNameType]::publicfoldersroot
$fldarry = new-object EWSUtil.EWS.BaseFolderIdType[] 1
$fldarry[0] = $parentFolder
$rfRootFolder = $ewc.getFolder($fldarry)
$rfRootFolder[0].DisplayName

$pfArray = $pfPublicFolderPath.Split("/")
$cfContactsFolder = $rfRootFolder[0]
for ($lint = 1; $lint -lt $pfArray.Length; $lint++) {
	$cfContactsFolder = $ewc.FindSubFolder($cfContactsFolder, $pfArray[$lint]);                     
}

$sySyncProp = new-object EWSUtil.EWS.PathToExtendedFieldType
$sySyncProp.DistinguishedPropertySetId = [EWSUtil.EWS.DistinguishedPropertySetType]::PublicStrings
$sySyncProp.DistinguishedPropertySetIdSpecified = $true
$sySyncProp.PropertyName = "SyncFileName"
$sySyncProp.PropertyType = [EWSUtil.EWS.MapiPropertyTypeType]::String
$biArray = new-object EWSUtil.EWS.BaseFolderIdType[] 1
$biArray[0] = $cfContactsFolder.FolderId 
$contactsHash = $ewc.FindSyncContacts($biArray, $sySyncProp, "filename.csv")
$fileContactsHash = @{ }

import-csv $fnFileName | foreach-object{
	if ($fileContactsHash.ContainsKey($_."E-mail Address") -eq $false){
		$fileContactsHash.Add($_."E-mail Address",1)
	}
	$newcontact = new-object EWSUtil.EWS.ContactItemType
	if ($_."First Name" -ne "") {$newcontact.GivenName = $_."First Name"}
	if ($_."Last Name" -ne "") {$newcontact.Surname = $_."Last Name"}
	if ($_."Middle Name" -ne "") {$newcontact.MiddleName = $_."Middle Name"}
	if ($_."Name" -ne "") {
		$newcontact.Subject = $_."Name"
		$newcontact.FileAs =  $_."Name"
		$newcontact.DisplayName = $_."Name"
	}
	if ($_."Nickname" -ne "") {$newcontact.Nickname = $_."Nickname"}
	$newcontact.EmailAddresses =  new-object EWSUtil.EWS.EmailAddressDictionaryEntryType[] 1
        $newcontact.EmailAddresses[0] = new-object EWSUtil.EWS.EmailAddressDictionaryEntryType
        $newcontact.EmailAddresses[0].Key = [EWSUtil.EWS.EmailAddressKeyType]::EmailAddress1
        $newcontact.EmailAddresses[0].Value = $_."E-mail Address"
	$newcontact.PhoneNumbers = new-object EWSUtil.EWS.PhoneNumberDictionaryEntryType[] 6
        $newcontact.PhoneNumbers[0] = new-object EWSUtil.EWS.PhoneNumberDictionaryEntryType
        $newcontact.PhoneNumbers[0].Key = [EWSUtil.EWS.PhoneNumberKeyType]::HomePhone;
        $newcontact.PhoneNumbers[0].Value = $_."Home Phone"
        $newcontact.PhoneNumbers[1] = new-object EWSUtil.EWS.PhoneNumberDictionaryEntryType
        $newcontact.PhoneNumbers[1].Key = [EWSUtil.EWS.PhoneNumberKeyType]::MobilePhone
        $newcontact.PhoneNumbers[1].Value = $_."Mobile Phone"
        $newcontact.PhoneNumbers[2] = new-object EWSUtil.EWS.PhoneNumberDictionaryEntryType
        $newcontact.PhoneNumbers[2].Key = [EWSUtil.EWS.PhoneNumberKeyType]::BusinessPhone
        $newcontact.PhoneNumbers[2].Value = $_."Business Phone"
        $newcontact.PhoneNumbers[3] = new-object EWSUtil.EWS.PhoneNumberDictionaryEntryType
        $newcontact.PhoneNumbers[3].Key = [EWSUtil.EWS.PhoneNumberKeyType]::BusinessFax
        $newcontact.PhoneNumbers[3].Value = $_."Business Fax"
        $newcontact.PhoneNumbers[4] = new-object EWSUtil.EWS.PhoneNumberDictionaryEntryType
        $newcontact.PhoneNumbers[4].Key = [EWSUtil.EWS.PhoneNumberKeyType]::HomeFax
        $newcontact.PhoneNumbers[4].Value = $_."Home Fax"
        $newcontact.PhoneNumbers[5] = new-object EWSUtil.EWS.PhoneNumberDictionaryEntryType
        $newcontact.PhoneNumbers[5].Key = [EWSUtil.EWS.PhoneNumberKeyType]::Pager
        $newcontact.PhoneNumbers[5].Value = $_."Pager"
	$newcontact.PhysicalAddresses = new-object EWSUtil.EWS.PhysicalAddressDictionaryEntryType[] 2
        $newcontact.PhysicalAddresses[0] = new-object EWSUtil.EWS.PhysicalAddressDictionaryEntryType
	$newcontact.PhysicalAddresses[0].Key = [EWSUtil.EWS.PhysicalAddressKeyType]::Home
        $newcontact.PhysicalAddresses[0].City = $_."Home City"
        $newcontact.PhysicalAddresses[0].CountryOrRegion = $_."Home Country/Region"
        $newcontact.PhysicalAddresses[0].PostalCode = $_."Home Postal Code"
        $newcontact.PhysicalAddresses[0].Street = $_."Home Street"
        $newcontact.PhysicalAddresses[0].State = $_."Home State"
        $newcontact.PhysicalAddresses[1] = new-object EWSUtil.EWS.PhysicalAddressDictionaryEntryType
	$newcontact.PhysicalAddresses[1].Key = [EWSUtil.EWS.PhysicalAddressKeyType]::Business
        $newcontact.PhysicalAddresses[1].City = $_."Business City"
        $newcontact.PhysicalAddresses[1].CountryOrRegion = $_."Business Country/Region"
        $newcontact.PhysicalAddresses[1].PostalCode = $_."Business Postal Code"
        $newcontact.PhysicalAddresses[1].Street = $_."Business Street"
        $newcontact.PhysicalAddresses[1].State = $_."Business State"
	if ($_."Business Web Page" -ne "") {$newcontact.BusinessHomePage = $_."Business Web Page"}
	if ($_."Company" -ne "") {$newcontact.CompanyName = $_."Company"}
	if ($_."Job Title" -ne "") {$newcontact.JobTitle = $_."Job Title"}
	if ($_."Department" -ne "") {$newcontact.Department = $_."Department"}
	if ($_."Office Location" -ne "") {$newcontact.OfficeLocation = $_."Office Location"}


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

	$newcontact.ExtendedProperty = new-object EWSUtil.EWS.ExtendedPropertyType[] 4
        $newcontact.ExtendedProperty[0] = new-object EWSUtil.EWS.ExtendedPropertyType
        $newcontact.ExtendedProperty[0].ExtendedFieldURI = $dnEmailDisplayName1
        $newcontact.ExtendedProperty[0].Item = $_."Name" + " (" + $_."E-mail Address" + ")" 

	$newcontact.ExtendedProperty[1] = new-object EWSUtil.EWS.ExtendedPropertyType
        $newcontact.ExtendedProperty[1].ExtendedFieldURI = $daEmailDisplayName1
        $newcontact.ExtendedProperty[1].Item = $_."E-mail Address"

	$newcontact.ExtendedProperty[2] = new-object EWSUtil.EWS.ExtendedPropertyType
        $newcontact.ExtendedProperty[2].ExtendedFieldURI = $atEmailAddressType1
        $newcontact.ExtendedProperty[2].Item = "SMTP"

	$newcontact.ExtendedProperty[3] = new-object EWSUtil.EWS.ExtendedPropertyType
        $newcontact.ExtendedProperty[3].ExtendedFieldURI = $sySyncProp
        $newcontact.ExtendedProperty[3].Item = $fnFileName

	$_."E-mail Address"
	if ($contactsHash.containskey($_."E-mail Address")){
		"Address Exist Test Sync"
		$dsContact =  $contactsHash[$_."E-mail Address"] 
		$diffs = $ewc.DiffConact($newcontact,$dsContact)
		if ($diffs.Count -ne 0){
			"Number Changes Found " + $diffs.Count
			[VOID]$ewc.UpdateContact($diffs,$dsContact.ItemId)
		}
			
	}else{
		"CreateContact"
		$ewc.AddNewContact($newcontact,$cfContactsFolder.FolderId )
	}

}
$diDelCollection = @()
$delitems = $false
foreach ($key in $contactsHash.Keys){
	if ($fileContactsHash.ContainsKey($key) -eq $false){
		$diDelCollection += $contactsHash[$key]	
		$delitems = $true

	}
}
if ($delitems -eq $true){
	$delbis = new-object EWSUtil.EWS.BaseItemIdType[] $diDelCollection.length
	for($delint=0;$delint -lt [Int]$diDelCollection.length;$delint++){
		$diDelCollection[$delint].Subject.ToString()
		$delbis[$delint] = $diDelCollection[$delint].ItemId}
		 
	$delbis
	$ewc.DeleteItems($delbis,[EWSUtil.EWS.DisposalType]::SoftDelete)
	"Items Deleted"
}
