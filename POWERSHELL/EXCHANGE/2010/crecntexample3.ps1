### Contact Property List
### 
### Normal : to Set eg SetProp "Normal" "GivenName" "MyfirstName"
### 	http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.contact_members%28v=EXCHG.80%29.aspx
### 
### Email :  to Set eg SetProp "Email" "EmailAddress1.Address" "glenscales@yahoo.com"
###	    EmailAddress1.Address
###         EmailAddress2.Address  
###         EmailAddress3.Address  
###	    EmailAddress1.Name
###         EmailAddress2.Name  
###         EmailAddress3.Name 
###
### Phone  :  to Set eg SetProp SetProp "Phone" "MobilePhone" "2345234523"
###	    AssistantPhone	The assistant's phone number.
###         BusinessFax	The business fax number.
###         BusinessPhone	The business phone number.
###         BusinessPhone2	The second business phone number.
###	    Callback	The callback number.
###	    CarPhone	The car phone number.
###	    CompanyMainPhone	The company's main phone number.
###	    HomeFax	The home fax number.
###	    HomePhone	The home phone number.
###	    HomePhone2	The second home phone number.
###	    Isdn	The ISDN number.
###	    MobilePhone	The mobile phone number.
###	    OtherFax	An alternate fax number.
###	    OtherTelephone	An alternate phone number.
###	    Pager	The pager number.
###	    PrimaryPhone	The primary phone number.
###	    RadioPhone	The radio phone number.
###	    Telex	The Telex number.
###	    TtyTddPhone	The TTY/TTD phone number. 
###
### Address  : to Set eg SetProp SetProp "Address" "Business.City" "Sydney"
###
###          Business.City
###	     Business.CountryOrRegion
###          Business.PostalCode
###	     Business.State
###          Business.Street
###
###          Home.City
###	     Home.CountryOrRegion
###          Home.PostalCode
###	     Home.State
###          Home.Street
###
###          Other.City
###	     Other.CountryOrRegion
###          Other.PostalCode
###	     Other.State
###          Other.Street
### 
### Extended :  to Set
### 
###   $AddressGuid =	new-object Guid("00062004-0000-0000-C000-000000000046")
###   
###   $email1DisplayNameProp = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition($AddressGuid,32896, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
###   SetProp "Extended" $email1DisplayNameProp "Fredoo"
###
###   $gender =  New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(14925,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Short)
###   SetProp "Extended" $gender 2
### 

$MailboxName = "user@domain.com"
$csvFile = "c:\allcustm.csv"

$AddressGuid =	new-object Guid("00062004-0000-0000-C000-000000000046")
$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)

function SetProp([String]$Type,[Object]$Name,[Object]$Value){
	$p1Prop1 = "" | select proptype,name,value
	$p1Prop1.proptype = $Type
	$p1Prop1.name = $Name
	$p1Prop1.value = $Value
	$ContactProps.Add($Name,$p1Prop1)
}

Function FindTargetFolder([String]$FolderPath){
	$tfTargetFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
	$pfArray = $FolderPath.Split("/")
	for ($lint = 1; $lint -lt $pfArray.Length; $lint++) {
		$pfArray[$lint]
		$fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)
		$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$pfArray[$lint])
                $findFolderResults = $service.FindFolders($tfTargetFolder.Id,$SfSearchFilter,$fvFolderView)
		if ($findFolderResults.TotalCount -gt 0){
			foreach($folder in $findFolderResults.Folders){
				$tfTargetFolder = $folder				
			}
		}
		else{
			"Error Folder Not Found"
			$tfTargetFolder = $null
			break
		}	
	}
	$Global:findFolder = $tfTargetFolder
}

function CreateContact($service,$ContactProps,$Folder){

$NewContact = new-object  Microsoft.Exchange.WebServices.Data.Contact($service)

$ContactProps.GetEnumerator() | foreach-object {
	$propName = $_.Value.name
	$propValue = $_.Value.value
	if ($_.Value.proptype -ne "Extended"){
		$psplit = $propName.split(".")
		$pval1 = $psplit[0] 
		$pval2 = $psplit[1]
	}
	Switch($_.Value.proptype){
		"Normal" {$NewContact.$propName = $propValue}
		"Email" {
				if ($NewContact.EmailAddresses.Contains([Microsoft.Exchange.WebServices.Data.EmailAddressKey]::$pval1)){
					$EmailEntry = $NewContact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::$pval1]
				}
				else{
					$EmailEntry = new-object Microsoft.Exchange.WebServices.Data.EmailAddress
				}
				$EmailEntry.$pval2 = $propValue
				$NewContact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::$pval1] = $EmailEntry
			}
		"Phone" {
				$NewContact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::$propName] = $propValue
			}
		"Address"{
				if ($NewContact.PhysicalAddresses.Contains([Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::$pval1)){
					$PhysicalAddressEntry = $NewContact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::$pval1]
				}
				else{
					$PhysicalAddressEntry = new-object Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
				}
				$PhysicalAddressEntry.$pval2 = $propValue
				$NewContact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::$pval1] = $PhysicalAddressEntry
			 }
		"Extended" {
			$NewContact.SetExtendedProperty($propName,$propValue)
		}
			
	
	}
}
$NewContact.Save($Global:findFolder.Id)
"Contact Created : " + $NewContact.FileAs
$Global:newContact = $NewContact
}


$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())


$folderid = FindTargetFolder ("/folder1/folder2")

$ContactProps = @{ }


import-csv $csvFile | foreach-object {
	$ContactProps.Clear()
	SetProp "Normal" "GivenName" $_.FirstName
	SetProp "Normal" "Surname" $_.LastName
	$fileasName = $_.FirstName + "," + $_.LastName
	SetProp "Normal" "Subject" $fileasName
	SetProp "Normal" "FileAs" $fileasName
	SetProp "Normal" "CompanyName" $_.Company
	SetProp "Email" "EmailAddress1.Address" $_.Email
	CreateContact $service $ContactProps 
	
}

