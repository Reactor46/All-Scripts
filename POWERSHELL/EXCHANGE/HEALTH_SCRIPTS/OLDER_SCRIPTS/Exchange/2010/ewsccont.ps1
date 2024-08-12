#Required Props
$subject = "Skippy(Skippys Email)"
$EmailDisplayAS = "Skippy Roo(Skippys Emaill)"
$EmailDisplayName =  "skippy@acme.com"
$EmailAddress1 = "skippy@acme.com"
$FileAs = "Roo, Skippy"

#Optional Props

$givenName = "Skippy"
$middleName = "The"
$Surname = "Roo"
$nickName = "Skip"
$CompanyName = "Skips Helicopters"
$HomeStreet = "32 Walaby Street"
$HomeCity = "Maroobra"
$HomeState = "NSW"
$HomeCountry = "Australia"
$HomePostCode = "2000"
$BusinessStreet = "45 Emu Place"
$BusinessCity = "Woolloomooloo"
$BusinessState = "NSW"
$BusinessCountry = "Austrlia"
$BusinessPostCode = "2000"
$HomePhone = "9999-99999"
$HomeFax = "9999-99999"
$BusinessPhone = "9999-99999"
$BusinessFax = "9999-99999"
$MobilePhone = "999-9999-99999"
$AssistantName = "Wally"
$BusinessWebPage = "www.skipshelis.com"
$Department = "Flying"
$OfficeLocation = "1st Floor"
$Profession = "Pilot"

$soapString = "<?xml version=`"1.0`" encoding=`"utf-8`" ?> ` 
<soap:Envelope xmlns:soap=`"http://schemas.xmlsoap.org/soap/envelope/`" xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`"  `
xmlns:xsd=`"http://www.w3.org/2001/XMLSchema`">`
<soap:Header>`
		<RequestServerVersion Version=`"Exchange2007_SP1`" xmlns=`"http://schemas.microsoft.com/exchange/services/2006/types`" /> `
	</soap:Header>`
<soap:Body> `
	<CreateItem MessageDisposition=`"SaveOnly`" xmlns=`"http://schemas.microsoft.com/exchange/services/2006/messages`">`
	<SavedItemFolderId> `
		<DistinguishedFolderId Id=`"contacts`" xmlns=`"http://schemas.microsoft.com/exchange/services/2006/types`" />` 
	</SavedItemFolderId>`
	<Items>`
	<Contact xmlns=`"http://schemas.microsoft.com/exchange/services/2006/types`"> `
	<Subject>" + $subject + "</Subject> `
	<ExtendedProperty>`
	<ExtendedFieldURI PropertySetId=`"00062004-0000-0000-C000-000000000046`" PropertyId=`"32898`" PropertyType=`"String`" /> `
	<Value>SMTP</Value> `
	</ExtendedProperty>`
	<ExtendedProperty>`
	<ExtendedFieldURI PropertySetId=`"00062004-0000-0000-C000-000000000046`" PropertyId=`"32896`" PropertyType=`"String`" /> `
	<Value>" + $EmailDisplayAS + "</Value> `
	</ExtendedProperty>`
	<ExtendedProperty>`
	<ExtendedFieldURI PropertySetId=`"00062004-0000-0000-C000-000000000046`" PropertyId=`"32900`" PropertyType=`"String`" /> `
	<Value>" + $EmailDisplayName + "</Value> `
	</ExtendedProperty>`
	<FileAs>" + $fileas + "</FileAs> `
	<GivenName>" + $givenName + "</GivenName> `
	<MiddleName>" + $middleName + "</MiddleName> `
	<Nickname>" + $nickName + "</Nickname> `
	<CompanyName>" + $CompanyName + "</CompanyName> `
	<EmailAddresses>`
	<Entry Key=`"EmailAddress1`">" + $EmailAddress1 + "</Entry> `
	</EmailAddresses>`
	<PhysicalAddresses>`
		<Entry Key=`"Home`">`
			<Street>" + $HomeStreet + "</Street> `
			<City>" + $HomeCity + "</City> `
			<State>" + $HomeState + "</State> `
			<CountryOrRegion>" + $HomeCountry + "</CountryOrRegion> `
			<PostalCode>" + $HomePostCode + "</PostalCode> `
		</Entry>`
		<Entry Key=`"Business`">
			<Street>" + $BusinessStreet + "</Street> `
			<City>" + $BusinessCity + "</City> `
			<State>" + $BusinessState + "</State> `
			<CountryOrRegion>" + $BusinessCountry + "</CountryOrRegion> `
			<PostalCode>" + $BusinessPostCode + "</PostalCode> `
		</Entry>`
	</PhysicalAddresses>`
	 <PhoneNumbers>`
		<Entry Key=`"HomePhone`">" + $HomePhone + "</Entry> `
		<Entry Key=`"HomeFax`">" + $HomeFax + "</Entry> `
		<Entry Key=`"BusinessPhone`">" + $BusinessPhone + "</Entry> `
		<Entry Key=`"BusinessFax`">" + $BusinessFax + "</Entry> `
		<Entry Key=`"MobilePhone`">" + $MobilePhone + "</Entry> `
	</PhoneNumbers>`
	<AssistantName>" + $AssistantName + "</AssistantName> `
	<BusinessHomePage>" + $BusinessWebPage + "</BusinessHomePage> `
	<Department>" + $Department + "</Department> `
	<OfficeLocation>" + $OfficeLocation + "</OfficeLocation> `
	<Profession>" + $Profession + "</Profession> `
	<Surname>" + $Surname + "</Surname> `
</Contact>`
</Items>`
</CreateItem>`
</soap:Body>`
</soap:Envelope>"

$unUserName = "username"
$psPassword = "password"
$dnDomainName = "domain"
$cdUsrCredentials = new-object System.Net.NetworkCredential($unUserName , $psPassword , $dnDomainName)

$mbURI = "https://servername/EWS/Exchange.asmx"

$WDRequest = [System.Net.WebRequest]::Create($mbURI)
$WDRequest.ContentType = "text/xml"
$WDRequest.Headers.Add("Translate", "F")
$WDRequest.Method = "POST"
$WDRequest.Credentials = $cdUsrCredentials
# $WDRequest.UseDefaultCredentials = $True
$bytes = [System.Text.Encoding]::UTF8.GetBytes($soapString)
$WDRequest.ContentLength = $bytes.Length
$RequestStream = $WDRequest.GetRequestStream()
$RequestStream.Write($bytes, 0, $bytes.Length)
$RequestStream.Close()
$WDResponse = $WDRequest.GetResponse()
$ResponseStream = $WDResponse.GetResponseStream()
$readStream = new-object System.IO.StreamReader $ResponseStream
$textresult = $readStream.ReadToEnd()
if ($textresult.indexofany("ResponseClass=`"Success`"")){
	"Contact Created Sucessfully"
	}
else {
	$textresult
}