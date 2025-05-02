$exportFolder = "c:\temp\"
$Mailboxes = Get-User -RecipientTypeDetails UserMailbox	
foreach($Mailbox in $Mailboxes){
	try{	
		
		$DisplayName = $Mailbox.DisplayName;
		write-host("Processing " + $DisplayName)
		$fileName =  $exportFolder + $DisplayName + "-" + [Guid]::NewGuid().ToString() + ".vcf"
		add-content -path $filename "BEGIN:VCARD"
		add-content -path $filename "VERSION:2.1"
		$givenName = $Mailbox.FirstName
		$surname = $Mailbox.LastName
		add-content -path $filename ("N:" + $surname + ";" + $givenName)
		add-content -path $filename ("FN:" + $Mailbox.DisplayName)
		$Department = $Mailbox.Department;
		add-content -path $filename ("EMAIL;PREF;INTERNET:" + $Mailbox.WindowsEmailAddress)
		$CompanyName = $Mailbox.Company
		add-content -path $filename ("ORG:" + $CompanyName + ";" + $Department)	
		if($Mailbox.Title -ne ""){
			add-content -path $filename ("TITLE:" + $Mailbox.Title)
		}
		$Country = ""
		$City = ""
		$Street = ""
		$State = ""
		$PCode = ""
		if($Mailbox.City -ne ""){
			$City = $Mailbox.City
		}
		if($Mailbox.StateOrProvince -ne ""){
			$State = $Mailbox.StateOrProvince
		}
		if($Mailbox.StreetAddress -ne ""){
			$Street = $Mailbox.StreetAddress
		}
		if($Mailbox.CountryOrRegion -ne ""){
			$Country = $Mailbox.CountryOrRegion
		}
		if($Mailbox.PostalCode -ne ""){
			$PCode = $Mailbox.PostalCode
		}
		$addr =  "ADR;WORK;PREF:;" + $Country + ";" + $Street + ";" +$City + ";" + $State + ";" + $PCode + ";" + $Country
		add-content -path $filename $addr
		if($Mailbox.MobilePhone -ne ""){
			add-content -path $filename ("TEL;CELL;VOICE:" + $Mailbox.MobilePhone)					
		}
		if($Mailbox.Phone -ne ""){
			add-content -path $filename ("TEL;WORK;VOICE:" + $Mailbox.Phone)
		}
		if($Mailbox.Fax -ne ""){
			add-content -path $filename ("TEL;WORK;FAX:" + $Mailbox.Fax)
		}
		if($Mailbox.HomePhone -ne ""){
			add-content -path $filename ("TEL;HOME;VOICE:" + $Mailbox.HomePhone)
		}
		if($Mailbox.WebPage -ne ""){
			add-content -path $filename ("URL;WORK:" + $Mailbox.WebPage)
		}
		Try{
				$sidbind = "LDAP://<SID=" + $Mailbox.Sid + ">"
				$userObj = [ADSI]$sidbind
				$photo = $userObj.thumbnailPhoto.value
				if($photo -eq $null){"No Photo"}
				if($photo -ne $null){
					add-content -path $filename "PHOTO;ENCODING=BASE64;TYPE=JPEG:"
					$ImageString = [System.Convert]::ToBase64String($photo,[System.Base64FormattingOptions]::InsertLineBreaks)
					add-content -path $filename $ImageString
					add-content -path $filename "`r`n"	
				}
		}
		catch{
		
		}	
		add-content -path $filename "END:VCARD"
		"Exported " + $filename 
	}
	catch{
	}
}
