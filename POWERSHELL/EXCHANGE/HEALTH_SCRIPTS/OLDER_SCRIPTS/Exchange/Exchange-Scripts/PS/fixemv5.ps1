$upall = 0
$skippall = 0
$mbcount = 0
Get-Mailbox -ResultSize unlimited | ForEach-Object{
	$mbcount++
	if (($mbcount % 100) -eq 0 ) {$mbcount.ToString() + " Mailboxes Processed"}
	$usrdn = $_.DistinguishedName.ToString()
	$raReplayAddress = $null
	foreach($maMailAddress in $_.EmailAddresses){
		if ($maMailAddress.IsPrimaryAddress -eq $true -band $maMailAddress.PrefixString -eq "SMTP"){
			
			$raReplayAddress = $maMailAddress.SmtpAddress
		}	
	}
	if ($_.WindowsEmailAddress.ToString() -ne $null -band $raReplayAddress -ne $null){
		if($_.WindowsEmailAddress.ToString().ToLower() -ne $raReplayAddress.ToLower()){
		"MissMatch " + $_.WindowsEmailAddress.ToString().ToLower() + "	" + $raReplayAddress.ToLower()
		if ($upall -eq 0 -band $skippall -eq 0){
			$answer = Read-Host "Do you want to modify this Object [Y] Yes [A] Yes to All [N] No [L] No to all "
			switch ($Answer)
			{
				"Y" { 	if ($raReplayAddress -ne $null){
						$usrDNLD = "LDAP://" + $usrdn 
						$uoUserobject = [ADSI]$usrDNLD 
						$uoUserobject.mail = $raReplayAddress
						$uoUserobject.SetInfo()
						"Address Updated"}
					else {"Proxy Address Null !! not updating"}
					}
				"" { "Not updating"}
				"A" { 	if ($raReplayAddress -ne $null){
						$usrDNLD = "LDAP://" + $usrdn 
						$uoUserobject = [ADSI]$usrDNLD 
						$uoUserobject.mail = $raReplayAddress
						$uoUserobject.SetInfo()
						$upall = 1
						"Address Updated"}
					else {"Proxy Address Null !! not updating"}
					}
				"N" {"Not updating" }
				"L" {"Not updating" 
				     $skippall = 1
					}
			}
			
			}
		else {if ($upall -eq 1 -band $skippall -eq 0){	
					if ($raReplayAddress -ne $null){
						$usrDNLD = "LDAP://" + $usrdn 
						$uoUserobject = [ADSI]$usrDNLD 
						$uoUserobject.mail = $raReplayAddress
						$uoUserobject.SetInfo()
						"Address Updated"}
					else {"Proxy Address Null !! not updating"}
					}			
					}
		}
	}
	else{
		if ($_.WindowsEmailAddress.ToString() -eq ""){"**** Null Ad Mail Property : " + $_.name}
		if ($raReplayAddress -eq $null){"***** No Primary SMTP Proxyaddress : " + $_.name}
			
	}

}
"Total number of Mailboxes Processed :" + $mbcount