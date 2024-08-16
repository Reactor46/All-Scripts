$root = [ADSI]'LDAP://RootDSE' 
$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
$dfRoot = [ADSI]$dfDefaultRootPath
$gfGALQueryFilter =  "(&(&(&(& (mailnickname=*)(objectCategory=person)(objectClass=user)(msExchHomeServerName=*)))))"
$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
$dfsearcher.PageSize = 900 
$dfsearcher.Filter = $gfGALQueryFilter
$srSearchResult = $dfsearcher.FindAll()
$mbcount = 0
$upall = 0
$skippall = 0
foreach ($emResult in $srSearchResult) {
	$mbcount++
	if (($mbcount % 100) -eq 0 ) {$mbcount.ToString() + " Mailboxes Processed"}
	$uoUserobject = New-Object System.DirectoryServices.directoryentry
	$uoUserobject = $emResult.GetDirectoryEntry()
	$raReplayAddress = ""
	foreach($maMailAddress in $uoUserobject.Proxyaddresses){
		if ($maMailAddress.indexofany("SMTP:") -eq 0){
			$raReplayAddress = $maMailAddress.ToString().Replace("SMTP:","")
		}	
	}
	if ($uoUserobject.mail.value -ne $null -band $raReplayAddress -ne ""){
		if($uoUserobject.mail.Value.ToLower() -ne $raReplayAddress.ToLower()){
		"MissMatch " + $uoUserobject.mail.Value.ToLower() + "	" + $raReplayAddress.ToLower()
		if ($upall -eq 0 -band $skippall -eq 0){
			$answer = Read-Host "Do you want to modify this Object [Y] Yes [A] Yes to All [N] No [L] No to all "
			switch ($Answer)
			{
				"Y" { 	if ($raReplayAddress -ne ""){
						$uoUserobject.mail.Value = $raReplayAddress
						$uoUserobject.SetInfo()
						"Address Updated"}
					else {"Proxy Address Null !! not updating"}
					}
				"" { "Not updating"}
				"A" { 	if ($raReplayAddress -ne ""){
						$uoUserobject.mail.Value = $raReplayAddress
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
					if ($raReplayAddress -ne ""){
					$uoUserobject.mail.Value = $raReplayAddress
					$uoUserobject.SetInfo()
					"Address Updated"}
					else {"Proxy Address Null !! not updating"}
					}			
					}
		}
	}
	else{
		if ($uoUserobject.mail.value -eq $null){"**** Null Ad Mail Property : " + $uoUserobject.name}
		if ($raReplayAddress -eq ""){"***** Null Proxyaddress : " + $uoUserobject.name}
			
	}
}

"Total number of Mailboxes Processed :" + $mbcount
