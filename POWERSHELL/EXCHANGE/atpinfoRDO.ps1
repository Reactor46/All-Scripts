[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") 


$snServerName = $args[0]

$Global:rptCollection = @()

function QueryMailbox($mb){
$mb
$rdoSession = new-object -com Redemption.RDOSession
$rdoSession.LogonExchangeMailbox($SmtpAddress,$snServerName)
$calendar = $rdoSession.Stores.DefaultStore.GetDefaultFolder(9)




$Itmcnt = "" | select DisplayName,SMTPAddress,ItemCount,Size
$Itmcnt.DisplayName = $displayName
$Itmcnt.SMTPAddress = $SmtpAddress
$Itmcnt.ItemCount = $calendar.Items.Count
$Itmcnt.Size = [System.Math]::Round(($calendar.Fields(235405315) /1024),2)
$Global:rptCollection += $Itmcnt
$Itmcnt
$rdoSession.logoff()
}

function GetUsers(){
$root = [ADSI]'LDAP://RootDSE' 
$cfConfigRootpath = "LDAP://" + $root.ConfigurationNamingContext.tostring()
$dfDefaultRootPath = "LDAP://" + $root.DefaultNamingContext.tostring()
$configRoot = [ADSI]$cfConfigRootpath 
$dfRoot = [ADSI]$dfDefaultRootPath
$searcher = new-object System.DirectoryServices.DirectorySearcher($configRoot)
$searcher.Filter = '(&(objectCategory=msExchExchangeServer)(cn=' + $snServerName  + '))'
$searcher.PropertiesToLoad.Add("cn")
$searcher.PropertiesToLoad.Add("gatewayProxy")
$searcher.PropertiesToLoad.Add("legacyExchangeDN")
$searcher1 = $searcher.FindAll()
foreach ($server in $searcher1){ 
	$snServerEntry = New-Object System.DirectoryServices.directoryentry 
        $snServerEntry = $server.GetDirectoryEntry() 
	$snServerName = $snServerEntry.cn
	$snExchangeDN = $snServerEntry.legacyExchangeDN
}
$searcher.Filter = '(&(objectCategory=msExchRecipientPolicy)(cn=Default Policy))'
$searcher1 = $searcher.FindAll()
foreach ($recppolicies in $searcher1){ 
     	$gwaddrrs = New-Object System.DirectoryServices.directoryentry 
        $gwaddrrs = $recppolicies.GetDirectoryEntry() 
	foreach ($address in $gwaddrrs.gatewayProxy){
		if($address.Substring(0,5) -ceq "SMTP:"){$dfAddress = $address.Replace("SMTP:@","")}
	}	
	
}
$arMbRoot = "https://" + $snServerName + "/exadmin/admin/" + $dfAddress + "/mbx/"
$gfGALQueryFilter =  "(&(&(&(& (mailnickname=*)(!msExchHideFromAddressLists=TRUE)(| (&(objectCategory=person)" `
+ "(objectClass=user)(msExchHomeServerName=" + $snExchangeDN + ")) )))))"
$dfsearcher = new-object System.DirectoryServices.DirectorySearcher($dfRoot)
$dfsearcher.Filter = $gfGALQueryFilter
$searcher2 = $dfsearcher.FindAll()
foreach ($uaUsers in $searcher2){ 
     	$uaUserAccount = New-Object System.DirectoryServices.directoryentry 
        $uaUserAccount = $uaUsers.GetDirectoryEntry() 
	foreach ($address in $uaUserAccount.proxyaddresses){
		if($address.Substring(0,5) -ceq "SMTP:"){$uaAddress = $address.Replace("SMTP:","")}
	}
	$SmtpAddress = $uaAddress
	$displayName = $uaUserAccount.DisplayName[0]
	QueryMailbox($uaAddress)
}
}

GetUsers
$Global:rptCollection | Export-Csv -NoTypeInformation c:\mailbox.csv