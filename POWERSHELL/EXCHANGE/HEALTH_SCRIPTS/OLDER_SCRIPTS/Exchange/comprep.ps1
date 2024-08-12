$snServerName = "servername"
$fname = "c:\mbreport.csv"
$usrquotas = @{ }
$mstoresquotas = @{ }
$mbcombCollection = @()

get-mailboxdatabase -server $snServerName | ForEach-Object{
	if ($_.ProhibitSendReceiveQuota.IsUnlimited -ne $true){
		$mstoresquotas.add($_.identity,$_.ProhibitSendReceiveQuota)
	}

}

$usrquotas = @{ }
Get-Mailbox -server $snServerName -ResultSize Unlimited | foreach-object{
	if($_.ProhibitSendReceiveQuota.IsUnlimited -ne $true){
		$usrquotas.add($_.ExchangeGuid,$_.ProhibitSendReceiveQuota)
	}
}
$quQuotaval = 0
get-mailboxstatistics -Server $snServerName  | Where {$_.DisconnectDate -eq $null} | ForEach-Object{
$quQuota = "0"
if ($usrquotas.ContainsKey($_.MailboxGUID)){
	if ($usrquotas[$_.MailboxGUID].Value -ne $null){
		$quQuotaval = $usrquotas[$_.MailboxGUID].Value.ToMB()
		$quQuota = "{0:P0}" -f ($_.TotalItemSize.Value.ToMB()/$usrquotas[$_.MailboxGUID].Value.ToMB())}
}
else{
	if ($mstoresquotas.ContainsKey($_.database)){
		if ($mstoresquotas[$_.database].Value -ne $null){
			$quQuotaval = $mstoresquotas[$_.database].Value.ToMB()
			$quQuota = "{0:P0}" -f ($_.TotalItemSize.Value.ToMB()/$mstoresquotas[$_.database].Value.ToMB())}}
}
$icount = 0
$tisize = 0
$disize = 0
if ($_.DisplayName -ne $null){$dname = $_.DisplayName}
if ($_.ItemCount -ne $null){$icount = $_.ItemCount}
if ($_.TotalItemSize.Value.ToMB() -ne $null){$tisize = $_.TotalItemSize.Value.ToMB()}
if ($_.TotalDeletedItemSize.Value.ToMB() -ne $null){$disize = $_.TotalDeletedItemSize.Value.ToMB()}  
$mbcomb = "" | select DisplayName,QuotaSize,TotalItemSize,Department,LastLogonTime
$mbcomb.DisplayName = $dname
$mbcomb.QuotaSize = $quQuotaval
$mbcomb.TotalItemSize = $tisize
$usrString = 'LDAP://' + $_.identity
$usr = [ADSI]$usrString
$mbcomb.Department = $usr.Department
$mbcomb.LastLogonTime = $_.LastLogonTime

$mbcombCollection += $mbcomb
}

$mbcombCollection | export-csv -noTypeInformation $fname 