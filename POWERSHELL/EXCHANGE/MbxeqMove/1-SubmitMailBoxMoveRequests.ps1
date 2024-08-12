##########################################################################
#			Author: Vikas Sukhija
#			Reviewer:
#			Date: 04/22/2016
#			Description: Intelligent Mailbox Movement Agent
##########################################################################

########Add Exchange 2010 Shell ####################

If ((Get-PSSnapin | where {$_.Name -match "Microsoft.Exchange.Management.PowerShell.E2010"}) -eq $null)
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
}



################Logs#####################

$date = get-date -format d
$time = get-date -format t
$month = get-date 
$month1 = $month.month
$year1 = $month.year


$date = $date.ToString().Replace(“/”, “-”)

$time = $time.ToString().Replace(":", "-")
$time = $time.ToString().Replace(" ", "")

$log1 = ".\Logs\MbxMovereq" + "\" + "Move_" + $date + $time + "_.log"
$errlog = ".\Logs\MbxMovereq" + "\" + "err_subm" + $date + $time + "_.log"

$report = ".\Logs\MbxMovereq" + "\" + "Move_Subm" + $date + "_" + $time + "_.csv"	

$smtpserver = "smtpserver"
$from = "MailboxMoveRequests@labtest.com"
$email1 = "vikassukhija@labtest.com"

####################Variables############


$mbxsrvsrc = "OLDdag01"
$mbxsrvdst = "Newdag01"

$countperday = 10

$coll = @()


#########################getmbx from dbs###################

$getsrcdbmbx = get-mailboxdatabase | where{$_.MasterServerOrAvailabilityGroup -like $mbxsrvsrc} 

$getdstdbmbx = get-mailboxdatabase | where{$_.MasterServerOrAvailabilityGroup -like $mbxsrvdst} 


foreach($db in $getsrcdbmbx){

$mbx += get-mailbox -database $db -resultsize $countperday | where{(($_.MailboxMoveStatus -like "None") -and ($_.Name -notlike "sys*") -and ($_.ProhibitSendQuota -gt "10 MB"))}


$date1=get-date
Write-host "$date1 : Processing only.. $countperday for $db" -foregroundcolor green}


Write-host "Total mailbox extracted for the move from all source dbs is "$mbx.count"" -foregroundcolor magenta
$mbxcount = $mbx.count
$date1=get-date
Add-content $log1 "$date1 :: Total mailbox extracted for the move from $db is $mbxcount"

$dbdstcount=[int]$getdstdbmbx.count

############distribute mailboxes on diffrent databases####################

if($mbxcount -gt "0"){

	$a=0
        for($i=0;$i -lt $dbdstcount;$i++){
        $srcmailbox = $mbx[$a]
	$targetdatabase = $getdstdbmbx[$i]
	New-moverequest $srcmailbox -targetdatabase $targetdatabase -SuspendWhenReadyToComplete
	
	Write-Host "$date1 :: $srcmailbox move request submited to $targetdatabase" -foregroundcolor magenta   
	$date1=get-date
	Add-content $log1 "$date1 :: $srcmailbox move request submited to $targetdatabase"

	$mcoll = "" | select Name,PrimarySmtpaddress,Sourcedatabase,Targetdatabase,UseDatabaseQuotaDefaults,ProhibitSendQuota
	$mcoll.Name = $srcmailbox.Name
	$mcoll.PrimarySmtpaddress = $srcmailbox.PrimarySmtpaddress
	$mcoll.UseDatabaseQuotaDefaults = $srcmailbox.UseDatabaseQuotaDefaults
	$mcoll.ProhibitSendQuota = $srcmailbox.ProhibitSendQuota
	$mcoll.Sourcedatabase = $srcmailbox.database
        $mcoll.Targetdatabase = $getdstdbmbx[$i].Name
	$coll+= $mcoll
        $a = $a +1
        if($i -eq $dbdstcount -1){$i = -1}
	if($a -eq $mbxcount){$i = $dbdstcount} 
        
 	}}

$coll | export-csv $report -notypeinfo

Add-content $errlog $error 

For($i = 1; $i -le "10"; $i++){ sleep 1;Write-Progress -Activity "Waiting to finish before sending email" -status "$i" -percentComplete ($i /10*100)}

######################send email report######################

$errlog1= get-item $errlog
$report1 = get-item $report
$date1=get-date
$message = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpserver)
 
$message.From = $from
$message.To.Add($email1)
$message.IsBodyHtml = $False

$attach1 = new-object Net.Mail.Attachment($errlog1.Fullname) 
$message.Attachments.Add($attach1)

$attach2 = new-object Net.Mail.Attachment($report1.Fullname) 
$message.Attachments.Add($attach2)

$message.Subject = "Mailbox Move requests: $date1"
$smtp.Send($message)

############################################################




		
