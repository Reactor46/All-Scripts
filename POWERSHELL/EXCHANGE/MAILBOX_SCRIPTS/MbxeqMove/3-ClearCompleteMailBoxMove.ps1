##########################################################################
#			Author: Vikas Sukhija
#			Reviewer:
#			Date: 04/22/2016
#			Description: Intelligent Mailbox Movement Agent3
#			This will clear the complted requests
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

$log1 = ".\Logs\ClearRequest" + "\" + "ClearRequest_" + $date + $time + "_.log"
$errlog = ".\Logs\ClearRequest" + "\" + "err_ClearRequest" + $date + $time + "_.log"

$report = ".\Logs\ClearRequest" + "\" + "ClearRequest_Move" + $date + "_" + $time + "_.csv"	

$smtpserver = "smtpserver"
$from = "MailboxMoveRequests@labtest.com"
$email1 = "vikassukhija@labtest.com"

$coll = @()


#############get move requests & finish on the basis of location########

$getmove = get-moverequest | where{$_.status -like "Completed"}

$getmove | foreach-object{

$Alias = $_.Alias
$TargetDb = $_.TargetDatabase

write-host "Clear move request for user $Alias" -foregroundcolor green

Remove-MoveRequest $_ -confirm:$false

$date1 = get-date 
Add-content $log1 "$date1 :: Clear move request for user $Alias"
	
	$mcoll = "" | select Name,Targetdatabase
	$mcoll.Name = $Alias
	$mcoll.Targetdatabase = $TargetDb
	$coll+= $mcoll

}
if($coll){
$coll | export-csv $report -notypeinfo}
Add-content $errlog $error 

For($i = 1; $i -le "10"; $i++){ sleep 1;Write-Progress -Activity "Waiting to finish before sending email" -status "$i" -percentComplete ($i /10*100)}

######################send email report######################

$errlog1= get-item $errlog
$report1 = get-item $report -ea silentlycontinue

$date1=get-date
$message = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpserver)
 
$message.From = $from
$message.To.Add($email1)
$message.IsBodyHtml = $False

$attach1 = new-object Net.Mail.Attachment($errlog1.Fullname) 
$message.Attachments.Add($attach1)

if($report1){
$attach2 = new-object Net.Mail.Attachment($report1.Fullname) 
$message.Attachments.Add($attach2)}

$message.Subject = "Completion Mailbox Clear Move requests: $date1"
$smtp.Send($message)

############################################################




		
