##########################################################################
#			Author: Vikas Sukhija
#			Reviewer:
#			Date: 04/22/2016
#			Description: Intelligent Mailbox Movement Agent2
#			This will start the completion on Paused Requests
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

$log1 = ".\Logs\resumemove" + "\" + "CompMove_" + $date + $time + "_.log"
$errlog = ".\Logs\resumemove" + "\" + "err_Compmove" + $date + $time + "_.log"

$report = ".\Logs\resumemove" + "\" + "resume_Move" + $date + "_" + $time + "_.csv"	

$smtpserver = "smtpserver"
$from = "MailboxMoveRequests@labtest.com"
$email1 = "vikassukhija@labtest.com"

$region = "*"

$coll = @()


#############get move requests & finish on the basis of location########

$getmove = get-moverequest | where{$_.status -like "AutoSuspended"}

$getmove | foreach-object{

$Alias = $_.Alias
$TargetDb = $_.TargetDatabase

$regn = get-user $Alias

$locality = $regn.CountryOrRegion



if(($locality -like $region) -or ($locality -like $null)){

resume-moverequest $_ 

write-host "Resuming move request for user $Alias" -foregroundcolor green
$date1 = get-date 
Add-content $log1 "$date1 :: Resuming move request for user $Alias"
	
	$mcoll = "" | select Name,Email,Targetdatabase
	$mcoll.Name = $regn.Name
	$mcoll.Email = $regn.WindowsEmailAddress
	$mcoll.Targetdatabase = $TargetDb
	$coll+= $mcoll}

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

$message.Subject = "Mailbox Resume Move requests: $date1"
$smtp.Send($message)

############################################################




		
