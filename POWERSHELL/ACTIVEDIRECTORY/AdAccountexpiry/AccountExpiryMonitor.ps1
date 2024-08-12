##########################################################################
##           Script to Monitor Account expiry                          
##           Author: Vikas Sukhija                  		 
##           Date: 04-30-2014                       		 
##   This scripts is used for monitoring service accounts expiry & 
##   sending email to teams before 15 days till account is unexpired 
##                                              
##########################################################################
##########################Define Variables################################

$date1 = get-date -format d
$date1 = $date1.ToString().Replace("/","-")

$logs = ".\Logs" + "\" + "Processed_" + $date1 + "_.log"

Start-Transcript -Path $logs

$date= get-date
$smtpserver = "smtpserver"
$from = "AccountExpiry@labtest.com"
$days = "-15"
$errormail = "vikassukhija@labtest.com"

If ((Get-PSSnapin | where {$_.Name -match "Quest.ActiveRoles.ADManagement"}) -eq $null)
{
	Add-PSSnapin Quest.ActiveRoles.ADManagement
}

$data=import-csv .\Accountexpiry.csv 

foreach($i in $data) {

$user = $i.account
$acc= Get-QADUser $user | select Name,AccountExpires
$nm= $acc.name


$accexpiry=$acc.AccountExpires

write-host "$nm - $accexpiry" -foregroundcolor yellow

if($accexpiry -eq $null)
{ 
write-host "Account expiration date is not set for $user" -foregroundcolor Green

}

else

{
$accexpiry1 = ($accexpiry).adddays($days)

if($accexpiry1 -le $date)

{

write-host "Account $user will expire on $accexpiry" -foregroundcolor red
$to = $i.Emailid

$message = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpserver)
$message.From = $from
$message.To.Add($to)
$message.bcc.ADD($errormail)
$message.IsBodyHtml = $False
$message.Subject = "Attention: Account $user will expire on $accexpiry"
$smtp.Send($message)
Write-host "Message Sent to $to: $to for Account $user" -foregroundcolor Blue


}

}

}

if ($error -ne $null)
      {
#SMTP Relay address
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)

#Mail sender
$msg.From = $from

#mail recipient
$msg.To.Add($errormail)
$msg.Subject = "Account expiry Script error"
$msg.Body = $error
$smtp.Send($msg)
$error.clear()
       }
  else

      {
    Write-host "no errors till now"
      }

Stop-Transcript

##############################################################################