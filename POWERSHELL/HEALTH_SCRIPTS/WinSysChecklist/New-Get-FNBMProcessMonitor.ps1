#this is two separate scripts which must be run in Tandem all files need to be in the same directory

################################ Script 1. ############################################

###################### Monitor Script ###########################
# By Ryan Jones                                                 #
# This is the script which will run the monitor script          #
# You can call this what you would like I call it Run This.ps1  #
# Change details (if required) anywhere you see **CHANGE THIS** #
#################################################################
 $date = Get-Date
 $d = $date.day
 $m = $date.month
 $y = $date.year
#**CHANGE THIS** Set the location the report is saved to (must exist)
$File ="C:\inetpub\wwwroot\DataLayerService\Reports\report_$d.$m.$y.txt"

"Running Reports"
.\monitor.ps1 | out-file -filepath $File -append

"Sending Email"
#**CHANGE THIS** change these to what you require
$EmailFrom = "ContosoDLSMONITOR@creditone.com"
$EmailTo = "John.Battista@CreditOne.com"
$EmailSubject = "Contoso Data Layer Service Memory Usage" 
$emailbody = @"
               Contoso Data Layer Service Memory Usage Report

Attached is the daily Contoso Service report for the $d/$m/$y please check and review issues.

"@
#**CHANGE THIS** change this to your SMTP server 
$SMTPServer = "mailgateway.Contoso.corp"

$emailattachment = $file

function send_email {
$mailmessage = New-Object system.net.mail.mailmessage
$mailmessage.from = ($emailfrom)
$mailmessage.To.add($emailto)
$mailmessage.Subject = $emailsubject
$mailmessage.Body = $emailbody

$attachment = New-Object System.Net.Mail.Attachment($emailattachment, 'text/plain')
  $mailmessage.Attachments.Add($attachment)


#$mailmessage.IsBodyHTML = $true
$SMTPClient = New-Object Net.Mail.SmtpClient $SmtpServer 
$SMTPClient.Send($mailmessage)
}
send_email
"Email Sent"


$smtphost = "mailgateway.Contoso.corp" 
$from = "ADHealthCheck@creditone.com" 
$email1 = "john.battista@creditone.com"
$timeout = "60"

$subject = "Active Directory Health Monitor" 
$body = Get-Content $filename, $filename2, $filename3, $ADHealthReport
$smtp= New-Object System.Net.Mail.SmtpClient $smtphost 
$msg = New-Object System.Net.Mail.MailMessage 
$msg.To.Add($email1)
$msg.from = $from
$msg.subject = $subject
$msg.body = $body 
$msg.isBodyhtml = $true 
$smtp.send($msg) 