############################################################################################################## DAILY CHECKS V2 ################################################################################################################

#HTML
 
$style = "<style>"
$style = $style + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$style = $style + "TH{border-width: 1px;padding: 10px;border-style: solid;border-color: black;}"
$style = $style + "TD{border-width: 1px;padding: 10px;border-style: solid;border-color: black;}"
$style = $style + "</style>"

#Set Date
$date = (get-date).ToString('dd/MM/yy')
$backupdatelimit = (Get-Date).AddDays(-5)

############################################################################################################## SCRIPTS ########################################################################################################################

############################################################################################################## QUEUES #########################################################################################################################
$Queues = Get-Queue | where {$_.Messagecount -gt "10"} 

if ($Queues -eq $null)
{
$a += "<p>No Queues with more than 10 emails queued.</p>"
}
else
{
    $a += $Queues |  Select-Object queueidentity, status,messagecount | ConvertTo-HTML -head $style -body "<H2>Mail Queues</H2>" | Write-Output 
}

############################################################################################################## SMTP ###########################################################################################################################
$SMTP = Test-SmtpConnectivity | where {$_.statuscode -ne "Success"} 

if ($SMTP -eq $null)
{
$a += "<p>No SMTP Problems.</p>"
}
else
{
    $a += $SMTP | select receiveconnector,statuscode | ConvertTo-HTML -head $style -body "<H2>SMTP Connectivity</H2>"
}
############################################################################################################## OWA ############################################################################################################################
$OWA = Test-OwaConnectivity | where {$_.EventType -ne "Success"}

if ($OWA -eq $null)
{
$a += "<p>No OWA Problems.</p>"
}
else
{
$a += $OWA | select url, scenario, result | ConvertTo-HTML -head $style -body "<H2>OWA Connectivity</H2>"
}

############################################################################################################## DATABASES ######################################################################################################################
$Databases = Get-Mailboxserver | Get-MailboxDatabaseCopyStatus |  where {$_.Status -ne "Mounted" -and $_.Status -ne "Healthy" -or $_.copyqueuelength -gt 10 -or $_.replayqueuelength -gt 10}
if ($Databases -eq $null)
{
$a += "<p>No Database Problems.</p>"
}
else
{
    $a += $Databases | select mailboxserver, name, status, copyqueuelength,replayqueuelength, contentindexstate | sort mailboxserver | ConvertTo-HTML -head $style -body "<H2>Mailbox Databases</H2>"
}

############################################################################################################## Backups ########################################################################################################################
$Backups = Get-MailboxDatabase -Status | where {$_.name -like "NAME*" -and $_.LastFullBackup -lt $backupdatelimit}
if ($backups -eq $null)
{
$a += "<p>The last full Exchange backup was $_.LastFullBackup.</p>"
}
else
{
    $a += $Backups | select name,lastfullbackup,lastincrementalbackup | sort name | ConvertTo-HTML -head $style -body "<H2>Database Backup Status</h2>"
}

############################################################################################################## Mailbox Size ##################################################################################################################
$MailboxSize = Get-Mailbox | Get-MailboxStatistics| where {$_.totalitemsize -gt "15GB"} 
if ($MailboxSize -eq $null)
{
$a +="<p>There are no mailboxes over 15GB.</p>" 
}
else
{
     $a += $MailboxSize | select displayname,totalitemsize, totaldeleteditemsize | sort totalitemsize | ConvertTo-HTML -head $style -body "<H2>Mailboxes over 15GB</h2>"
}

####################################################################################################### Create attachments ##################################################################################################################

Get-Queue | Select-Object queueidentity, status,messagecount | export-csv “c:\scripts\DailyChecks\attachments\Queues.csv” -NoTypeInformation

Test-SmtpConnectivity | select receiveconnector,statuscode | export-csv “c:\scripts\DailyChecks\attachments\SMTP.csv” -NoTypeInformation

Test-OwaConnectivity | select url, scenario, result | export-csv “c:\scripts\DailyChecks\attachments\OWA.csv” -NoTypeInformation

Test-ActiveSyncConnectivity | select ClientAccessServerShortName,scenario,result,error| export-csv “c:\scripts\DailyChecks\attachments\ActiveSync.csv” -NoTypeInformation

Get-Mailbox | Get-MailboxStatistics| where {$_.totalitemsize -gt "10GB"} | select displayname,totalitemsize, totaldeleteditemsize | sort totalitemsize | export-csv “c:\scripts\DailyChecks\attachments\Mailbox Size.csv” -NoTypeInformation

Get-MailboxDatabase -Status | where {$_.name -like "NAME*"} | select name,lastfullbackup,lastincrementalbackup | sort name | export-csv “c:\scripts\DailyChecks\attachments\Databases.csv” -NoTypeInformation

Get-Mailboxserver | where {$_.name -like "NAME*"} | Get-MailboxDatabaseCopyStatus | select mailboxserver, name, status, copyqueuelength,replayqueuelength, contentindexstate | sort mailboxserver | export-csv “c:\scripts\DailyChecks\attachments\Databasecopystatus.csv” -NoTypeInformation

$QueuesAttachment = “c:\scripts\DailyChecks\attachments\Queues.csv”

$smtpAttachment = “c:\scripts\DailyChecks\attachments\SMTP.csv”

$owaAttachment = “c:\scripts\DailyChecks\attachments\OWA.csv”

$ActiveSyncAttachment = “c:\scripts\DailyChecks\attachments\ActiveSync.csv”

$MailboxSizeAttachment = “c:\scripts\DailyChecks\attachments\Mailbox Size.csv”

$DatabasesAttachment = “c:\scripts\DailyChecks\attachments\Databases.csv”

$DatabaseCopyStatusAttachment = “c:\scripts\DailyChecks\attachments\Databasecopystatus.csv”

####################################################################################################### Send Emails ##########################################################################################################################

Send-MailMessage -SmtpServer SMTPSERVER -To EMAILADDRESS -From DailyChecks@DOMAIN.COM -Subject "Exchange Environment Daily Checks: SERVER NAME - $date" -Body $a -attachment $QueuesAttachment,$smtpAttachment,$owaAttachment,$ActiveSyncAttachment,$MailboxSizeAttachment,$DatabasesAttachment,$DatabaseCopyStatusAttachment -BodyAsHtml  
