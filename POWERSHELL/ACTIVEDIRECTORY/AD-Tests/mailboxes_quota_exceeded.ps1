#This script is designed to determine which users are over their quota determined limit and enable an archive mailbox on the archive datastore
#Created by Max Maskevich
#For questions, feedback, etc please email me at max.maskevich@gmail.com

#Retrieves mailboxes where the size limit has generated a warning sent to the user.
Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics | where {$_.StorageLimitStatus -ne "BelowLimit"} | Select Displayname | Export-CSV "C:\scripts\exceededquotas.csv" -NoTypeInformation

#Takes the list of warned users and enables their archive, and places them in the retention policy used.
Import-Csv "c:\scripts\exceededquotas.csv" | ForEach {enable-Mailbox -Identity $_.displayname -Archive -ArchiveDatabase "Archive Datastore Name"} | ForEach {set-Mailbox -Identity $_.displayname -retentionpolicy "Archival Policy Name"} | out-file "c:\scripts\archiveresults.txt"

#This section generates an email if the archive results file has enough content to not be blank.  It only sends an email if a change is made.
#This email includes the users who had their archiving/rentention policy settings changed.
if( (get-item C:\scripts\archiveresults.txt).length -gt 2)
{
Send-MailMessage -To "Exchange Admin <exchangeadmin@domainname.com>" -From "Exchange Server <exchange@domainname.com>" -Subject "Archive Script Log" -SmtpServer "127.0.0.1" -Body "The exchange archive script has run - Here are the results." -Attachments "C:\scripts\archiveresults.txt"
}