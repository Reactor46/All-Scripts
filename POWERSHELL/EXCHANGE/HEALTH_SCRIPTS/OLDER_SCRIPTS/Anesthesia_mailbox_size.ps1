$Users = Get-Content E:\Scripts\Anesthesia.txt

Foreach($user in $Users){

Get-MailboxStatistics $user | Select TotalItemSize | Measure #Export-CSV -Path E:\Scripts\Anesthesia_Mailboxes.csv -Delimiter "," -NoTypeInformation -Verbose
}