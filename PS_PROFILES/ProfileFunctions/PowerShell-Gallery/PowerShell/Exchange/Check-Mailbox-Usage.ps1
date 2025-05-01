#
# Get mailbox statistics for all mailboxes and select relevant properties
#
$mailboxStats = Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics | Select DisplayName, TotalItemSize
#
# Convert mailbox size to gigabytes and sort the mailboxes by size
#
$mailboxStats | Select DisplayName, @{Name="TotalSizeGB";Expression={[math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","") / 1GB),2)}} | Sort TotalSizeGB -Descending | Export-Csv -Path "C:\MailboxUsage.csv" -NoTypeInformation