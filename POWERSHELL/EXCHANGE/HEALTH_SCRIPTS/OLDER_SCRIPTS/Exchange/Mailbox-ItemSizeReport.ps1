$mailboxes = @(Get-Mailbox -ResultSize Unlimited)
$report = @()
 
foreach ($mailbox in $mailboxes)
{
    $mboxstats = Get-MailboxFolderStatistics $mailbox
 
    $mbObj = New-Object PSObject
    $mbObj | Add-Member -MemberType NoteProperty -Name "Display Name" -Value $mailbox.DisplayName
    $mbObj | Add-Member -MemberType NoteProperty -Name "Inbox Size (Mb)" -Value $mboxstats.FolderandSubFolderSize.ToMB()
    $mbObj | Add-Member -MemberType NoteProperty -Name "Inbox Items" -Value $mboxstats.ItemsinFolderandSubfolders
    $mbObj | Add-Member -MemberType NoteProperty -Name "Oldest Item" -Value $mboxstats.OldestItemReceivedDate
    $report += $mbObj
}
 
$report