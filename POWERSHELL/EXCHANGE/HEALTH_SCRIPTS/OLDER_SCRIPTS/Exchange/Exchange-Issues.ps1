Search-Mailbox -Identity "Larry Smith" -SearchDumpsterOnly -TargetMailbox "Discovery Search Mailbox" -TargetFolder "LarrySmith-RecoverableItems" -DeleteContent
Search-Mailbox -Identity "Alan Eufemio" -SearchDumpsterOnly -TargetMailbox "Discovery Search Mailbox" -TargetFolder "AlanEufemio-RecoverableItems" -DeleteContent
Get-MailboxFolderStatistics -Identity "Alan Eufemio" -FolderScope RecoverableItems | Format-Table Name,FolderAndSubfolderSize,ItemsInFolderAndSubfolders -Auto
Get-MailboxFolderStatistics -Identity "Larry Smith" -FolderScope RecoverableItems | Format-Table Name,FolderAndSubfolderSize,ItemsInFolderAndSubfolders -Auto
Get-MailboxFolderStatistics -Identity "Larry Smith" | Select Name,FolderSize,ItemsinFolder
Get-MailboxFolderStatistics -Identity "Alan Eufemio" | Select Name,FolderSize,ItemsinFolder