Get-MailboxFolderPermission lsmith | select-object foldername,accessrights,user

Get-MailboxFolderStatistics -Identity 'lsmith' -FolderScope RecoverableItems | Format-Table Name,FolderPath,ItemsInFolder,FolderAndSubfolderSize