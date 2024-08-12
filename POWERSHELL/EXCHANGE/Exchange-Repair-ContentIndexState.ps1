Get-MailboxDatabaseCopyStatus * | ? {$_.Status -ne “Mounted” –and $_.ContentIndexState -ne “Healthy”}
Get-MailboxDatabaseCopyStatus * | ? {$_.Status -ne “Mounted” -and $_.ContentIndexState -ne “Healthy”} | Update-MailboxDatabaseCopy -DeleteExistingFiles
