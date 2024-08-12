(Get-MailboxDatabase) | foreach {write-Host $_; (Get-Mailbox -database $_ -resultsize "Unlimited" |ft).Count}

Method that lists count next to db name rather than under it.
(Get-MailboxDatabase) | ForEach-Object {Write-Host $_.Name (Get-Mailbox -Database $_.Name -resultsize "Unlimited").Count}


eseutil /d /p "H:\Students_16\exchange\mailbox\students_16\students_16.edb" /t"H:\Students_16\exchange\mailbox\students_16\students_16_df.edb"




