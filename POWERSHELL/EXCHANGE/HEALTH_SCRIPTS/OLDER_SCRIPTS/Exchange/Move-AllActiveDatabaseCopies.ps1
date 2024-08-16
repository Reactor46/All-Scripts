Move all DBs on a node to another node
Get-MailboxDatabase | Where-Object {$_.Server -eq "ANEXMBX03"} | Move-ActiveMailboxDatabase -ActivateOnServer ANEXMBX02 -MountDialOverride:None


Move all dbs matching name pattern to another node
Get-MailboxDatabase | Where-Object {$_.Server -eq "ANEXMBX01" -And $_.Name -Like "Students*"} | Move-ActiveMailboxDatabase -ActivateOnServer ANEXMBX03 -MountDialOverride:None