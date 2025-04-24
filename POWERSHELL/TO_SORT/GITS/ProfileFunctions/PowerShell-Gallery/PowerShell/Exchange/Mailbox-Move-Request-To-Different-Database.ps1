Get-MailboxDatabase

Get-MoveRequest -Identity "user@domain.example" |  Get-MoveRequestStatistics

New-MoveRequest -Identity "user@domain.example" -TargetDatabase "Mailbox Database 2134000348"

New-MoveRequest -Identity "user@domain.example" -TargetDatabase "Mailbox Database 2134000348" -ArchiveTargetDatabase "ArchiveDatabase"

Remove-MoveRequest -Identity "user@domain.example"

Get-MigrationBatch -Identity "Mailbox 1 and Mailbox 2" | Select-Object Identity, Status, TotalCount, FinalizedCount, FailedCount

New-MoveRequest -Identity "user@example.com" -TargetDatabase "Mailbox Database 1986552740"

New-MoveRequest -Identity "user@example.com" -TargetDatabase "Mailbox Database 2134000348" -ArchiveTargetDatabase "ArchiveDatabase"

Get-MoveRequest -Identity "user@example.com" | Get-MoveRequestStatistics

Get-MoveRequest | Get-MoveRequestStatistics

Remove-MoveRequest -Identity "user@example.com"

Get-Mailbox -Database "Mailbox Database 1986552740" | Sort-Object TotalItemSize

Get-Mailbox -Database "ArchiveDatabase2019" | Sort-Object DisplayName

Get-MailboxDatabase -Status | Format-Table Name, DatabaseSize, AvailableNewMailboxSpace -Auto