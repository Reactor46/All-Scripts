New-MailboxExportRequest -Mailbox "user@example.com" -FilePath "\\exchange\export\User.pst" -Name "CustomExportName" -Description "Description of the export"

New-MailboxExportRequest -Mailbox "user@example.com" -IsArchive -FilePath "\\exchange\export\UserArchive.pst" -Name "User Archive"

Get-MailboxExportRequest -Mailbox "User Display Name or Mailbox"

Get-MailboxExportRequest -Identity "User Display Name or Mailbox"

Get-MailboxExportRequest -Status

Get-MailboxExportRequest -Status InProgress

Remove-MailboxExportRequest -Identity "User Display Name or Mailbox"

Get-MailboxDatabase -Status | Format-Table Name, DatabaseSize, AvailableNewMailboxSpace -Auto