## The "Exchange Trusted Subsystem" accound read permission on the directory that stores the pst file!
New-MailboxExportRequest -mailbox <username> -FilePath <filename>.pst

Check status of specific request
Get-MailboxExportRequestStatistics <RequestGUID>


Check status of all
Get-MailboxExportRequest | Get-MailboxExportRequestStatistics