Get-MailboxDatabaseCopyStatus -Server ANEXMBX01 | Suspend-MailboxDatabaseCopy -ActivationOnly -Confirm:$False -SuspendComment "Hardware Installtion - Memory"