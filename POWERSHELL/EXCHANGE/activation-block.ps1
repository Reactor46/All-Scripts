Suspend-MailboxDatabaseCopy –identity DB1\EXCH3 –ActivationOnly

get-mailboxdatabasecopystatus *\<server name> | suspend-mailboxdatabasecopy -activationOnly:$TRUE

Reverse
Resume-MailboxDatabaseCopy –identity DB1\EXCH3

get-mailboxdatabasecopystatus *\<server name> | resume-mailboxdatabasecopy