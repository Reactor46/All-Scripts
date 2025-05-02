Set-MailboxServer <ServerName> -DatabaseCopyActivationDisabledAndMoveNow $True
Set-ServerComponentState –Identity <ServerName> –Component HubTransport –State Draining –Requester Maintenance
Suspend-ClusterNode –Name <ServerName>
Set-MailboxServer –Identity <ServerName> –DatabaseCopyAutoActivationPolicy Blocked