Set-MailboxServer –Identity <ServerName> –DatabaseCopyAutoActivationPolicy Unrestricted
Resume-ClusterNode –Name <ServerName>
Set-ServerComponentState –Identity <ServerName> –Component HubTransport –State Active –Requester Maintenance 
Set-MailboxServer –Identity <ServerName> –DatabaseCopyActivationDisabledAndMoveNow $False