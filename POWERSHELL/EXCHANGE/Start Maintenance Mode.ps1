Set-MailboxServer <ServerName> -DatabaseCopyActivationDisabledAndMoveNow $True
****Put Exchange 2016 DAG Member into maintenance mode****

Set-ServerComponentState USONVSVREX01 -Component HubTransport -State Draining -Requester Maintenance

Restart-Service MSExchangeTransport

Restart-Service MSExchangeFrontEndTransport

Redirect-Message -Server USONVSVRDAG01 -Target USONVSVRDAG02 ***Any Exchange server not in Maintenance Mode***

Suspend-ClusterNode USONVSVREX01

Set-MailboxServer USONVSVRDAG01 -DatabaseCopyActivationDisabledAndMoveNow $True

Get-MailboxServer USONVSVRDAG02 | Select DatabaseCopyAutoActivationPolicy

Set-MailboxServer USONVSVRDAG02 -DatabaseCopyAutoActivationPolicy Blocked

Set-ServerComponentState USONVSVREX01 -Component ServerWideOffline -State Inactive -Requester Maintenance




Set-ServerComponentState –Identity <ServerName> –Component HubTransport –State Draining –Requester Maintenance
Suspend-ClusterNode –Name <ServerName>
Set-MailboxServer –Identity <ServerName> –DatabaseCopyAutoActivationPolicy Blocked