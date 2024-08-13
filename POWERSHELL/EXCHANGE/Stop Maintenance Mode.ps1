****Take Exchage 2016 DAG Member out of maintenance mode****

Set-ServerComponentState USONVSVREX01 -Component ServerWideOffline -State Active -Requester Maintenance

Resume-ClusterNode USONVSVREX01

Set-MailboxServer USONVSVREX01 -DatabaseCopyActivationDisabledAndMoveNow $False

Set-MailboxServer LASEXCH01 -DatabaseCopyAutoActivationPolicy Unrestricted

Set-ServerComponentState LASEXCH01 -Component HubTransport -State Active -Requester Maintenance

Restart-Service MSExchangeTransport

Restart-Service MSExchangeFrontEndTransport

****Rebalance BDs****
C:\Program Files\Microsoft\Exchange Server\V15\Scripts> .\RedistributeActiveDatabases.ps1 -BalanceDbsByActivationPreference -Confirm:$false





<#
Set-MailboxServer –Identity <ServerName> –DatabaseCopyAutoActivationPolicy Unrestricted
Resume-ClusterNode –Name <ServerName>
Set-ServerComponentState –Identity <ServerName> –Component HubTransport –State Active –Requester Maintenance 
Set-MailboxServer –Identity <ServerName> –DatabaseCopyActivationDisabledAndMoveNow $False
#>