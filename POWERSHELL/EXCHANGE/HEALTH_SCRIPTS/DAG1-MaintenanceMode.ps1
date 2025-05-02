## USONVSVRDAG01 Maintenance Mode
#Get-MailboxDatabase | where {$_.Server -eq "USONVSVRDAG01"}
#Move-ActiveMailboxDatabase "DAG01 MBX-1" -ActivateOnServer USONVSVRDAG02
#Move-ActiveMailboxDatabase "DAG02 MBX-1" -ActivateOnServer USONVSVRDAG02
#Get-MailboxServer USONVSVRDAG01 | fl Name,DatabaseCopyAutoActivationPolicy

Get-MailboxDatabase | where {$_.Server -eq "USONVSVRDAG02"} | Move-ActiveMailboxDatabase -ActivateOnServer USONVSVRDAG01 -Confirm:$false
Set-MailboxServer USONVSVRDAG02 -DatabaseCopyAutoActivationPolicy Blocked

$exscripts\StartDagServerMaintenance.ps1 -serverName USONVSVRDAG02

#$exscripts\StopDagServerMaintenance.ps1 -serverName USONVSVRDAG01


# Run after Each Maintenance Start and Stop

.\RedistributeActiveDatabases.ps1 -DagName ExchangeDAG1 -ShowDatabaseDistributionByServer | ft

.\RedistributeActiveDatabases.ps1 -DagName ExchangeDAG1 -BalanceDbsByActivationPreference
Get-MailboxServer -Status | fl Name,DatabaseCopyAutoActivationPolicy

Get-MailboxDatabase | ft name, server, activationpreference -AutoSize

Get-DatabaseAvailabilityGroup -Status | FL Name, ServersInMaintenance