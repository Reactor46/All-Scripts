# api: multitool
# version: 0.5
# title: Exchange health checks
# description: probe various server performance/database stats
# type: inline
# depends: funcs_base
# category: exchange
# icon: firewall
# param: exchangeserver
# hidden: 0
# key: m10|exchangetests?
# status: beta
# config: -
#
# Run basic Exchange server health checks
#  ❏ Test-ServiceHealth
#  ❏ Test-Mailflow
#  ❏ Get-MailboxDatabase
#  ❏ Get-MailboxDatabase


Param($server = (Read-Host "exchangeserver"));


#-- conn
Import-ExchangeSession


#-- tests
Write-Host -f Green "❏ Test-ServiceHealth"
Test-ServiceHealth | FL | Out-String | Write-Host

Write-Host -f Green "❏ Test-EcpConnectivity"
Test-EcpConnectivity -ClientAccessServer $server | Out-String | Write-Host

Write-Host -f Green "❏ Test-Mailflow"
Test-Mailflow -Targetmailboxserver $server | FL -Prop * | Out-String | Write-Host

Write-Host -f Green "❏ Get-MailboxDatabase -Status"
Get-MailboxDatabase -Status -Server $server | FT name,server,mounted,replicationtype,recovery -Auto -Wrap | Out-String | Write-Host

Write-Host -f Green "❏ Get-Queue"
Get-Queue -Server $server | FL -Prop PSComputerName,Identity,IsValid,Status,MessageCount,RetryCount,LastError,RiskLevel,IncomingRate,OutgoingRate,PriorityDescriptions,DeferredMessageCount,LockedMessageCount | Out-String | Write-Host








