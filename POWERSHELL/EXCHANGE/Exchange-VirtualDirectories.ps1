Start-Transcript D:\DNS\EnvironmentBackup.txt

Get-OabVirtualDirectory | fl server, Name, ExternalURL, InternalURL, *auth*
Get-WebServicesVirtualDirectory | fl server, Name,ExternalURL, InternalURL, *auth*
Get-EcpVirtualDirectory | fl server, Name, ExternalURL, InternalURL, *auth*
Get-OutlookAnywhere | fl server, Name, *hostname*, *auth*
Get-OwaVirtualDirectory | fl server, Name, ExternalURL, InternalURL, *auth*
Get-MapiVirtualDirectory | fl server, Name,ExternalURL,InternalURL, *auth*
Get-OutlookProvider
Get-ClientAccessServer | fl Name,OutlookAnywhereEnabled, AutodiscoverServiceInternalUri
Get-ExchangeCertificate | fl Services, CertificateDomains, Issuer, *not*
Get-OutlookAnywhere | fl Name, *hostname*, *auth*
Get-ClientAccessArray | fl
Get-ActiveSyncVirtualDirectory | Format-List
Get-AutodiscoverVirtualDirectory | Format-List
Get-PowerShellVirtualDirectory | Format-List
Get-SendConnector | Where-Object {$_.Enabled -eq $true} | Format-List
Get-SendConnector | Where-Object {$_.Enabled -eq $true} | Get-ADPermission | Where-Object { $_.extendedrights -like '*routing*' } | fl identity, user, *rights
Resolve-DnsName -Type A -Name mail.optummso.com
Resolve-DnsName -Type A -Name autodiscover.optummso.com
Resolve-DnsName -Type A -Name mail.optummso.com -Server 10.20.11.11
Resolve-DnsName -Type A -Name autodiscover.domain.com -Server 10.20.11.11
Resolve-DnsName -Type MX -Name optummso.com -Server 10.20.11.11
Resolve-DnsName -Type TXT -Name optummso.com -Server 10.20.11.11
#Resolve-DnsName -Type A -Name i-should-not-exist.domain.com -Server 10.20.11.11
Stop-Transcript