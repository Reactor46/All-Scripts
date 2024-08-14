Connect-VIServer -Server 192.168.1.201 -User "administrator@vsphere.local" -Password "J@bb3rJ4w" -ErrorAction Continue

Get-VM | Get-CDDrive | Where {$_.ISOPath -ne $null} | Set-CDDrive -NoMedia -Confirm:$false
Get-VM | FT Name, @{Label="ISO file"; Expression = { ($_ | Get-CDDrive).ISOPath }} 