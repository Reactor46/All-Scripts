# Script to restart IIS on all farm servers.

# Main code

$servers = (Get-SPFarm).Servers
foreach ($server in $servers) {
 if ($server.Role -eq [Microsoft.SharePoint.Administration.SPServerRole]::Invalid) {
  continue
 }

 Write-Host "Restarting service on" $server.Name
 Restart-Service -InputObject (Get-Service -ComputerName $server.Name -Name "w3svc")
}

Write-Host "Done!" -ForegroundColor Green

# End main code