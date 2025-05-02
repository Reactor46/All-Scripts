Write-Warning "The following commands must be run on a remote computer."

$ComputerName = Read-Host "FQDN of the RDS Server (COMP-REMOTE01)"

## Installing Remote Desktop Services
Import-Module RemoteDesktop

New-RDSessionDeployment -ConnectionBroker $ComputerName `
    -WebAccessServer $ComputerName `
    -SessionHost $ComputerName

Write-Verbose "Successfully created new RDS deployment on : $ComputerName"

Write-Output "Please continue with the RDSConfiguration.ps1 script on $ComputerName."