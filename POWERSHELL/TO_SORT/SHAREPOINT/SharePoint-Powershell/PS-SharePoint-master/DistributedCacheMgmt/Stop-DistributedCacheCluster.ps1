[CmdletBinding()]
param ()

begin {
}

process {
    Write-Output "$(Get-Date -Format T) : Stopping Distributed Cache Service Instance on all servers..."
    Get-SPServiceInstance | Where-Object { $_.TypeName -eq "Distributed Cache" } | Stop-SPServiceInstance -Confirm:$false | Out-Null
    While (Get-SPServiceInstance | Where-Object { $_.TypeName -eq "Distributed Cache" -and $_.Status -ne "Disabled" }) {
        Write-Output "$(Get-Date -Format T) : Waiting for all Distributed Cache Service Instances to stop..."
        Start-Sleep -Seconds 15
    }
    Write-Output "$(Get-Date -Format T) : All Distributed Cache Service Instances stopped!"
}