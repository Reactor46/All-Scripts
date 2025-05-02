[CmdletBinding()]
param ()

begin {
}

process {
    Write-Output "$(Get-Date -Format T) : Starting Distributed Cache Service Instance on all servers..."
    Get-SPServiceInstance | Where-Object { $_.TypeName -eq "Distributed Cache" } | Start-SPServiceInstance | Out-Null
    While (Get-SPServiceInstance | Where-Object { $_.TypeName -eq "Distributed Cache" -and $_.Status -ne "Online" }) {
        Write-Output "$(Get-Date -Format T) : Waiting for all Distributed Cache Service Instances to start..."
        Start-Sleep -Seconds 15
    }
    Write-Output "$(Get-Date -Format T) : All Distributed Cache Service Instances started!"
}