Set-Location C:\Scripts\Repository\jbattista\Web\reports\AdminScripts\ChatServers

Get-Content -Path .\ChatServers.txt |
    ForEach-Object {Get-WmiObject -Namespace ROOT\StandardCIMV2 -Class MSFT_NetTCPConnection -Computer $_ |
        Select-Object RemoteAddress, RemotePort, OwningProcess, PSComputerName |
            Export-CSV ".\ChatServersConnections.csv" -Append -NoTypeInformation}