
C:\LazyWinAdmin\Networking\Get-NetConfig.ps1 (GC C:\LazyWinAdmin\VMWare\Nic-Change.txt) |
Select ComputerName, IPAddress, SubnetMask, Gateway, DNSServers, MACAddress | Export-CSV C:\LazyWinAdmin\Networking\VM-Nics.csv -NoTypeInformation -Append 
