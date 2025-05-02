$vCenter = "192.168.1.200"
$Hosts = "192.168.1.1","192.168.1.2"."192.168.1.7"
$Destination = "C:\LazyWinAdmin\VMWare\VMware Reports\"

Connect-VIServer -Server 192.168.1.200  -User "administrator@vsphere.local" -Password "J@bb3rJ4w" -ErrorAction SilentlyContinue -WarningAction 0
#Get-Log -Bundle -DestinationPath $Destination

Foreach($host in $Hosts){
Get-VMHost $host | Get-Log -Bundle -DestinationPath $Destination
}
