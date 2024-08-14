$esxlist = Get-Content ".\phxvcenter01-hosts.txt"
foreach($item in $esxlist){
Connect-VIServer $item -User root -Password !p6+_6Cu5u
$item | Set-VMHostSnmp -Enabled:$true
Set-VMHostSnmp -HostSnmp $_ -ReadOnlyCommunity "pilot" -TargetHost "192.168.99.1"
Disconnect-VIServer -Confirm:$false }