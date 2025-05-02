#Script Variables
$sESXiHost = 'phxvcenter01.phx.fnbm.corp'
$sCommunity = 'pilot'
$sTarget = '192.168.99.1'
$sPort = '161'

#Connect to ESXi host
Connect-VIServer -Server $sESXiHost

#Clear SNMP Settings
Get-VMHostSnmp | Set-VMHostSnmp -ReadonlyCommunity @()

#Add SNMP Settings
Get-VMHostSnmp | Set-VMHostSnmp -Enabled:$true -AddTarget -TargetCommunity $sCommunity -TargetHost $sTarget -TargetPort $sPort -ReadOnlyCommunity $sCommunity

#Get SNMP Settings
$Cmd= Get-EsxCli -V2 -VMHost $sESXiHost
$Cmd.System.Snmp.Get()