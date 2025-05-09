# Function for receiving information about ports of switching equipment,
# which are connected to the uplink ports of the ESXi host.
# Parameters: ESXi host, vCenter name, Datacenter name.
#
# Lyubimov Roman, 2015-2017

function Get-HostUplinkInfo
{
	Param(
		[Parameter(Mandatory = $True)]
		[VMware.VimAutomation.Types.VMHost] $hst,
		
		[Parameter(Mandatory = $True)]
		[string] $vcName,
		
		[Parameter(Mandatory = $True)]
		[string] $dcName
	)
	
	$report = @()
		
	$vmhost = Get-View $hst
	$networkSystem = Get-View $vmhost.ConfigManager.NetworkSystem

	foreach($physNic in $networkSystem.NetworkInfo.Pnic) {
		$physNicInfo = $networkSystem.QueryNetworkHint($physNic.Device)

		foreach($hint in $physNicInfo)
		{
			$reportLine = New-Object psobject
			$reportLine | Add-Member -Type noteproperty -Name vCenter -Value $vcName
			$reportLine | Add-Member -Type noteproperty -Name Datacenter -Value $dcName
			$reportLine | Add-Member -Type noteproperty -Name Parent -Value $hst.Parent.Name
			$reportLine | Add-Member -Type noteproperty -Name Host -Value $hst.Name
			$reportLine | Add-Member -Type noteproperty -Name Uplink -Value $physNic.Device
			$reportLine | Add-Member -Type noteproperty -Name Device -Value $hint.ConnectedSwitchPort.DevId
			$reportLine | Add-Member -Type noteproperty -Name Address -Value $hint.ConnectedSwitchPort.Address
			$reportLine | Add-Member -Type noteproperty -Name Port -Value $hint.ConnectedSwitchPort.PortId
		}
			
		$report += $reportLine
	}
			
	return $report
}