function Get-HostInfo
{
	Param(
		[Parameter(Mandatory = $True)]
		[VMware.VimAutomation.Types.VMHost] $hst,
		
		[Parameter(Mandatory = $True)]
		[string] $vcName,
		
		[Parameter(Mandatory = $True)]
		[string] $dcName
	)

	$vmCount = 0
	$virtualCores = 0
	$virtualRam = 0
		
	$hst | Get-VM | % {
		$vmCount += 1
		$virtualCores += $_.NumCpu
		$virtualRam += $_.MemoryGB
	}
			
	$report = New-Object psobject
		
	$report | Add-Member -type noteproperty -name vCenter -Value $vcName
	$report | Add-Member -type noteproperty -name Datacenter -Value $dcName
	$report | Add-Member -type noteproperty -name Parent -Value $hst.Parent.Name
	$report | Add-Member -type noteproperty -name Host -Value $hst.Name
	$report | Add-Member -type noteproperty -name ESXiVersion -Value $hst.Version
	$report | Add-Member -type noteproperty -name ESXiBuild -Value $hst.Build
	$report | Add-Member -type noteproperty -name Manufacturer -Value $hst.Manufacturer
	$report | Add-Member -type noteproperty -name Model -Value $hst.Model
	$report | Add-Member -type noteproperty -name CPUType -Value $hst.ProcessorType
	$report | Add-Member -type noteproperty -name CPUCount -Value $hst.ExtensionData.Hardware.CpuInfo.NumCpuPackages
	$report | Add-Member -type noteproperty -name PhysCores -Value $hst.ExtensionData.Hardware.CpuInfo.NumCpuCores
	$report | Add-Member -type noteproperty -name LogicCores -Value $hst.ExtensionData.Hardware.CpuInfo.NumCpuThreads
	$report | Add-Member -type noteproperty -name PhysRAMGB -Value ([Math]::Round($hst.MemoryTotalGB))	
	$report | Add-Member -type noteproperty -name VMCount -Value $vmCount
	$report | Add-Member -type noteproperty -name VirtCores -Value $virtualCores        
	$report | Add-Member -type noteproperty -name VirtRAMGB -Value ([Math]::Round($virtualRAM))
	
	$cpuP2VRatio = "Error"
	if ($report.PhysCores -gt 0) {
		$cpuP2VRatio = "1::" + [Math]::Round($report.VirtCores / $report.PhysCores)
	}
	$report | Add-Member -type noteproperty -name CPUP2VRatio -Value $cpuP2VRatio
	$report | Add-Member -type noteproperty -name RAMOverprovisionGB -Value ($report.VirtRAMGB - $report.PhysRAMGB)
				
	return $report
}