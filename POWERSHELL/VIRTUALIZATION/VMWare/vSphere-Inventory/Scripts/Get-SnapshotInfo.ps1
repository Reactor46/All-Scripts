function Get-SnapshotInfo
{
	Param(
		[Parameter(Mandatory = $True)]
		[VMware.VimAutomation.Types.VirtualMachine] $vm,
		
		[Parameter(Mandatory = $True)]
		[string] $vcName,
		
		[Parameter(Mandatory = $True)]
		[string] $dcName
	)
	
	$report = @()

	$vm | Get-Snapshot | % { 
		$SizeGB = [Math]::Round($_.SizeGB, 3)
		
		$reportLine = New-Object psobject
		$reportLine | Add-Member -Type noteproperty -Name vCenter -Value $vcName
		$reportLine | Add-Member -Type noteproperty -Name Datacenter -Value $dcName
		$reportLine | Add-Member -Type noteproperty -Name VM -Value $_.VM
		$reportLine | Add-Member -type noteproperty -name Name -Value $_.Name
		$reportLine | Add-Member -type noteproperty -name Description -Value $_.Description
		$reportLine | Add-Member -type noteproperty -name Created -Value $_.Created
		$reportLine | Add-Member -type noteproperty -name SizeGB -Value $SizeGB
		
		$report += $reportLine
	}
	
	return $report
}