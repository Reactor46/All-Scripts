# Функция для получения информации о стораджах.
# Параметры: Datacenter, имя vCenter.
#
# Любимов Роман, 2015-2017

function Get-StorageInfo
{
	Param(
		[Parameter(Mandatory = $True)]
		[VMware.VimAutomation.Types.Datacenter] $dc,
		
		[Parameter(Mandatory = $True)]
		[string] $vcName
	)
	
	$report = @()

	Get-View -ViewType Datastore -SearchRoot $dc.Id | % { 
		$capacityGB    = [math]::Round($_.Summary.Capacity / 1GB, 0)
		$provisionedGB = [math]::Round(($_.Summary.Capacity - $_.Summary.FreeSpace + $_.Summary.Uncommitted) / 1GB, 0)
		$uncommittedGB = [math]::Round($_.Summary.Uncommitted / 1GB, 0)
		$freeGB        = [math]::Round($_.Summary.FreeSpace / 1GB, 0)
		
		$parent = ""
		if ($_.Parent.Type -eq "Folder") {
			$parent = Get-Folder -Id $_.Parent
		}
		if ($_.Parent.Type -eq "StoragePod") {
			$parent = Get-DatastoreCluster -Id $_.Parent
		}
		
		$reportLine = New-Object psobject
		$reportLine | Add-Member -Type noteproperty -Name vCenter -Value $vcName
		$reportLine | Add-Member -Type noteproperty -Name Datacenter -Value $dc.Name
		$reportLine | Add-Member -Type noteproperty -Name Parent -Value $parent.Name
		$reportLine | Add-Member -type noteproperty -name Storage -Value $_.Name
		$reportLine | Add-Member -type noteproperty -name Type -Value $_.Summary.Type
		$reportLine | Add-Member -type noteproperty -name MultipleHostAccess -Value $_.Summary.MultipleHostAccess
		$reportLine | Add-Member -type noteproperty -name SIOCEnabled -Value $_.IormConfiguration.Enabled
		$reportLine | Add-Member -type noteproperty -name CapacityGB -Value $capacityGB
		$reportLine | Add-Member -type noteproperty -name ProvisionedGB -Value $provisionedGB
		$reportLine | Add-Member -type noteproperty -name UncommittedGB -Value $uncommittedGB
		$reportLine | Add-Member -type noteproperty -name FreeSpaceGB -Value $freeGB
		$reportLine | Add-Member -type noteproperty -name OverprovisionGB -Value ($provisionedGB - $capacityGB)

		$report += $reportLine
	}
	
	return $report | sort Parent, Storage
}