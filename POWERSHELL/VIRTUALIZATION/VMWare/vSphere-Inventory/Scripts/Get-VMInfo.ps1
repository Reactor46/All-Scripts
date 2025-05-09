# Function to get information about virtual machines.
# Parameters: VM, vCenter name, Datacenter name.
#
# Lyubimov Roman, 2015-2017

function Get-VMInfo
{
	Param(
		[Parameter(Mandatory = $True)]
		[VMware.VimAutomation.Types.VirtualMachine] $vm,
		
		[Parameter(Mandatory = $True)]
		[string] $vcName,
		
		[Parameter(Mandatory = $True)]
		[string] $dcName
	)
	
	# Get the path to the VM in VMs and Templates.
	$path = $vm.Folder.Name
	$currentFolder = $vm.Folder
	while ($currentFolder.Name -ne "vm") {
		$currentFolder = $currentFolder.Parent
		$path = $currentFolder.Name + "/" + $path
	}
	
	# Select IPv4 addresses,combine into string.
	$ip = ""
	$vm.Guest.IPAddress | % {
		if ($_ -Match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}") {
			$ip += $_ + " "
		}
	}
		
	# Remove line breaks
	$notes = $vm.Notes -replace "`r"," " -replace "`n"," " -replace "`t"," "
	
	# Сustom Fields.
		
	$report = New-Object psobject
	$report | Add-Member -type noteproperty -name vCenter -Value $vcName
	$report | Add-Member -type noteproperty -name Datacenter -Value $dcName
	$report | Add-Member -Type NoteProperty -Name "VM Path" -Value $path
	$report | Add-Member -Type NoteProperty -Name "VM Name" -Value $vm.Name
	$report | Add-Member -Type NoteProperty -Name "CPU Count" -Value $vm.NumCpu
	$report | Add-Member -Type NoteProperty -Name "RAM GB" -Value $vm.MemoryGB
	$report | Add-Member -Type NoteProperty -Name "Provisioned Space GB" -Value ([Math]::Round($vm.ProvisionedSpaceGB))
	$report | Add-Member -Type NoteProperty -Name "Used Space GB" -Value ([Math]::Round($vm.UsedSpaceGB))
	$report | Add-Member -Type NoteProperty -Name "DNS Name" -Value $vm.Guest.HostName
	$report | Add-Member -Type NoteProperty -Name "IP" -Value $ip
	$report | Add-Member -Type NoteProperty -Name "OS" -Value $vm.Guest.OSFullName
	$report | Add-Member -Type NoteProperty -Name "Power State" -Value $vm.PowerState
	$report | Add-Member -Type NoteProperty -Name "Notes" -Value $notes
			
	return $report
}