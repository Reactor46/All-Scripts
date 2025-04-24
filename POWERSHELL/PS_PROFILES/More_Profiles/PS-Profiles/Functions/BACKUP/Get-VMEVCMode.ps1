function Get-VMEvcMode {
<#  
.SYNOPSIS  
    Gathers information on the EVC status of a VM
.DESCRIPTION 
    Will provide the EVC status for the specified VM
.NOTES  
    Author:  Kyle Ruddy, @kmruddy, thatcouldbeaproblem.com
.PARAMETER Name
    VM name which the function should be ran against
.EXAMPLE
	Get-VMEvcMode -Name vmName
	Retreives the EVC status of the provided VM 
#>
[CmdletBinding()] 
	param(
	[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true)]
        $Name
  	)

    Process {
        $evVM = @()

        if ($name -is [string]) {$evVM += Get-VM -Name $Name -ErrorAction SilentlyContinue}
        elseif ($name -is [array]) {

            if ($name[0] -is [string]) {
                $name | foreach {
                    $evVM += Get-VM -Name $_ -ErrorAction SilentlyContinue
                }
            }
            elseif ($name[0] -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM = $name}

        }
        elseif ($name -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]) {$evVM += $name}
        
        if ($evVM -eq $null) {Write-Warning "No VMs found."}
        else {
            $output = @()
            foreach ($v in $evVM) {

                $report = "" | select Name,EVCMode
                $report.Name = $v.Name
                $report.EVCMode = $v.ExtensionData.Runtime.MinRequiredEVCModeKey
                $output += $report

            }

        return $output

        }

    }

}
