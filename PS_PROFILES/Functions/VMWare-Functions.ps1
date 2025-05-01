# Enable or Disable Hot Add Memory/CPU
# Enable-MemHotAdd $ServerName
# Disable-MemHotAdd $ServerName
# Enable-vCPUHotAdd $ServerName
# Disable-vCPUHotAdd $ServerName




Function Enable-MemHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Disable-MemHotAdd($vm){
$vmview = Get-VM $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Enable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

Function Disable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}

function Get-VMEVCMode {
    <#  .Description
        Code to get VMs' EVC mode and that of the cluster in which the VMs reside.  May 2014, vNugglets.com
        .Example
        Get-VMEVCMode -Cluster myCluster | ?{$_.VMEVCMode -ne $_.ClusterEVCMode}
        Get all VMs in given clusters and return, for each, an object with the VM's- and its cluster's EVC mode, if any
        .Outputs
        PSCustomObject
    #>
    param(
        ## Cluster name pattern (regex) to use for getting the clusters whose VMs to get
        [string]$Cluster_str = ".+"
    )
 
    process {
        ## get the matching cluster View objects
        Get-View -ViewType ClusterComputeResource -Property Name,Summary -Filter @{"Name" = $Cluster_str} | Foreach-Object {
            $viewThisCluster = $_
            ## get the VMs Views in this cluster
            Get-View -ViewType VirtualMachine -Property Name,Runtime.PowerState,Summary.Runtime.MinRequiredEVCModeKey -SearchRoot $viewThisCluster.MoRef | Foreach-Object {
                ## create new PSObject with some nice info
                New-Object -Type PSObject -Property ([ordered]@{
                    Name = $_.Name
                    PowerState = $_.Runtime.PowerState
                    VMEVCMode = $_.Summary.Runtime.MinRequiredEVCModeKey
                    ClusterEVCMode = $viewThisCluster.Summary.CurrentEVCModeKey
                    ClusterName = $viewThisCluster.Name
                })
            } ## end foreach-object
        } ## end foreach-object
    } ## end process
} ## end function
