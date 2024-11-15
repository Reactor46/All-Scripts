﻿Function Set-VMOvfProperty {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
    ===========================================================================
    .DESCRIPTION
        This function updates the OVF Properties (vAppConfig Property) for a VM
    .PARAMETER VM
        VM object returned from Get-VM
    .PARAMETER ovfChanges
        Hashtable mapping OVF property ID to Value
    .EXAMPLE
        $VMNetwork = "sddc-cgw-network-1"
        $VMDatastore = "WorkloadDatastore"
        $VMNetmask = "255.255.255.0"
        $VMGateway = "192.168.1.1"
        $VMDNS = "192.168.1.254"
        $VMNTP = "50.116.52.97"
        $VMPassword = "VMware1!"
        $VMDomain = "vmware.local"
        $VMSyslog = "192.168.1.10"

        $ovfPropertyChanges = @{
            "guestinfo.syslog"=$VMSyslog
            "guestinfo.domain"=$VMDomain
            "guestinfo.gateway"=$VMGateway
            "guestinfo.ntp"=$VMNTP
            "guestinfo.password"=$VMPassword
            "guestinfo.hostname"=$VMIPAddress
            "guestinfo.dns"=$VMDNS
            "guestinfo.ipaddress"=$VMIPAddress
            "guestinfo.netmask"=$VMNetmask
        }

        Set-VMOvfProperty -VM (Get-VM -Name "vesxi65-1-1") -ovfChanges $ovfPropertyChanges
#>
    param(
        [Parameter(Mandatory=$true)]$VM,
        [Parameter(Mandatory=$true)]$ovfChanges
    )

    # Retrieve existing OVF properties from VM
    $vappProperties = $VM.ExtensionData.Config.VAppConfig.Property

    # Create a new Update spec based on the # of OVF properties to update
    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $spec.vAppConfig = New-Object VMware.Vim.VmConfigSpec
    $propertySpec = New-Object VMware.Vim.VAppPropertySpec[]($ovfChanges.count)

    # Find OVF property Id and update the Update Spec
    foreach ($vappProperty in $vappProperties) {
        if($ovfChanges.ContainsKey($vappProperty.Id)) {
            $tmp = New-Object VMware.Vim.VAppPropertySpec
            $tmp.Operation = "edit"
            $tmp.Info = New-Object VMware.Vim.VAppPropertyInfo
            $tmp.Info.Key = $vappProperty.Key
            $tmp.Info.value = $ovfChanges[$vappProperty.Id]
            $propertySpec+=($tmp)
        }
    }
    $spec.VAppConfig.Property = $propertySpec

    Write-Host "Updating OVF Properties ..."
    $task = $vm.ExtensionData.ReconfigVM_Task($spec)
    $task1 = Get-Task -Id ("Task-$($task.value)")
    $task1 | Wait-Task
}

Function Get-VMOvfProperty {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
    ===========================================================================
    .DESCRIPTION
        This function retrieves the OVF Properties (vAppConfig Property) for a VM
    .PARAMETER VM
        VM object returned from Get-VM
    .EXAMPLE
        #Get-VMOvfProperty -VM (Get-VM -Name "vesxi65-1-1")
#>
    param(
        [Parameter(Mandatory=$true)]$VM
    )
    $vappProperties = $VM.ExtensionData.Config.VAppConfig.Property

    $results = @()
    foreach ($vappProperty in $vappProperties | Sort-Object -Property Id) {
        $tmp = [pscustomobject] @{
            Id = $vappProperty.Id;
            Label = $vappProperty.Label;
            Value = $vappProperty.Value;
            Description = $vappProperty.Description;
        }
        $results+=$tmp
    }
    $results
}