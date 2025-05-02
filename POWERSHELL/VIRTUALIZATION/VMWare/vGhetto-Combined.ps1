

$vcname = "192.168.1.200"
$vcuser = "administrator@vsphere.local"
$vcpass = "J@bb3rJ4w"

#### DO NOT EDIT BEYOND HERE ####

$vcenter = Connect-VIServer $vcname -User $vcuser -Password $vcpass -WarningAction SilentlyContinue

$cluster_ref = Get-Cluster $cluster

$tasks = @()
foreach($esxhost in $esxhosts) {
    Write-Host "Adding $esxhost to $cluster ..."
    Add-VMHost -Name $esxhost -Location $cluster_ref -User $esxuser -Password $esxpass -Force | out-null
}

$spec = New-Object VMware.Vim.ClusterConfigSpecEx
$vsanconfig = New-Object VMware.Vim.VsanClusterConfigInfo
$defaultconfig = New-Object VMware.Vim.VsanClusterConfigInfoHostDefaultInfo
$defaultconfig.AutoClaimStorage = $true
$vsanconfig.DefaultConfig = $defaultconfig
$vsanconfig.enabled = $true
$spec.VsanConfig = $vsanconfig

Write-Host "Enabling VSAN Cluster on $cluster ..."
$task = $cluster_ref.ExtensionData.ReconfigureComputeResource_Task($spec,$true)
$task1 = Get-Task -Id ("Task-$($task.value)")
$task1 | Wait-Task | out-null

Disconnect-VIServer $vcenter -Confirm:$false
# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script demonstrating vSphere MOB Automation using PowerShell
# Reference: http://www.virtuallyghetto.com/2016/07/how-to-automate-vsphere-mob-operations-using-powershell.html

$vc_server = "192.168.1.51"
$vc_username = "administrator@vghetto.local"
$vc_password = "VMware1!"
$mob_url = "https://$vc_server/mob/?moid=VpxSettings&method=queryView"

$secpasswd = ConvertTo-SecureString $vc_password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($vc_username, $secpasswd)

# Ingore SSL Warnings
add-type -TypeDefinition  @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

# Initial login to vSphere MOB using GET and store session using $vmware variable
$results = Invoke-WebRequest -Uri $mob_url -SessionVariable vmware -Credential $credential -Method GET

# Extract hidden vmware-session-nonce which must be included in future requests to prevent CSRF error
# Credit to https://blog.netnerds.net/2013/07/use-powershell-to-keep-a-cookiejar-and-post-to-a-web-form/ for parsing vmware-session-nonce via Powershell
if($results.StatusCode -eq 200) {
    $null = $results -match 'name="vmware-session-nonce" type="hidden" value="?([^\s^"]+)"'
    $sessionnonce = $matches[1]
} else {
    $results
    Write-host "Failed to login to vSphere MOB"
    exit 1
}

# The POST data payload must include the vmware-session-nonce varaible + URL-encoded
$body = @"
vmware-session-nonce=$sessionnonce&name=VirtualCenter.InstanceName
"@

# Second request using a POST and specifying our session from initial login + body request
$results = Invoke-WebRequest -Uri $mob_url -WebSession $vmware -Method POST -Body $body

# Logout out of vSphere MOB
$mob_logout_url = "https://$vc_server/mob/logout"
Invoke-WebRequest -Uri $mob_logout_url -WebSession $vmware -Method GET

# Clean up the results for further processing
# Extract InnerText, split into string array & remove empty lines
$cleanedUpResults = $results.ParsedHtml.body.innertext.split("`n").replace("`"","") | ? {$_.trim() -ne ""}

# Loop through results looking for valuestring which contains the data we want
foreach ($parsedResults in $cleanedUpResults) {
    if($parsedResults -like "valuestring*") {
        $parsedResults.replace("valuestring","")
    }
}
﻿<#
.SYNOPSIS Retrieve the current VMFS Unmap priority for VMFS 6 datastore
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/10/configure-new-automatic-space-reclamation-vmfs-unmap-using-vsphere-6-5-apis.html
.PARAMETER Datastore
  VMFS 6 Datastore to enable or disable VMFS Unamp
.EXAMPLE
  Get-Datastore "mini-local-datastore-hdd" | Get-VMFSUnmap
#>

Function Get-VMFSUnmap {
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [VMware.VimAutomation.ViCore.Impl.V1.DatastoreManagement.DatastoreImpl]$Datastore
     )

     $datastoreInfo = $Datastore.ExtensionData.Info

     if($datastoreInfo -is [VMware.Vim.VmfsDatastoreInfo] -and $datastoreInfo.Vmfs.MajorVersion -eq 6) {
        $datastoreInfo.Vmfs | select Name, UnmapPriority, UnmapGranularity
     } else {
        Write-Host "Not a VMFS Datastore and/or VMFS version is not 6.0"
     }
}

<#
.SYNOPSIS Configure the VMFS Unmap priority for VMFS 6 datastore
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/10/configure-new-automatic-space-reclamation-vmfs-unmap-using-vsphere-6-5-apis.html
.PARAMETER Datastore
  VMFS 6 Datastore to enable or disable VMFS Unamp
.EXAMPLE
  Get-Datastore "mini-local-datastore-hdd" | Set-VMFSUnmap -Enabled $true
.EXAMPLE
  Get-Datastore "mini-local-datastore-hdd" | Set-VMFSUnmap -Enabled $false
#>

Function Set-VMFSUnmap {
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [VMware.VimAutomation.ViCore.Impl.V1.DatastoreManagement.DatastoreImpl]$Datastore,
        [String]$Enabled
     )

    $vmhostView = ($Datastore | Get-VMHost).ExtensionData
    $storageSystem = Get-View $vmhostView.ConfigManager.StorageSystem

    if($Enabled -eq $true) {
        $enableUNMAP = "low"
        $reconfigMessage = "Enabling Automatic VMFS Unmap for $Datastore"
    } else {
        $enableUNMAP = "none"
        $reconfigMessage = "Disabling Automatic VMFS Unmap for $Datastore"
    }

    $uuid = $datastore.ExtensionData.Info.Vmfs.Uuid

    Write-Host "$reconfigMessage ..."
    $storageSystem.UpdateVmfsUnmapPriority($uuid,$enableUNMAP)
}
<#
.SYNOPSIS
This script accepts the name of a VM and the credentials to its
source vCenter Server as well as destination vCenter Server and its
credentials to check if there are any MAC Address conflicts prior to 
issuing a xVC-vMotion of VM (applicable to same and differnet SSO Domain)
.NOTES
File Name : check-vm-mac-conflict.ps1
Author : William Lam - @lamw
Version : 1.0
.LINK
http://www.virtuallyghetto.com/2015/03/duplicate-mac-address-concerns-with-xvc-vmotion-in-vsphere-6-0.html
.LINK
https://github.com/lamw
.INPUTS
sourceVC, sourceVCUsername, sourceVCPassword,destVC, destVCUsername, destVCPassword, vmname
.OUTPUTS
Console output
.PARAMETER sourceVC
The hostname or IP Address of the source vCenter Server
.PARAMETER sourceVCUsername
The username to connect to source vCenter Server
.PARAMETER sourceVCPassword
The password to connect to source vCenter Server
.PARAMETER destVC
The hostname or IP Address of the destination vCenter Server
.PARAMETER destVCUsername
The username to connect to the destination vCenter Server
.PARAMETER destVCPassword
The password to connect to the destination vCenter Server
.PARAMETER vmname
The name of the source VM to check for duplicated MAC Addresses
#>
param
(
   [Parameter(Mandatory=$true)]
   [string]
   $sourceVC,
   [Parameter(Mandatory=$true)]
   [string]
   $sourceVCUsername,
   [Parameter(Mandatory=$true)]
   [string]
   $sourceVCPassword,
   [Parameter(Mandatory=$true)]
   [string]
   $destVC,
   [Parameter(Mandatory=$true)]
   [string]
   $destVCUsername,
   [Parameter(Mandatory=$true)]
   [string]
   $destVCPassword,
   [Parameter(Mandatory=$true)]
   [string]
   $vmname
);

# Debug
#$sourceVC = "vcenter60-1.primp-industries.com"
#$sourceVCUsername = "administrator@vghetto.local"
#$sourceVCPassword = "VMware1!"
#
#$destVC = "vcenter60-2.primp-industries.com"
#$destVCUsername = "administrator@vghetto.local"
#$destVCPassword = "VMware1!"
#
#$vmname = "VM1"

# Connect to Source vCenter Server
$sourceVCConn = Connect-VIServer -Server $sourceVC -user $sourceVCUsername -password $sourceVCPassword

# Connect to Destination vCenter Server
$destVCConn = Connect-VIServer -Server $destVC -user $destVCUsername -password $destVCpassword

# Retrieve Source VM MAC Addresses
$sourceVMMACs = (Get-NetworkAdapter -Server $sourceVCConn -VM $vmname).MacAddress

# Retrieve ALL VM Mac Addresses from Destination vCenter Server
$allVMMacs = @{}
$vms = Get-View -Server $destVCConn -ViewType VirtualMachine -Property Name,Config.Hardware.Device -Filter @{"Config.Template" = "False"}
foreach ($vm in $vms) {
	$devices = $vm.Config.Hardware.Device
	foreach ($device in $devices) {
		if($device -is  [VMware.Vim.VirtualEthernetCard]) {
			# Store hash of Mac to VM to be used later for later comparison
			$allVMMacs.add($device.MacAddress,$vm.Name)
		}
	}
}

# Disconnect from Source/Dest vCenter Servers as it is no longer needed
Disconnect-VIServer -Server $sourceVCConn -Confirm:$false
Disconnect-VIServer -Server $destVCConn -Confirm:$false

# Check for duplicated MAC Addresses in destionation vCenter Server
Write-Host "`nChecking to see if there are MAC Address conflicts with" $vmname "at destination vCenter Server...`n"

foreach ($mac in $sourceVMMACs) {
	if($allVMMacs[$mac]) {
		Write-Host $allVMMacs[$mac] "also has MAC Address: $mac"
	}
}
# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script to enable MultiWriter Flag for existing VMDK in vSphere 6.x
# Reference: http://www.virtuallyghetto.com/2015/10/new-method-of-enabling-multiwriter-vmdk-flag-in-vsphere-6-0-update-1.html

$vcname = "192.168.1.150"
$vcuser = "administrator@vghetto.local"
$vcpass = "VMware1!"

$vmName = "vm-1"
$diskName = "Hard disk 2"

#### DO NOT EDIT BEYOND HERE ####

$server = Connect-VIServer -Server $vcname -User $vcuser -Password $vcpass

# Retrieve VM and only its Devices
$vm = Get-View -Server $server -ViewType VirtualMachine -Property Name,Config.Hardware.Device -Filter @{"Name" = $vmName}

# Array of Devices on VM
$vmDevices = $vm.Config.Hardware.Device

# Find the Virtual Disk that we care about
foreach ($device in $vmDevices) {
	if($device -is  [VMware.Vim.VirtualDisk] -and $device.deviceInfo.Label -eq $diskName) {
		$diskDevice = $device
		$diskDeviceBaking = $device.backing
		break
	}
}

# Create VM Config Spec to Edit existing VMDK & Enable Multi-Writer Flag
$spec = New-Object VMware.Vim.VirtualMachineConfigSpec
$spec.deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec
$spec.deviceChange[0].operation = 'edit'
$spec.deviceChange[0].device = New-Object VMware.Vim.VirtualDisk
$spec.deviceChange[0].device = $diskDevice
$spec.DeviceChange[0].device.backing = New-Object VMware.Vim.VirtualDiskFlatVer2BackingInfo
$spec.DeviceChange[0].device.backing = $diskDeviceBaking
$spec.DeviceChange[0].device.Backing.Sharing = "sharingMultiWriter"

Write-Host "`nEnabling Multiwriter flag on on VMDK:" $diskName "for VM:" $vmname
$task = $vm.ReconfigVM_Task($spec)
$task1 = Get-Task -Id ("Task-$($task.value)")
$task1 | Wait-Task

Disconnect-VIServer $server -Confirm:$false
# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script to configure per-VMDK IOPS Reservations on a VM in vSphere 6.0
# Reference: http://www.virtuallyghetto.com/2015/05/configuring-per-vmdk-iops-reservations-in-vsphere-6-0

$server = Connect-VIServer -Server 192.168.1.60 -User administrator@vghetto.local -Password VMware1!

# Fill out with your VM Name, Disk Label & IOPS Reservation
$vmName = "Photon"
$diskName = "Hard disk 1"
$iopsReservation = "2000"

### DO NOT EDIT BEYOND HERE ###

# Retrieve VM and only its Devices
$vm = Get-View -Server $server -ViewType VirtualMachine -Property Name,Config.Hardware.Device -Filter @{"Name" = $vmName}

# Array of Devices on VM
$vmDevices = $vm.Config.Hardware.Device

# Find the Virtual Disk that we care about
foreach ($device in $vmDevices) {
	if($device -is  [VMware.Vim.VirtualDisk] -and $device.deviceInfo.Label -eq $diskName) {
			$diskDevice = $device
			break
	}
}

# Create VM Config Spec
$spec = New-Object VMware.Vim.VirtualMachineConfigSpec
$spec.deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec
$spec.deviceChange[0].operation = 'edit'
$spec.deviceChange[0].device = New-Object VMware.Vim.VirtualDisk
$spec.deviceChange[0].device = $diskDevice
$spec.deviceChange[0].device.storageIOAllocation.reservation = $iopsReservation
Write-Host "Configuring IOPS Reservation:" $iopsReservation "on VMDK:" $diskName "for VM:" $vmname
$vm.ReconfigVM($spec)

# Uncomment the following snippet if you wish to verify as part of the reconfiguration operation 

#$vm.UpdateViewData()
#$vmDevices = $vm.Config.Hardware.Device
#foreach ($device in $vmDevices) {
#	if($device -is  [VMware.Vim.VirtualDisk] -and $device.deviceInfo.Label -eq $diskName) {
#			$device.storageIOAllocation
#	}
#}

Disconnect-VIServer $server -Confirm:$false$ActivationKey = "<FILL ME>"
$HCXServer = "mgmt-hcxm-02.cpbu.corp"
$VAMIUsername = "admin"
$VAMIPassword = "VMware1!"
$VIServer = "mgmt-vcsa-01.cpbu.corp"
$VIUsername = "administrator@vsphere.local"
$VIPassword = "VMware1!"
$NSXServer = "mgmt-nsxm-01.cpbu.corp"
$NSXUsername = "admin"
$NSXPassword = "VMware1!"

Connect-HcxVAMI -Server $HCXServer -Username $VAMIUsername -Password $VAMIPassword

Set-HcxLicense -LicenseKey $ActivationKey

Set-HcxVCConfig -VIServer $VIServer -VIUsername $VIUsername -VIPassword $VIPassword -PSCServer $VIServer

Set-HcxNSXConfig -NSXServer $NSXServer -NSXUsername $NSXUsername -NSXPassword $NSXPassword

Set-HcxLocation -City "Santa Barbara" -Country "United States of America"

Set-HcxRoleMapping -SystemAdminGroup @("vsphere.local\Administrators","cpbu.corp\Administrators") -EnterpriseAdminGroup @("vsphere.local\Administrators","cpbu.corp\Administrators")
# William Lam
# www.virtuallyghetto.com

$vcname = "192.168.1.150"
$vcuser = "administrator@vghetto.local"
$vcpass = "VMware1!"

$ovffile = "Z:\Desktop\Nested_ESXi_Appliance.ovf"

$cluster = "MacMini-Cluster"
$vmnetwork = "VM Network"
$datastore = "mini-local-datastore-1"
$iprange = "192.168.1"
$netmask = "255.255.255.0"
$gateway = "192.168.1.1"
$dns = "192.168.1.1"
$dnsdomain = "primp-industries.com"
$ntp = "192.168.1.1"
$syslog = "192.168.1.150"
$password = "VMware1!"
$ssh = "True"

#### DO NOT EDIT BEYOND HERE ####

$vcenter = Connect-VIServer $vcname -User $vcuser -Password $vcpass -WarningAction SilentlyContinue

$datastore_ref = Get-Datastore -Name $datastore
$network_ref = Get-VirtualPortGroup -Name $vmnetwork
$cluster_ref = Get-Cluster -Name $cluster
$vmhost_ref = $cluster_ref | Get-VMHost | Select -First 1

$ovfconfig = Get-OvfConfiguration $ovffile
$ovfconfig.NetworkMapping.VM_Network.value = $network_ref

190..192 | Foreach {
    $ipaddress = "$iprange.$_"
    # Try to perform DNS lookup
    try {
        $vmname = ([System.Net.Dns]::GetHostEntry($ipaddress).HostName).split(".")[0]
    }
    Catch [system.exception]
    {
        $vmname = "vesxi-vsan-$ipaddress"
    }
    $ovfconfig.common.guestinfo.hostname.value = $vmname
    $ovfconfig.common.guestinfo.ipaddress.value = $ipaddress
    $ovfconfig.common.guestinfo.netmask.value = $netmask
    $ovfconfig.common.guestinfo.gateway.value = $gateway
    $ovfconfig.common.guestinfo.dns.value = $dns
    $ovfconfig.common.guestinfo.domain.value = $dnsdomain
    $ovfconfig.common.guestinfo.ntp.value = $ntp
    $ovfconfig.common.guestinfo.syslog.value = $syslog
    $ovfconfig.common.guestinfo.password.value = $password
    $ovfconfig.common.guestinfo.ssh.value = $ssh

    # Deploy the OVF/OVA with the config parameters
    Write-Host "Deploying $vmname ..."
    $vm = Import-VApp -Source $ovffile -OvfConfiguration $ovfconfig -Name $vmname -Location $cluster_ref -VMHost $vmhost_ref -Datastore $datastore_ref -DiskStorageFormat thin
    $vm | Start-Vm -RunAsync | Out-Null
}

Disconnect-VIServer $vcenter -Confirm:$false
﻿# Load OVF/OVA configuration into a variable
$ovffile = "C:\Users\william\Desktop\VMware-HCX-Enterprise-3.5.1-10027070.ova"
$ovfconfig = Get-OvfConfiguration $ovffile

# vSphere Cluster + VM Network configurations
$Cluster = "Cluster-01"
$VMName = "MGMT-HCXM-02"
$VMNetwork = "SJC-CORP-MGMT-EP"
$HCXAddressToVerify = "mgmt-hcxm-02.cpbu.corp"

$VMHost = Get-Cluster $Cluster | Get-VMHost | Sort MemoryGB | Select -first 1
$Datastore = $VMHost | Get-datastore | Sort FreeSpaceGB -Descending | Select -first 1
$Network = Get-VDPortGroup -Name $VMNetwork

# Fill out the OVF/OVA configuration parameters

# vSphere Portgroup Network Mapping
$ovfconfig.NetworkMapping.VSMgmt.value = $Network

# IP Address
$ovfConfig.common.mgr_ip_0.value = "172.17.31.50"

# Netmask
$ovfConfig.common.mgr_prefix_ip_0.value = "24"

# Gateway
$ovfConfig.common.mgr_gateway_0.value = "172.17.31.253"

# DNS Server
$ovfConfig.common.mgr_dns_list.value = "172.17.31.5"

# DNS Domain
$ovfConfig.common.mgr_domain_search_list.value  = "cpbu.corp"

# Hostname
$ovfconfig.Common.hostname.Value = "mgmt-hcxm-02.cpbu.corp"

# NTP
$ovfconfig.Common.mgr_ntp_list.Value = "172.17.31.5"

# SSH
$ovfconfig.Common.mgr_isSSHEnabled.Value = $true

# Password
$ovfconfig.Common.mgr_cli_passwd.Value = "VMware1!"
$ovfconfig.Common.mgr_root_passwd.Value = "VMware1!"

# Deploy the OVF/OVA with the config parameters
Write-Host -ForegroundColor Green "Deploying HCX Manager OVA ..."
$vm = Import-VApp -Source $ovffile -OvfConfiguration $ovfconfig -Name $VMName -VMHost $vmhost -Datastore $datastore -DiskStorageFormat thin

# Power On the HCX Manager VM after deployment
Write-Host -ForegroundColor Green "Powering on HCX Manager ..."
$vm | Start-VM -Confirm:$false | Out-Null

# Waiting for HCX Manager to initialize
while(1) {
    try {
        if($PSVersionTable.PSEdition -eq "Core") {
            $requests = Invoke-WebRequest -Uri "https://$($HCXAddressToVerify):9443" -Method GET -SkipCertificateCheck -TimeoutSec 5
        } else {
            $requests = Invoke-WebRequest -Uri "https://$($HCXAddressToVerify):9443" -Method GET -TimeoutSec 5
        }
        if($requests.StatusCode -eq 200) {
            Write-Host -ForegroundColor Green "HCX Manager is now ready to be configured!"
            break
        }
    }
    catch {
        Write-Host -ForegroundColor Yellow "HCX Manager is not ready yet, sleeping for 120 seconds ..."
        sleep 120
    }
}# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script to disable/enable vMotion capability for a specific VM
# Reference: http://www.virtuallyghetto.com/2016/07/how-to-easily-disable-vmotion-cross-vcenter-vmotion-for-a-particular-virtual-machine.html

Function Enable-vSphereMethod {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [String]$vmmoref,
    [string]$vc_server,
    [String]$vc_username,
    [String]$vc_password,
    [String]$enable_method
    )

    $secpasswd = ConvertTo-SecureString $vc_password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($vc_username, $secpasswd)

    # vSphere MOB URL to private enableMethods
    $mob_url = "https://$vc_server/mob/?moid=AuthorizationManager&method=enableMethods"

# Ingore SSL Warnings
add-type -TypeDefinition  @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

    # Initial login to vSphere MOB using GET and store session using $vmware variable
    $results = Invoke-WebRequest -Uri $mob_url -SessionVariable vmware -Credential $credential -Method GET

    # Extract hidden vmware-session-nonce which must be included in future requests to prevent CSRF error
    # Credit to https://blog.netnerds.net/2013/07/use-powershell-to-keep-a-cookiejar-and-post-to-a-web-form/ for parsing vmware-session-nonce via Powershell
    if($results.StatusCode -eq 200) {
        $null = $results -match 'name="vmware-session-nonce" type="hidden" value="?([^\s^"]+)"'
        $sessionnonce = $matches[1]
    } else {
        Write-host "Failed to login to vSphere MOB"
        exit 1
    }

    # The POST data payload must include the vmware-session-nonce variable + URL-encoded
    $body = @"
vmware-session-nonce=$sessionnonce&entity=%3Centity+type%3D%22ManagedEntity%22+xsi%3Atype%3D%22ManagedObjectReference%22%3E$vmmoref%3C%2Fentity%3E%0D%0A&method=%3Cmethod%3E$enable_method%3C%2Fmethod%3E
"@

    # Second request using a POST and specifying our session from initial login + body request
    $results = Invoke-WebRequest -Uri $mob_url -WebSession $vmware -Method POST -Body $body

    # Logout out of vSphere MOB
    $mob_logout_url = "https://$vc_server/mob/logout"
    Invoke-WebRequest -Uri $mob_logout_url -WebSession $vmware -Method GET    
}

Function Disable-vSphereMethod {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [String]$vmmoref,
    [string]$vc_server,
    [String]$vc_username,
    [String]$vc_password,
    [String]$disable_method
    )

    $secpasswd = ConvertTo-SecureString $vc_password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($vc_username, $secpasswd)

    # vSphere MOB URL to private disableMethods
    $mob_url = "https://$vc_server/mob/?moid=AuthorizationManager&method=disableMethods"

# Ingore SSL Warnings
add-type -TypeDefinition  @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

    # Initial login to vSphere MOB using GET and store session using $vmware variable
    $results = Invoke-WebRequest -Uri $mob_url -SessionVariable vmware -Credential $credential -Method GET

    # Extract hidden vmware-session-nonce which must be included in future requests to prevent CSRF error
    # Credit to https://blog.netnerds.net/2013/07/use-powershell-to-keep-a-cookiejar-and-post-to-a-web-form/ for parsing vmware-session-nonce via Powershell
    if($results.StatusCode -eq 200) {
        $null = $results -match 'name="vmware-session-nonce" type="hidden" value="?([^\s^"]+)"'
        $sessionnonce = $matches[1]
    } else {
        Write-host "Failed to login to vSphere MOB"
        exit 1
    }

    # The POST data payload must include the vmware-session-nonce variable + URL-encoded
    $body = @"
vmware-session-nonce=$sessionnonce&entity=%3Centity+type%3D%22ManagedEntity%22+xsi%3Atype%3D%22ManagedObjectReference%22%3E$vmmoref%3C%2Fentity%3E%0D%0A%0D%0A&method=%3CDisabledMethodRequest%3E%0D%0A+++%3Cmethod%3E$disable_method%3C%2Fmethod%3E%0D%0A%3C%2FDisabledMethodRequest%3E%0D%0A%0D%0A&sourceId=self
"@

    # Second request using a POST and specifying our session from initial login + body request
    $results = Invoke-WebRequest -Uri $mob_url -WebSession $vmware -Method POST -Body $body
}

### Sample Usage of Enable/Disable functions ###

$vc_server = "192.168.1.51"
$vc_username = "administrator@vghetto.local"
$vc_password = "VMware1!"
$vm_name = "TestVM-1"
$method_name = "MigrateVM_Task"

# Connect to vCenter Server
$server = Connect-VIServer -Server $vc_server -User $vc_username -Password $vc_password

$vm = Get-VM -Name $vm_name
$vm_moref = (Get-View $vm).MoRef.Value

#Disable-vSphereMethod -vc_server $vc_server -vc_username $vc_username -vc_password $vc_password -vmmoref $vm_moref -disable_method $method_name

#Enable-vSphereMethod -vc_server $vc_server -vc_username $vc_username -vc_password $vc_password -vmmoref $vm_moref -enable_method $method_name

# Disconnect from vCenter Server
Disconnect-viserver $server -confirm:$false
﻿Function Get-ESXiBootDevice {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function identifies how an ESXi host was booted up along with its boot
        device (if applicable). This supports both local installation to Auto Deploy as
        well as Boot from SAN.
    .PARAMETER VMHostname
        The name of an individual ESXi host managed by vCenter Server
    .EXAMPLE
        Get-ESXiBootDevice
    .EXAMPLE
        Get-ESXiBootDevice -VMHostname esxi-01
#>
    param(
        [Parameter(Mandatory=$false)][String]$VMHostname
    )

    if($VMHostname) {
        $vmhosts = Get-VMhost -Name $VMHostname
    } else {
        $vmhosts = Get-VMHost
    }

    $results = @()
    foreach ($vmhost in ($vmhosts | Sort-Object -Property Name)) {
        $esxcli = Get-EsxCli -V2 -VMHost $vmhost
        $bootDetails = $esxcli.system.boot.device.get.Invoke()

        # Check to see if ESXi booted over the network
        $networkBoot = $false
        if($bootDetails.BootNIC) {
            $networkBoot = $true
            $bootDevice = $bootDetails.BootNIC
        } elseif ($bootDetails.StatelessBootNIC) {
            $networkBoot = $true
            $bootDevice = $bootDetails.StatelessBootNIC
        }

        # If ESXi booted over network, check to see if deployment
        # is Stateless, Stateless w/Caching or Stateful
        if($networkBoot) {
            $option = $esxcli.system.settings.advanced.list.CreateArgs()
            $option.option = "/UserVars/ImageCachedSystem"
            try {
                $optionValue = $esxcli.system.settings.advanced.list.Invoke($option)
            } catch {
                $bootType = "stateless"
            }
            $bootType = $optionValue.StringValue
        }

        # Loop through all storage devices to identify boot device
        $devices = $esxcli.storage.core.device.list.Invoke()
        $foundBootDevice = $false
        foreach ($device in $devices) {
            if($device.IsBootDevice -eq $true) {
                $foundBootDevice = $true

                if($device.IsLocal -eq $true -and $networkBoot -and $bootType -ne "stateful") {
                    $bootType = "stateless caching"
                } elseif($device.IsLocal -eq $true -and $networkBoot -eq $false) {
                    $bootType = "local"
                } elseif($device.IsLocal -eq $false -and $networkBoot -eq $false) {
                    $bootType = "remote"
                }

                $bootDevice = $device.Device
                $bootModel = $device.Model
                $bootVendor = $device.VEndor
                $bootSize = $device.Size
                $bootIsSAS = $device.IsSAS
                $bootIsSSD = $device.IsSSD
                $bootIsUSB = $device.IsUSB
            }
        }

        # Pure Stateless (e.g. No USB or Disk for boot)
        if($networkBoot-and $foundBootDevice -eq $false) {
            $bootModel = "N/A"
            $bootVendor = "N/A"
            $bootSize = "N/A"
            $bootIsSAS = "N/A"
            $bootIsSSD = "N/A"
            $bootIsUSB = "N/A"
        }

        $tmp = [pscustomobject] @{
            Host = $vmhost.Name;
            Device = $bootDevice;
            BootType = $bootType;
            Vendor = $bootVendor;
            Model = $bootModel;
            SizeMB = $bootSize;
            IsSAS = $bootIsSAS;
            IsSSD = $bootIsSSD;
            IsUSB = $bootIsUSB;
        }
        $results+=$tmp
    }
    $results | FT -AutoSize
}﻿Function Get-ESXiDPC {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retreives the current disabled TLS protocols for all ESXi
        hosts within a vSphere Cluster
    .SYNOPSIS
        Returns current disabled TLS protocols for Hostd, Authd, sfcbd & VSANVP/IOFilter 
    .PARAMETER Cluster
        The name of the vSphere Cluster
    .EXAMPLE
        Get-ESXiDPC -Cluster VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$Cluster
    )

    $debug = $false

    Function Get-SFCBDConf {
        param(
            [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$VMHost
        )

        $url = "https://$vmhost/host/sfcb.cfg"

        $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

        $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
        $spec.Method = "httpGet"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie = New-Object System.Net.Cookie
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhost.name
        $websession.Cookies.Add($cookie)
        $result = Invoke-WebRequest -Uri $url -WebSession $websession
        $sfcbConf = $result.content
        
        # Extract the TLS fields if they exists
        $sfcbResults = @()
        $usingDefault = $true
        foreach ($line in $sfcbConf.Split("`n")) {
            if($line -match "enableTLSv1:") {
                ($key,$value) = $line.Split(":")
                if($value -match "false") {
                    $sfcbResults+="tlsv1"
                }
                $usingDefault = $false
            }
            if($line -match "enableTLSv1_1:") {
                ($key,$value) = $line.Split(":")
                if($value -match "false") {
                    $sfcbResults+="tlsv1.1"
                }
                $usingDefault = $false
            }
            if($line -match "enableTLSv1_2:") {
                ($key,$value) = $line.Split(":")
                if($value -match "false") {
                    $sfcbResults+="tlsv1.2"
                }
                $usingDefault = $false
            }
        }
        if($usingDefault -or ($sfcbResults.Length -eq 0)) {
            $sfcbResults = "tlsv1,tlsv1.1,sslv3"
            return $sfcbResults
        } else {
            $sfcbResults+="sslv3"
            return $sfcbResults -join ","
        }
    }

    $results = @()
    foreach ($vmhost in (Get-Cluster -Name $Cluster | Get-VMHost)) {
        if( ($vmhost.ApiVersion -eq "6.0" -and (Get-AdvancedSetting -Entity $vmhost -Name "Misc.HostAgentUpdateLevel").value -eq "3") -or ($vmhost.ApiVersion -eq "6.5") ) {
            $esxiVersion = ($vmhost.ApiVersion) + " Update " + (Get-AdvancedSetting -Entity $vmhost -Name "Misc.HostAgentUpdateLevel").value
            
            $vps = (Get-AdvancedSetting -Entity $vmhost -Name "UserVars.ESXiVPsDisabledProtocols" -ErrorAction SilentlyContinue).value
            # ESXi 6.5 - UserVars.ESXiVPsDisabledProtocols covers both VPs+rHTTP
            if($vmhost.ApiVersion -eq "6.5") {
                $rhttpProxy = $vps
                # Only TLS 1.2 is enabled 
                $vmauth = "tlsv1,tlsv1.1,sslv3"
            } else {
                $rhttpProxy = (Get-AdvancedSetting -Entity $vmhost -Name "UserVars.ESXiRhttpproxyDisabledProtocols" -ErrorAction SilentlyContinue).value
                $vmauth = (Get-AdvancedSetting -Entity $vmhost -Name "UserVars.VMAuthdDisabledProtocols" -ErrorAction SilentlyContinue).value
            }
            $sfcbd = Get-SFCBDConf -vmhost $vmhost

            $hostTLSSettings = [pscustomobject] @{
                vmhost = $vmhost.name;
                version = $esxiVersion;
                hostd = $rhttpProxy;
                authd = $vmauth;
                sfcbd = $sfcbd
                ioFilterVSANVP = $vps
            }
            $results+=$hostTLSSettings
        }
    }
    Write-Host -NoNewline "`nDisabled Protocols on all ESXi hosts:"
    $results
}

Function Set-ESXiDPC {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function configures the TLS protocols to disable for all 
        ESXi hosts within a vSphere Cluster
    .SYNOPSIS
        Configures the disabled TLS protocols for Hostd, Authd, sfcbd & VSANVP/IOFilter 
    .PARAMETER Cluster
        The name of the vSphere Cluster
    .EXAMPLE
        Set-ESXiDPC -Cluster VSAN-Cluster -TLS1 $true -TLS1_1 $true -TLS1_2 $false -SSLV3 $true
#>
    param(
        [Parameter(Mandatory=$true)][String]$Cluster,
        [Parameter(Mandatory=$true)][Boolean]$TLS1,
        [Parameter(Mandatory=$true)][Boolean]$TLS1_1,
        [Parameter(Mandatory=$true)][Boolean]$TLS1_2,
        [Parameter(Mandatory=$true)][Boolean]$SSLV3
    )

    Function UpdateSFCBConfig {
        param(
            [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$VMHost
        )

        $url = "https://$vmhost/host/sfcb.cfg"

        $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

        $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
        $spec.Method = "httpGet"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie = New-Object System.Net.Cookie
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhost.name
        $websession.Cookies.Add($cookie)
        $result = Invoke-WebRequest -Uri $url -WebSession $websession
        $sfcbConf = $result.content
        
        # Download the current sfcb.cfg and ignore existing TLS configuration
        $sfcbResults = ""
        foreach ($line in $sfcbConf.Split("`n")) {
            if($line -notmatch "enableTLSv1:" -and $line -notmatch "enableTLSv1_1:" -and $line -notmatch "enableTLSv1_2:" -and $line -ne "") {
                $sfcbResults+="$line`n"
            }
        }
        
        # Append the TLS protocols based on user input to the configuration file
        $sfcbResults+="enableTLSv1: " + (!$TLS1).ToString().ToLower() + "`n"
        $sfcbResults+="enableTLSv1_1: " + (!$TLS1_1).ToString().ToLower() + "`n"
        $sfcbResults+="enableTLSv1_2: " + (!$TLS1_2).ToString().ToLower() +"`n"

        # Create HTTP PUT spec
        $spec.Method = "httpPut"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

        # Upload sfcb.cfg back to ESXi host
        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhost.name
        $websession.Cookies.Add($cookie)
        $result = Invoke-WebRequest -Uri $url -WebSession $websession -Body $sfcbResults -Method Put -ContentType "plain/text"
        if($result.StatusCode -eq 200) {
            Write-Host "`tSuccessfully updated sfcb.cfg file"
        } else {
            Write-Host "Failed to upload sfcb.cfg file"
            break
        }
    }

    # Build TLS string based on user input for setting ESXi Advanced Settings
    if($TLS1 -and $TLS1_1 -and $TLS1_2 -and $SSLV3) {
        Write-Host -ForegroundColor Red "Error: You must at least enable one of the TLS protocols"
        break
    }

    $tlsString = @()
    if($TLS1) { $tlsString += "tlsv1" }
    if($TLS1_1) { $tlsString += "tlsv1.1" }
    if($TLS1_2) { $tlsString += "tlsv1.2" }
    if($SSLV3) { $tlsString += "sslv3" }
    $tlsString = $tlsString -join ","

    Write-Host "`nDisabling the following TLS protocols: $tlsString on ESXi hosts ...`n"
    foreach ($vmhost in (Get-Cluster -Name $Cluster | Get-VMHost)) {
        if( ($vmhost.ApiVersion -eq "6.0" -and (Get-AdvancedSetting -Entity $vmhost -Name "Misc.HostAgentUpdateLevel").value -eq "3") -or ($vmhost.ApiVersion -eq "6.5") ) {
            Write-Host "Updating $vmhost ..."

            Write-Host "`tUpdating sfcb.cfg ..."
            UpdateSFCBConfig -vmhost $vmhost

            if($vmhost.ApiVersion -eq "6.0") {
                Write-Host "`tUpdating UserVars.ESXiRhttpproxyDisabledProtocols ..."
                Get-AdvancedSetting -Entity $vmhost -Name "UserVars.ESXiRhttpproxyDisabledProtocols" | Set-AdvancedSetting -Value $tlsString -Confirm:$false | Out-Null

                Write-Host "`tUpdating UserVars.VMAuthdDisabledProtocols ..."
                Get-AdvancedSetting -Entity $vmhost -Name "UserVars.VMAuthdDisabledProtocols" | Set-AdvancedSetting -Value $tlsString -Confirm:$false | Out-Null
            }
            Write-Host "`tUpdating UserVars.ESXiVPsDisabledProtocols ..."
            Get-AdvancedSetting -Entity $vmhost -Name "UserVars.ESXiVPsDisabledProtocols" | Set-AdvancedSetting -Value $tlsString -Confirm:$false | Out-Null
        }
    }
}
﻿<#
.SYNOPSIS Retrieve the installation date of an ESXi host
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/10/super-easy-way-of-getting-esxi-installation-date-in-vsphere-6-5.html
.PARAMETER Vmhost
  ESXi host to query installation date
.EXAMPLE
  Get-Vmhost "mini" | Get-ESXInstallDate
#>

Function Get-ESXInstallDate {
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$Vmhost
     )

     if($Vmhost.Version -eq "6.5.0") {
        $imageManager = Get-View ($Vmhost.ExtensionData.ConfigManager.ImageConfigManager)
        $installDate = $imageManager.installDate()

        Write-Host "$Vmhost was installed on $installDate"
     } else {
        Write-Host "ESXi must be running 6.5"
     }
}
﻿<#
.SYNOPSIS Retrieve the installation date of an ESXi host
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.PARAMETER Vmhost
  ESXi host to query installed ESXi VIBs
.EXAMPLE
  Get-ESXInstalledVibs -Vmhost (Get-Vmhost "mini")
.EXAMPLE
  Get-ESXInstalledVibs -Vmhost (Get-Vmhost "mini") -vibname vsan
#>

Function Get-ESXInstalledVibs {
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$Vmhost,
        [Parameter(
            Mandatory=$false,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [String]$vibname=""
     )

     $imageManager = Get-View ($Vmhost.ExtensionData.ConfigManager.ImageConfigManager)
     $vibs = $imageManager.fetchSoftwarePackages()

     foreach ($vib in $vibs) {
        if($vibname -ne "") {
            if($vib.name -eq $vibname) {
                return $vib | select Name, Version, Vendor, CreationDate, Summary
            }
        } else {
            $vib | select Name, Version, Vendor, CreationDate, Summary
        }
     }
}
﻿$esxiVersions = @("5.1.0", "5.5.0", "6.0.0", "6.5.0", "6.7.0")
$pathToStoreMetdataFile = $env:TMP

Add-Type -Assembly System.IO.Compression.FileSystem

Write-Host "Downloading ESXi Metadata Files ..."
foreach ($esxiVersion in $esxiVersions) {
    $metadataUrl = "https://hostupdate.vmware.com/software/VUM/PRODUCTION/main/esx/vmw/vmw-ESXi-$esxiVersion-metadata.zip"
    $metadataDownloadPath = $pathToStoreMetdataFile + "\" + $esxiVersion + ".zip"
    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile($metadataUrl,$metadataDownloadPath)

    #https://stackoverflow.com/a/41575369
    $zip = [IO.Compression.ZipFile]::OpenRead($metadataDownloadPath)
    $metadataFileExtractionPath = $pathToStoreMetdataFile + "\$esxiVersion.xml"
    $zip.Entries | where {$_.Name -like 'vmware.xml'} | foreach {[System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $metadataFileExtractionPath, $true)}
    $zip.Dispose()
    Remove-Item -Path $metadataDownloadPath -Force
}

Write-Host "Processing ESXi Metadata Files ..."
$esxiBulletinCVEesults = @()
foreach ($esxiVersion in $esxiVersions) {
    $metadataFileExtractionPath = $pathToStoreMetdataFile + "\$esxiVersion.xml"
    [xml]$XmlDocument = Get-Content -Path $metadataFileExtractionPath

    Write-Host "Extracting KB Information & CVE URLs for $esxiVersion ..." 
    foreach ($bulletin in $XmlDocument.metadataResponse.bulletin) {
        if($bulletin.category -eq "security") {
            $bulletinId = $bulletin.id
            $kbId = ($bulletin.kbUrl).Replace("http://kb.vmware.com/kb/","")

            $results = Invoke-WebRequest -Uri https://kb.vmware.com/articleview?docid=$kbId -UseBasicParsing

            $cveIds = @()
            foreach ($link in $results.Links) {
                if($link.href -match "CVE") {
                    $cveIds += ($link.href).Replace("http://cve.mitre.org/cgi-bin/cvename.cgi?name=","")
                }
            }

            if($cveIds) {
                foreach ($cveId in $cveIds) {
                    # CVE API to retrieve CVE details
                    $results = Invoke-WebRequest -Uri  http://cve.circl.lu/api/cve/$cveId -UseBasicParsing
                    $jsonResults = $results.Content | ConvertFrom-Json
                    $cvssScore = $jsonResults.cvss
                    $cvssComplexity = $jsonResults.access.complexity

                    if($cvssScore -eq $null) {
                        $cvssScore = "N/A"
                    }
                    if($cvssComplexity -eq $null) {
                        $cvssComplexity = "N/A"
                    }

                    $tmp = [PSCustomObject] @{
                        Bulletin = $bulletinId;
                        CVEId = $cveId;
                        CVSSScore = $cvssScore;
                        CVSSComplexity = $cvssComplexity;
                    }
                    $esxiBulletinCVEesults += $tmp
                }
            }
        }
    }
}

$esxiBulletinCVEesults﻿Function Get-VMConsoleURL {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function generates the HTML5 VM Console URL (default) as seen when using the
        vSphere Web Client connected to a vCenter Server 6.5 environment. You must
        already be logged in for the URL to be valid or else you will be prompted
        to login before being re-directed to VM Console. You have option of also generating
        either the Standalone VMRC or WebMKS URL using additional parameters
    .PARAMETER VMName
        The name of the VM
    .PARAMETER webmksUrl
        Set to true to generate the WebMKS URL instead (e.g. wss://<host>/ticket/<ticket>)
    .PARAMETER vmrcUrl
        Set to true to generate the VMRC URL instead (e.g. vmrc://...)
    .EXAMPLE
        Get-VMConsoleURL -VMName "Embedded-VCSA1"
    .EXAMPLE
        Get-VMConsoleURL -VMName "Embedded-VCSA1" -vmrcUrl $true
    .EXAMPLE
        Get-VMConsoleURL -VMName "Embedded-VCSA1" -webmksUrl $true
        #>
    param(
        [Parameter(Mandatory=$true)][String]$VMName,
        [Parameter(Mandatory=$false)][Boolean]$vmrcUrl,
        [Parameter(Mandatory=$false)][Boolean]$webmksUrl
    )

    Function Get-SSLThumbprint {
        param(
        [Parameter(
            Position=0,
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [Alias('FullName')]
        [String]$URL
        )

        add-type @"
            using System.Net;
            using System.Security.Cryptography.X509Certificates;

                public class IDontCarePolicy : ICertificatePolicy {
                public IDontCarePolicy() {}
                public bool CheckValidationResult(
                    ServicePoint sPoint, X509Certificate cert,
                    WebRequest wRequest, int certProb) {
                    return true;
                }
                }
"@
        [System.Net.ServicePointManager]::CertificatePolicy = new-object IDontCarePolicy

        # Need to connect using simple GET operation for this to work
        Invoke-RestMethod -Uri $URL -Method Get | Out-Null

        $ENDPOINT_REQUEST = [System.Net.Webrequest]::Create("$URL")
        $SSL_THUMBPRINT = $ENDPOINT_REQUEST.ServicePoint.Certificate.GetCertHashString()

        return $SSL_THUMBPRINT -replace '(..(?!$))','$1:'
    }

    $VM = Get-VM -Name $VMName
    $VMMoref = $VM.ExtensionData.MoRef.Value

    if($webmksUrl) {
        $WebMKSTicket = $VM.ExtensionData.AcquireTicket("webmks")
        $VMHostName = $WebMKSTicket.host
        $Ticket = $WebMKSTicket.Ticket
        $URL = "wss://$VMHostName`:443/ticket/$Ticket"
    } elseif($vmrcUrl) {
        $VCName = $global:DefaultVIServer.Name
        $SessionMgr = Get-View $DefaultViserver.ExtensionData.Content.SessionManager
        $Ticket = $SessionMgr.AcquireCloneTicket()
        $URL = "vmrc://clone`:$Ticket@$VCName`:443/?moid=$VMMoref"
    } else {
        $VCInstasnceUUID = $global:DefaultVIServer.InstanceUuid
        $VCName = $global:DefaultVIServer.Name
        $SessionMgr = Get-View $DefaultViserver.ExtensionData.Content.SessionManager
        $Ticket = $SessionMgr.AcquireCloneTicket()
        $VCSSLThumbprint = Get-SSLThumbprint "https://$VCname"
        $URL = "https://$VCName`:9443/vsphere-client/webconsole.html?vmId=$VMMoref&vmName=$VMname&serverGuid=$VCInstasnceUUID&locale=en_US&host=$VCName`:443&sessionTicket=$Ticket&thumbprint=$VCSSLThumbprint”
    }
    $URL
}
﻿<#
.SYNOPSIS Remoting collecting esxcfg-info from an ESXi host using vCenter Server
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/06/using-the-vsphere-api-to-remotely-collect-esxi-esxcfg-info.html
.PARAMETER Vmhost
  ESXi host
.EXAMPLE
  PS> Get-VMHost -Name "esxi-1" | Get-Esxcfginfo
#>

Function Get-Esxcfginfo {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl[]]$VMHost
    )

    $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

    # URL to the ESXi esxcfg-info info
    $url = "https://" + $vmhost.Name + "/cgi-bin/esxcfg-info.cgi?xml"

    $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
    $spec.Method = "httpGet"
    $spec.Url = $url
    $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

    # Append the cookie generated from VC
    $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $cookie = New-Object System.Net.Cookie
    $cookie.Name = "vmware_cgi_ticket"
    $cookie.Value = $ticket.id
    $cookie.Domain = $vmhost.name
    $websession.Cookies.Add($cookie)

    # Retrieve file
    $result = Invoke-WebRequest -Uri $url -WebSession $websession -ContentType "application/xml"
    
    # cast output as an XML object
    return [ xml]$result.content
}

Connect-VIServer -Server 192.168.1.51 -User administrator@vghetto.local -password VMware1! | Out-Null

$xmlResult = Get-VMHost -Name "192.168.1.190" | Get-Esxcfginfo

# Extracting device-name, vendor-name & vendor-id as an example
foreach ($childnodes in $xmlResult.host.'hardware-info'.'pci-info'.'all-pci-devices'.'pci-device') {
   foreach ($childnode in $childnodes | select -ExpandProperty childnodes) {
    if($childnode.name -eq 'device-name') {
        $deviceName = $childnode.'#text'
    } elseif($childnode.name -eq 'vendor-name') {
        $vendorName = $childnode.'#text'
    } elseif($childnode.name -eq 'vendor-id') {
        $vendorId = $childnode.'#text'
    }
   }
   $deviceName
   $vendorName
   $vendorId
   Write-Host ""
}

Disconnect-VIServer * -Confirm:$false
﻿<#
.SYNOPSIS Remoting collecting ESXi configuration files using vCenter Server
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/06/using-the-vsphere-api-to-remotely-collect-esxi-configuration-files.html
.PARAMETER Vmhost
  ESXi host
.EXAMPLE
  PS> Get-VMHost -Name "esxi-1" | Get-Esxconf
#>

Function Get-Esxconf {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl[]]$VMHost
    )

    $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

    # URL to ESXi's esx.conf configuration file (can use any that show up under https://esxi_ip/host)
    $url = "https://192.168.1.190/host/esx.conf"

    # URL to the ESXi configuration file
    $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
    $spec.Method = "httpGet"
    $spec.Url = $url
    $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

    # Append the cookie generated from VC
    $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $cookie = New-Object System.Net.Cookie
    $cookie.Name = "vmware_cgi_ticket"
    $cookie.Value = $ticket.id
    $cookie.Domain = $vmhost.name
    $websession.Cookies.Add($cookie)

    # Retrieve file
    $result = Invoke-WebRequest -Uri $url -WebSession $websession
    return $result.content
}

Connect-VIServer -Server 192.168.1.51 -User administrator@vghetto.local -password VMware1! | Out-Null

$esxConf = Get-VMHost -Name "192.168.1.190" | Get-Esxconf

$esxConf

Disconnect-VIServer * -Confirm:$false﻿<#
.SYNOPSIS Using the vSphere API in vCenter Server to collect ESXTOP & vscsiStats metrics
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2017/02/using-the-vsphere-api-in-vcenter-server-to-collect-esxtop-vscsistats-metrics.html
.PARAMETER Vmhost
  ESXi host
.EXAMPLE
  PS> Get-VMHost -Name "esxi-1" | Get-EsxtopAPI
#>

Function Get-EsxtopAPI {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$VMHost
    )

    $serviceManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.serviceManager) -property "" -ErrorAction SilentlyContinue

    $locationString = "vmware.host." + $VMHost.Name
    $services = $serviceManager.QueryServiceList($null,$locationString)
    foreach ($service in $services) {
        if($service.serviceName -eq "Esxtop") {
            $serviceView = Get-View $services.Service -Property "entity"
            $serviceView.ExecuteSimpleCommand("CounterInfo")
            break
        }
    }
}

Connect-VIServer -Server 192.168.1.51 -User administrator@vsphere.local -password VMware1! | Out-Null

Get-VMHost -Name "192.168.1.50" | Get-EsxtopAPI

Disconnect-VIServer * -Confirm:$false# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script querying remote ESXi host without adding to vCenter Server
# Reference: http://www.virtuallyghetto.com/2016/07/remotely-query-an-esxi-host-without-adding-to-vcenter-server.html

Function Get-RemoteESXi {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [String]$hostname,
    [string]$username,
    [String]$password,
    [string]$port = 443
    )

    # Function to retrieve SSL Thumbprint of a host
    # https://gist.github.com/lamw/988e4599c0f88d9fc25c9f2af8b72c92
    Function Get-SSLThumbprint {
        param(
        [Parameter(
            Position=0,
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [Alias('FullName')]
        [String]$URL
        )

    add-type @"
            using System.Net;
            using System.Security.Cryptography.X509Certificates;
                public class IDontCarePolicy : ICertificatePolicy {
                public IDontCarePolicy() {}
                public bool CheckValidationResult(
                    ServicePoint sPoint, X509Certificate cert,
                    WebRequest wRequest, int certProb) {
                    return true;
                }
            }
"@
        [System.Net.ServicePointManager]::CertificatePolicy = new-object IDontCarePolicy

        # Need to connect using simple GET operation for this to work
        Invoke-RestMethod -Uri $URL -Method Get | Out-Null

        $ENDPOINT_REQUEST = [System.Net.Webrequest]::Create("$URL")
        $SSL_THUMBPRINT = $ENDPOINT_REQUEST.ServicePoint.Certificate.GetCertHashString()

        return $SSL_THUMBPRINT -replace '(..(?!$))','$1:'
    }

    # Host Connection Spec
    $spec = New-Object VMware.Vim.HostConnectSpec
    $spec.Force = $False
    $spec.HostName = $hostname
    $spec.UserName = $username
    $spec.Password = $password
    $spec.Port = $port
    # Retrieve the SSL Thumbprint from ESXi host
    $spec.SslThumbprint = Get-SSLThumbprint "https://$hostname"

    # Using first available Datacenter object to query remote ESXi host 
    return (Get-Datacenter)[0].ExtensionData.QueryConnectionInfoViaSpec($spec)
}

# vCenter Server credentials
$vc_server = "192.168.1.51"
$vc_username = "administrator@vghetto.local"
$vc_password = "VMware1!"

# Remote ESXi credentials to connect
$remote_esxi_hostname = "192.168.1.190"
$remote_esxi_username = "root"
$remote_esxi_password = "vmware123"

$server = Connect-VIServer -Server $vc_server -User $vc_username -Password $vc_password

$result = Get-RemoteESXi -hostname $remote_esxi_hostname -username $remote_esxi_username -password $remote_esxi_password

$result

Disconnect-VIServer $server -Confirm:$false
﻿<#
.SYNOPSIS  Query vCenter Server Database (VCDB) for its
           current usage of the Core, Alarm, Events & Stats table
.DESCRIPTION Script that performs SQL Query against a VCDB running either
             MSSQL & Oracle and collects current usage data for the
             following tables Core, Alarm, Events & Stats table. In
             Addition, if you wish to use the VCSA Migration Tool, the script
             can also calculate the estimated downtime required for either
             migration Option 1 or 2.
.NOTES  Author:    William Lam - @lamw
.NOTES  Site:      www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/09/how-to-check-the-size-of-your-config-stats-events-alarms-tasks-seat-data-in-the-vcdb.html
.PARAMETER dbType
  mssql or oracle
.PARAMETER connectionType
  local (mssql) and remote (mssql or oracle)
.PARAMETER dbServer
   VCDB Server
.PARAMETER dbPort
   VCDB Server Port
.PARAMETER dbUsername
   VCDB Username
.PARAMETER dbPassword
   VCDB Password
.PARAMETER dbInstance
   VCDB Instance Name
.PARAMETER estimate_migration_type
   option1 or option2 for those looking to calculate Windows VC to VCSA Migration (vSphere 6.0 U2m only)
.EXAMPLE
  Run the script locally on the Microsoft SQL Server hosting the vCenter Server Database
  Get-VCDBUsage -dbType mssql -connectionType local -dbServer sql.primp-industries.com
.EXAMPLE
  Run the script remotely on the Microsoft SQL Server hosting the vCenter Server Database
  Get-VCDBUsage -dbType mssql -connectionType local -dbServer sql.primp-industries.com -dbPort 1433 -dbInstance VCDB -dbUsername sa -dbPassword VMware1!
.EXAMPLE
  Run the script remotely on the Microsoft SQL Server hosting the vCenter Server Database & calculate VCSA migration downtime w/option1
  Get-VCDBUsage -dbType mssql -connectionType local -dbServer sql.primp-industries.com -dbPort 1433 -dbInstance VCDB -dbUsername sa -dbPassword VMware1! -migration_type option1
.EXAMPLE
  Run the script remotely to connect to Oracle Sever hosting the vCenter Server Database
  Get-VCDBUsage -dbType oracle -connectionType remote -dbServer oracle.primp-industries.com -dbPort 1521 -dbInstance VCDB -dbUsername vpxuser -dbPassword VMware1!
#>

function UpdateGitHubStats ([string] $csv_stats)
{
    #
    # github token test in psh
    #

    $encoded_token = "YmUxMzZlZWI4ZGI1ZTY3NmJjMGQ1ZmI1MDhjOTYzZGExZDEyNDkzZA=="
    $github_token = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($encoded_token))
    $github_repository = "https://api.github.com/repos/migrate2vcsa/stats/contents/vcsadb.csv?access_token=$github_token"

    $HttpRes = ""

    # Fetch the current file content/commit data (GET)
    try {
        $HttpRes = Invoke-RestMethod -Uri $github_repository -Method "GET" -ContentType "application/json"
    }
    catch {
        Write-Host -ForegroundColor Red "Error connecting to $github_repository"
        Write-Host -ForegroundColor Red $_.Exception.Message
    }


    # Decode base64 text
    $content = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($HttpRes.content))

    # Append any new stuff to the current text file
    $newcontent = $content + "$csv_stats`n"

    # Encode back to base64
    $encoded_content = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($newcontent))

    # Fetch commit sha
    $sha = $HttpRes.sha

    # Generate json response
    $json = @"
    {
        "sha": "$sha",
        "content": "$encoded_content",
        "message": "Updated file",
        "committer": {
            "name" : "vS0ciety",
            "email" : "migratetovcsa@gmail.com"
        }
    }
"@

    # Create the commit request (PUT)
    try {
        $HttpRes = Invoke-RestMethod -Uri $github_repository -Method "PUT" -Body $json -ContentType "application/json"
    }
    catch {
        Write-Host -ForegroundColor Red "Error connecting to $github_repository"
        Write-Host -ForegroundColor Red $_.Exception.Message
    }
}

Function Get-VCDBMigrationTime {
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][Double]$alarmData,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][Double]$coreData,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][Double]$eventData,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][Double]$statData,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$migration_type
    )

    # Sum up total size of the selected migration option
    switch($migration_type) {
        option1 {$VCDBSize=[math]::Round(($coreData + $alarmData),2);break}
        option2 {$VCDBSize=[math]::Round(($coreData + $alarmData + $eventData + $statData),2);break}
    }

    # Formulas extracted from excel spreadsheet from https://kb.vmware.com/kb/2146420
    $H5 = [math]::round(1.62*[math]::pow(2.5, [math]::Log($VCDBSize/75,2)) + (5.47-1.62)/75*$VCDBSize,2)
    $H7 = [math]::round(1.62*[math]::pow(2.5, [math]::Log($VCDBSize/75,2)) + (3.93-1.62)/75*$VCDBSize,2)
    $H6 = $H5 - $H7

    # Calculate timings
    $totalTimeHours = [math]::floor($H5)
    $totalTimeMinutes = [math]::round($H5 - $totalTimeHours,2)*60
    $exportTimeHours = [math]::floor($H6)
    $exportTimeMinutes = [math]::round($H6 - $exportTimeHours,2)*60
    $importTimeHours = [math]::floor($H7)
    $importtTimeminutes = [math]::round($H7 - $importTimeHours,2)*60

    # Return nice description string of selected migration option
    switch($migration_type) {
        option1 { $migrationDescription = "(Core + Alarm = " + $VCDBSize + " GB)";break}
        option2 { $migrationDescription = "(Core + Alarm + Event + Stat = " + $VCDBSize + " GB)";break}
    }

    Write-Host -ForegroundColor Yellow "`nvCenter Server Migration Estimates for"$migration_type $migrationDescription"`n"
    Write-Host "Total  Time :" $totalTimeHours "Hours" $totalTimeMinutes "Minutes"
    Write-Host "Export Time :" $exportTimeHours "Hours" $exportTimeMinutes "Minutes"
    Write-Host "Import Time :" $importTimeHours "Hours" $importtTimeminutes "Minutes`n"
}

Function Get-VCDBUsage {
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$dbType,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$connectionType,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$dbServer,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][int]$dbPort,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$dbUsername,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$dbPassword,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$dbInstance,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$estimate_migration_type
    )

    $mssql_vcdb_usage_query = @"
use $dbInstance;
select tabletype, sum(rowcounts) as rowcounts,
       sum(spc.usedspaceKB)/1024.0 as usedspaceMB
from
    (select
        s.name as schemaname,
        t.name as tablename,
        p.rows as rowcounts,
        sum(a.used_pages) * 8 as usedspaceKB,
        case
            when t.name like 'VPX_ALARM%' then 'Alarm'
            when t.name like 'VPX_EVENT%' then 'ET'
            when t.name like 'VPX_TASK%' then 'ET'
            when t.name like 'VPX_HIST_STAT%' then 'Stats'
            when t.name = 'VPX_STAT_COUNTER' then 'Stats'
            when t.name = 'VPX_TOPN%' then 'Stats'
            else 'Core'
        end as tabletype
    from
        sys.tables t
    inner join
        sys.schemas s on s.schema_id = t.schema_id
    inner join
        sys.indexes i on t.object_id = i.object_id
    inner join
        sys.partitions p on i.object_id = p.object_id and i.index_id = p.index_id
    inner join
        sys.allocation_units a on p.partition_id = a.container_id
    where
        t.name not like 'dt%'
        and t.is_ms_shipped = 0
        and i.object_id >= 255
    group by
        t.name, s.name, p.rows) as spc

group by tabletype;
"@

    $oracle_vcdb_usage_query = @"
SELECT tabletype,
       SUM(CASE rn WHEN 1 THEN row_cnt ELSE 0 END) AS rowcount,
       ROUND(SUM(sized)/(1024*1024)) usedspaceMB
 FROM (
      SELECT
            CASE
               WHEN segment_name LIKE '%ALARM%' THEN 'Alarm'
               WHEN segment_name LIKE '%EVENT%' THEN 'ET'
               WHEN segment_name LIKE '%TASK%' THEN 'ET'
               WHEN segment_name LIKE '%HIST_STAT%' THEN 'Stats'
               WHEN segment_name LIKE 'VPX_TOPN%' THEN 'Stats'
               ELSE 'Core'
            END AS tabletype,
            row_cnt,
            sized ,
            ROW_NUMBER () OVER (PARTITION BY table_name ORDER BY segment_name) AS rn
       FROM (
            SELECT
                  t.table_name, t.table_name segment_name,
                  t.NUM_ROWS AS row_cnt, s.bytes AS sized
             FROM user_segments s
             JOIN user_tables t ON s.segment_name = t.table_name AND s.segment_type = 'TABLE'
             UNION ALL
            SELECT
                  ti.table_name,i.index_name, ti.NUM_ROWS,s.bytes
             FROM user_segments s
             JOIN user_indexes i ON s.segment_name = i.index_name AND s.segment_type = 'INDEX'
             JOIN user_tables ti ON i.table_name = ti.table_name) table_index ) type_cnt_size
GROUP BY tabletype
"@

    $oracle_odbc_dll_path = "C:\Oracle\odp.net\managed\common\Oracle.ManagedDataAccess.dll"

    Function Run-VCDBMSSQLQuery {

        Function Run-LocalMSSQLQuery {
            Write-Host -ForegroundColor Green "`nRunning Local MSSQL VCDB Usage Query"

            # Check whether Invoke-Sqlcmd cmdlet exists
            if( (Get-Command "Invoke-Sqlcmd" -errorAction SilentlyContinue -CommandType Cmdlet) -eq $null) {
               Write-Host -ForegroundColor Red "Invoke-Sqlcmd cmdlet does not exists on this system, you will need to install SQL Tools or run remotely with DB credentials`n"
               exit
            }

            try {
                $results = Invoke-Sqlcmd -Query $mssql_vcdb_usage_query -ServerInstance $dbServer
            }
            catch { Write-Host -ForegroundColor Red "Unable to connect to the SQL Server. Its possible the SQL Server is not configured to allow remote connections`n"; exit }

            foreach ($result in $results) {
                switch($result.tabletype) {
                    Alarm { $alarm_usage=$result.usedspaceMB; $alarm_rows=$result.rowcounts; break}
                    Core { $core_usage=$result.usedspaceMB; $core_rows=$result.rowcounts; break}
                    ET { $event_usage=$result.usedspaceMB; $event_rows=$result.rowcounts; break}
                    Stats { $stat_usage=$result.usedspaceMB; $stat_rows=$result.rowcounts; break}
                }
            }

            return ($alarm_usage,$core_usage,$event_usage,$stat_usage,$alarm_rows,$core_rows,$event_rows,$stat_rows)
        }

        Function Run-RemoteMSSQLQuery {
            if($dbServer -eq $null -eq $null -or $dbInstance -eq $null -or $dbUsername -eq $null -or $dbPassword -eq $null) {
                Write-host -ForegroundColor Red "One or more parameters is missing for the remote MSSQL Query option`n"
                exit
            }

            Write-Host -ForegroundColor Green "`nRunning Remote MSSQL VCDB Usage Query"

            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            if($dbPort -eq 0) {
              $SqlConnection.ConnectionString = "Server = $dbServer; Database = $dbInstance; User ID = $dbUsername; Password = $dbPassword;"
            } else {
              $SqlConnection.ConnectionString = "Server = $dbServer, $dbPort; Database = $dbInstance; User ID = $dbUsername; Password = $dbPassword;"
            }

            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandText = $mssql_vcdb_usage_query
            $SqlCmd.Connection = $SqlConnection

            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
            $SqlAdapter.SelectCommand = $SqlCmd

            try {
                $DataSet = New-Object System.Data.DataSet
                $numRecords = $SqlAdapter.Fill($DataSet)
            } catch { Write-Host -ForegroundColor Red "Unable to connect and execute query on the SQL Server. Its possible the SQL Server is not configured to allow remote connections`n"; exit }

            $SqlConnection.Close()

            foreach ($result in $DataSet.Tables[0]) {
                switch($result.tabletype) {
                    Alarm { $alarm_usage=$result.usedspaceMB; $alarm_rows=$result.rowcounts; break}
                    Core { $core_usage=$result.usedspaceMB; $core_rows=$result.rowcounts; break}
                    ET { $event_usage=$result.usedspaceMB; $event_rows=$result.rowcounts; break}
                    Stats { $stat_usage=$result.usedspaceMB; $stat_rows=$result.rowcounts; break}
                }
            }

            return ($alarm_usage,$core_usage,$event_usage,$stat_usage,$alarm_rows,$core_rows,$event_rows,$stat_rows)
        }

        switch($connectionType) {
            local { Run-LocalMSSQLQuery;break}
            remote { Run-RemoteMSSQLQuery;break}
        }
    }

    Function Run-VCDBOracleQuery {
        if($dbServer -eq $null -or $dbPort -eq $null -or $dbInstance -eq $null -or $dbUsername -eq $null -or $dbPassword -eq $null) {
            Write-host -ForegroundColor Red "One or more parameters is missing for the remote Oracle Query option`n"
            exit
        }

        Write-Host -ForegroundColor Green "`nRunning Remote Oracle VCDB Usage Query"

        if(Test-Path "$oracle_odbc_dll_path") {
            Add-Type -Path "$oracle_odbc_dll_path"

            $connectionString="Data Source = (DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$dbserver)(PORT=$dbport))(CONNECT_DATA=(SERVICE_NAME=$dbinstance)));User Id=$dbusername;Password=$dbpassword;"

            $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionString)

            try {
                $connection.open()
                $command=$connection.CreateCommand()
                $command.CommandText=$query
                $reader=$command.ExecuteReader()
            } catch { Write-Host -ForegroundColor Red "Unable to connect to Oracle DB. Ensure your connection info is correct and system allows for remote connections`n"; exit }

            while ($reader.Read()) {
                $table_name = $reader.getValue(0)
                $table_rows = $reader.getValue(1)
                $table_size = $reader.getValue(2)
                switch($table_name) {
                    Alarm { $alarm_usage=$table_size; $alarm_rows=$table_rows; break}
                    Core { $core_usage=$table_size; $core_rows=$table_rows; break}
                    ET { $event_usage=$table_size; $event_rows=$table_rows; break}
                    Stats { $stat_usage=$table_size; $stat_rows=$table_rows; break}
                }
            }
            $connection.Close()

        } else {
            Write-Host -ForegroundColor Red "Unable to find Oracle ODBC DLL which has been defined in the following path: $oracle_odbc_dll_path"
            exit
        }
        return ($alarm_usage,$core_usage,$event_usage,$stat_usage,$alarm_rows,$core_rows,$event_rows,$stat_rows)
    }

    # Run selected DB query and return 4 expected tables from VCDB
    ($alarmData,$coreData,$eventData,$statData,$alarm_rows,$core_rows,$event_rows,$stat_rows) = (0,0,0,0,0,0,0,0)
    switch($dbType) {
        mssql { ($alarmData,$coreData,$eventData,$statData,$alarm_rows,$core_rows,$event_rows,$stat_rows) = Run-VCDBMSSQLQuery; break }
        oracle { ($alarmData,$coreData,$eventData,$statData,$alarm_rows,$core_rows,$event_rows,$stat_rows) = Run-VCDBOracleQuery; break }
        default { Write-Host "mssql or oracle are the only valid dbType options" }
    }

    # Convert data from MB to GB
    $coreData = [math]::Round(($coreData*1024*1024)/1GB,2)
    $alarmData = [math]::Round(($alarmData*1024*1024)/1GB,2)
    $eventData = [math]::Round(($eventData*1024*1024)/1GB,2)
    $statData = [math]::Round(($statData*1024*1024)/1GB,2)

    Write-Host "`nCore Data :"$coreData" GB (rows:"$core_rows")"
    Write-Host "Alarm Data:"$alarmData" GB (rows:"$alarm_rows")"
    Write-Host "Event Data:"$eventData" GB (rows:"$event_rows")"
    Write-Host "Stat Data :"$statData" GB (rows:"$stat_rows")"

    # If user wants VCSA migration estimates, run the additional calculation
    if($estimate_migration_type -eq "option1" -or $estimate_migration_type -eq "option2") {
        Get-VCDBMigrationTime -alarmData $alarmData -coreData $coreData -eventData $eventData -statData $statData -migration_type $estimate_migration_type
    }

    Write-Host -ForegroundColor Magenta `
    "`nWould you like to be able to compare your VCDB Stats with others? `
If so, when prompted, type yes and only the size & # of rows will `
be sent to https://github.com/migrate2vcsa for further processing`n"
    $answer = Read-Host -Prompt "Do you accept (Y or N)"
    if($answer -eq "Y" -or $answer -eq "y") {
        UpdateGitHubStats("$dbType,$alarmData,$coreData,$eventData,$statData,$alarm_rows,$core_rows,$event_rows,$stat_rows")
    }
}

# Please replace variables your own VCDB details
$dbType = "mssql"
$connectionType = "remote"
$dbServer = "sql.primp-industries.com"
$dbPort = "1433"
$dbInstance = "VCDB"
$dbUsername = "sa"
$dbPassword = "VMware1!"

Get-VCDBUsage -connectionType $connectionType -dbType $dbType -dbServer $dbServer -dbPort $dbPort -dbInstance $dbInstance -dbUsername $dbUsername -dbPassword $dbPassword
<#
.SYNOPSIS  Query vCenter Server Database (VCDB) for its
           current usage of the Core, Alarm, Events & Stats table
.DESCRIPTION Script that performs SQL Query against a VCDB running
            on a vPostgres DB
.NOTES  Author:    William Lam - @lamw
.NOTES  Site:      www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/10/how-to-check-the-size-of-your-config-seat-data-in-the-vcdb-in-vpostgres.html
.PARAMETER dbServer
   VCDB Server
.PARAMETER dbName
   VCDB Instance Name
.PARAMETER dbUsername
   VCDB Username
.PARAMETER dbPassword
   VCDB Password
.EXAMPLE
  Get-VCDBUsagevPostgres -dbServer vcenter60-1.primp-industries.com -dbName VCDB -dbUser vc -dbPass "VMware1!"
#>

Function Get-VCDBUsagevPostgres{
    param(
          [string]$dbServer,
          [string]$dbName,
          [string]$dbUser,
          [string]$dbPass
         )

         $query = @"
         SELECT   tabletype,
         sum(reltuples) as rowcount,
         ceil(sum(pg_total_relation_size(oid)) / (1024*1024)) as usedspaceMB
FROM  (
      SELECT   CASE
                  WHEN c.relname LIKE 'vpx_alarm%' THEN 'Alarm'
                  WHEN c.relname LIKE 'vpx_event%' THEN 'ET'
                  WHEN c.relname LIKE 'vpx_task%' THEN 'ET'
                  WHEN c.relname LIKE 'vpx_hist_stat%' THEN 'Stats'
                  WHEN c.relname LIKE 'vpx_topn%' THEN 'Stats'
                  ELSE 'Core'
               END AS tabletype,
               c.reltuples, c.oid
        FROM pg_class C
        LEFT JOIN pg_namespace N
          ON N.oid = C.relnamespace
       WHERE nspname IN ('vc', 'vpx') and relkind in ('r', 't')) t
GROUP BY tabletype;
"@

    $conn = New-Object System.Data.Odbc.OdbcConnection
    $conn.ConnectionString = "Driver={PostgreSQL UNICODE(x64)};Server=$dbServer;Port=5432;Database=$dbName;Uid=$dbUser;Pwd=$dbPass;ReadOnly=1"
    $conn.open()
    $cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
    $ds = New-Object system.Data.DataSet
    (New-Object system.Data.odbc.odbcDataAdapter($cmd)).fill($ds) | out-null
    $conn.close()
    $ds.Tables[0]
}

# Please replace variables your own VCDB details
$dbServer = "vcenter60-1.primp-industries.com"
$dbInstance = "VCDB"
$dbUsername = "vc"
$dbPassword = "ezbo3wrMqkJB6{7t"

Get-VCDBUsagevPostgres $dbServer $dbInstance $dbUsername $dbPassword
<#
.SYNOPSIS  Returns configuration changes for a VM
.DESCRIPTION The function will return the list of configuration changes
    for a given Virtual Machine
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Comment: Modified example from Lucd's blog post http://www.lucd.info/2009/12/18/events-part-3-auditing-vm-device-changes/
.PARAMETER Vm
  Virtual Machine object to query configuration changes
.PARAMETER Hour
  The number of hours to to search for configuration changes, default 8hrs
.EXAMPLE
  PS> Get-VMConfigChanges -vm $VM
.EXAMPLE
  PS> Get-VMConfigChanges -vm $VM -hours 8
#>

Function Get-VMConfigChanges {
    param($vm, $hours=8)

    # Modified code from http://powershell.com/cs/blogs/tips/archive/2012/11/28/removing-empty-object-properties.aspx
    Function prettyPrintEventObject($vmChangeSpec,$task) {
    	$hashtable = $vmChangeSpec |
	    Get-Member -MemberType *Property |
    	Select-Object -ExpandProperty Name |
	    Sort-Object |
    	ForEach-Object -Begin {
  	    	[System.Collections.Specialized.OrderedDictionary]$rv=@{}
  	    	} -process {
  		    if ($vmChangeSpec.$_ -ne $null) {
    		    $rv.$_ = $vmChangeSpec.$_
      		}
	    } -end {$rv}

    	# Add in additional info to the return object (Thanks to Luc's Code)
    	$hashtable.Add('VMName',$task.EntityName)
	    $hashtable.Add('Start', $task.StartTime)
    	$hashtable.Add('End', $task.CompleteTime)
	    $hashtable.Add('State', $task.State)
    	$hashtable.Add('User', $task.Reason.UserName)
      $hashtable.Add('ChainID', $task.EventChainId)

    	# Device Change
	    $vmChangeSpec.DeviceChange | % {
		    if($_.Device -ne $null) {
          $hashtable.Add('Device', $_.Device.GetType().Name)
			    $hashtable.Add('Operation', $_.Operation)
        }
	    }
	    $newVMChangeSpec = New-Object PSObject
	    $newVMChangeSpec | Add-Member ($hashtable) -ErrorAction SilentlyContinue
	    return $newVMChangeSpec
    }

    # Modified code from Luc Dekens http://www.lucd.info/2009/12/18/events-part-3-auditing-vm-device-changes/
    $tasknumber = 999 # Windowsize for task collector
    $eventnumber = 100 # Windowsize for event collector

    $report = @()
    $taskMgr = Get-View TaskManager
    $eventMgr = Get-View eventManager

    $tFilter = New-Object VMware.Vim.TaskFilterSpec
    $tFilter.Time = New-Object VMware.Vim.TaskFilterSpecByTime
    $tFilter.Time.beginTime = (Get-Date).AddHours(-$hours)
    $tFilter.Time.timeType = "startedTime"
    $tFilter.Entity = New-Object VMware.Vim.TaskFilterSpecByEntity
    $tFilter.Entity.Entity = $vm.ExtensionData.MoRef
    $tFilter.Entity.Recursion = New-Object VMware.Vim.TaskFilterSpecRecursionOption
    $tFilter.Entity.Recursion = "self"

    $tCollector = Get-View ($taskMgr.CreateCollectorForTasks($tFilter))

    $dummy = $tCollector.RewindCollector
    $tasks = $tCollector.ReadNextTasks($tasknumber)

    while($tasks){
      $tasks | where {$_.Name -eq "ReconfigVM_Task"} | % {
        $task = $_
        $eFilter = New-Object VMware.Vim.EventFilterSpec
        $eFilter.eventChainId = $task.EventChainId

        $eCollector = Get-View ($eventMgr.CreateCollectorForEvents($eFilter))
        $events = $eCollector.ReadNextEvents($eventnumber)
        while($events){
          $events | % {
            $event = $_
            switch($event.GetType().Name){
              "VmReconfiguredEvent" {
                $event.ConfigSpec | % {
				    $report += prettyPrintEventObject $_ $task
                }
              }
              Default {}
            }
          }
          $events = $eCollector.ReadNextEvents($eventnumber)
        }
        $ecollection = $eCollector.ReadNextEvents($eventnumber)
	    # By default 32 event collectors are allowed. Destroy this event collector.
        $eCollector.DestroyCollector()
      }
      $tasks = $tCollector.ReadNextTasks($tasknumber)
    }

    # By default 32 task collectors are allowed. Destroy this task collector.
    $tCollector.DestroyCollector()

    $report
}

$vcserver = "192.168.1.150"
$vcusername = "administrator@vghetto.local"
$vcpassword = "VMware1!"

Connect-VIServer -Server $vcserver -User $vcusername -Password $vcpassword

$vm = Get-VM "Test-VM"

Get-VMConfigChanges -vm $vm -hours 1

Disconnect-VIServer -Server $vcserver -Confirm:$false
<#
.SYNOPSIS  Returns configuration changes for a VM using vCenter Server Alarm
.DESCRIPTION The function will return the list of configuration changes
    for a given Virtual Machine trigged by vCenter Server Alarm based on
    VmReconfigureEvent
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Comment: Modified example from Lucd's blog post http://www.lucd.info/2009/12/18/events-part-3-auditing-vm-device-changes/
.PARAMETER Moref
  Virtual Machine MoRef ID that generated vCenter Server Alarm
.PARAMETER EventId
  The ID correlating to the ReconfigVM operation
.EXAMPLE
  PS> Get-VMConfigChanges -moref vm-125 -eventId 8389
#>

Function Get-VMConfigChangesFromAlarm {
    param($moref, $eventId)

    # Construct VM object from MoRef ID
    $vm = New-Object VMware.Vim.ManagedObjectReference
    $vm.Type = "VirtualMachine"
    $vm.Value = $moref

    # Modified code from http://powershell.com/cs/blogs/tips/archive/2012/11/28/removing-empty-object-properties.aspx
    Function prettyPrintEventObject($vmChangeSpec,$task) {
    	$hashtable = $vmChangeSpec |
	    Get-Member -MemberType *Property |
    	Select-Object -ExpandProperty Name |
	    Sort-Object |
    	ForEach-Object -Begin {
  	    	[System.Collections.Specialized.OrderedDictionary]$rv=@{}
  	    	} -process {
  		    if ($vmChangeSpec.$_ -ne $null) {
    		    $rv.$_ = $vmChangeSpec.$_
      		}
	    } -end {$rv}

    	# Add in additional info to the return object (Thanks to Luc's Code)
    	$hashtable.Add('VMName',$task.EntityName)
	    $hashtable.Add('Start', $task.StartTime)
    	$hashtable.Add('End', $task.CompleteTime)
	    $hashtable.Add('State', $task.State)
    	$hashtable.Add('User', $task.Reason.UserName)

    	# Device Change
	    $vmChangeSpec.DeviceChange | % {
		    if($_.Device -ne $null) {
		        $hashtable.Add('Device', $_.Device.GetType().Name)
			    $hashtable.Add('Operation', $_.Operation)
            }
	    }
	    $newVMChangeSpec = New-Object PSObject
	    $newVMChangeSpec | Add-Member ($hashtable) -ErrorAction SilentlyContinue
	    return $newVMChangeSpec
    }

    # Modified code from Luc Dekens http://www.lucd.info/2009/12/18/events-part-3-auditing-vm-device-changes/
    $tasknumber = 999 # Windowsize for task collector
    $eventnumber = 100 # Windowsize for event collector

    $report = @()
    $taskMgr = Get-View TaskManager
    $eventMgr = Get-View eventManager

    $tFilter = New-Object VMware.Vim.TaskFilterSpec
    # Need to take eventId substract 1 to get real event
    $tFilter.eventChainId = ([int]$eventId - 1)
    $tFilter.Entity = New-Object VMware.Vim.TaskFilterSpecByEntity
    $tFilter.Entity.Entity = $vm
    $tFilter.Entity.Recursion = New-Object VMware.Vim.TaskFilterSpecRecursionOption
    $tFilter.Entity.Recursion = "self"

    $tCollector = Get-View ($taskMgr.CreateCollectorForTasks($tFilter))

    $dummy = $tCollector.RewindCollector
    $tasks = $tCollector.ReadNextTasks($tasknumber)

    while($tasks){
      $tasks | where {$_.Name -eq "ReconfigVM_Task"} | % {
        $task = $_
        $eFilter = New-Object VMware.Vim.EventFilterSpec
        $eFilter.eventChainId = $task.EventChainId

        $eCollector = Get-View ($eventMgr.CreateCollectorForEvents($eFilter))
        $events = $eCollector.ReadNextEvents($eventnumber)
        while($events){
          $events | % {
            $event = $_
            switch($event.GetType().Name){
              "VmReconfiguredEvent" {
                $event.ConfigSpec | % {
				    $report += prettyPrintEventObject $_ $task
                }
              }
              Default {}
            }
          }
          $events = $eCollector.ReadNextEvents($eventnumber)
        }
        $ecollection = $eCollector.ReadNextEvents($eventnumber)
	    # By default 32 event collectors are allowed. Destroy this event collector.
        $eCollector.DestroyCollector()
      }
      $tasks = $tCollector.ReadNextTasks($tasknumber)
    }

    # By default 32 task collectors are allowed. Destroy this task collector.
    $tCollector.DestroyCollector()

    $report | Out-File -filepath C:\Users\primp\Desktop\alarm.txt -Append
}

$vcserver = "172.30.0.112"
$vcusername = "administrator@vghetto.local"
$vcpassword = "VMware1!"

Connect-VIServer -Server $vcserver -User $vcusername -Password $vcpassword

# Parse vCenter Server Alarm environmental variables
$eventid_from_alarm = $env:VMWARE_ALARM_TRIGGERINGSUMMARY
$moref_from_alarm = $env:VMWARE_ALARM_TARGET_ID

# regex for string within paren http://powershell.com/cs/forums/p/7360/11988.aspx
$regex = [regex]"\((.*)\)"
$string = [regex]::match($eventid_from_alarm, $regex).Groups[1]
$eventid = $string.value

Get-VMConfigChangesFromAlarm -moref $moref_from_alarm -eventId $eventid

Disconnect-VIServer -Server $vcserver -Confirm:$false
﻿# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Retrieves the VM memory overhead for given VM
# Reference: http://www.virtuallyghetto.com/2015/12/easily-retrieve-vm-memory-overhead-using-the-vsphere-6-0-api.html

<#
.SYNOPSIS  Returns VM Ovehead a VM
.DESCRIPTION The function will return VM memory overhead
    for a given Virtual Machine
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.PARAMETER Vm
  Virtual Machine object to query VM memory overhead
.EXAMPLE
  PS> Get-VM "vcenter60-2" | Get-VMMemOverhead
#>

Function Get-VMMemOverhead {
    param(  
    [Parameter(
        Position=0, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [Alias('FullName')]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl[]]$VM
    ) 

    process {
        # Retrieve VM & ESXi MoRef
        $vmMoref = $VM.ExtensionData.MoRef
        $vmHostMoref = $VM.ExtensionData.Runtime.Host

        # Retrieve Overhead Memory Manager
        $overheadMgr = Get-View ($global:DefaultVIServer.ExtensionData.Content.OverheadMemoryManager)

        # Get VM Memory overhead
        $overhead = $overheadMgr.LookupVmOverheadMemory($vmMoref,$vmHostMoref)
        Write-Host $VM.Name "has overhead of" ([math]::Round($overhead/1MB,2)).ToString() "MB memory`n"
    }
}<#
.SYNOPSIS  Retrieve the VSAN Policy for a given VM(s) which includes filtering
    of VMs that do not contain a policy (None) or policies in which contains
    Thick Provisioning (e.g Object Space Reservation set to 100)
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.PARAMETER Vm
  Virtual Machine(s) object to query for VSAN VM Storage Policies
.EXAMPLE
  PS> Get-VM * | Get-VSANPolicy -datastore "vsanDatastore"
  PS> Get-VM * | Get-VSANPolicy -datastore "vsanDatastore" -nopolicy $false -thick $true -details $true
#>

Function Get-VSANPolicy {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl[]]$vms,
    [String]$details=$false,
    [String]$datastore,
    [String]$nopolicy=$false,
    [String]$thick=$false
    )

    process {
        foreach ($vm in $vms) {
            # Extract the VSAN UUID for VM Home
            $vm_dir,$vm_vmx = ($vm.ExtensionData.Config.Files.vmPathName).split('/').replace('[','').replace(']','')
            $vmdatastore,$vmhome_obj_uuid = ($vm_dir).split(' ')

            # Process only if we have a match on the specified datastore
            if($vmdatastore -eq $datastore) {
                $cmmds_queries = @()
                $disks_to_uuid_mapping = @{}
                $disks_to_uuid_mapping[$vmhome_obj_uuid] = "VM Home"

                # Create query object for VM home
                $vmhome_query = New-Object VMware.vim.HostVsanInternalSystemCmmdsQuery
                $vmhome_query.Type = "POLICY"
                $vmhome_query.Uuid = $vmhome_obj_uuid

                # Add the VM Home query object to overall cmmds query spec
                $cmmds_queries += $vmhome_query

                # Go through all VMDKs & build query object for each disk
                $devices = $vm.ExtensionData.Config.Hardware.Device
                foreach ($device in $devices) {
                    if($device -is [VMware.Vim.VirtualDisk]) {
                        if($device.backing.backingObjectId) {
                            $disks_to_uuid_mapping[$device.backing.backingObjectId] = $device.deviceInfo.label
                            $disk_query = New-Object VMware.vim.HostVsanInternalSystemCmmdsQuery
                            $disk_query.Type = "POLICY"
                            $disk_query.Uuid = $device.backing.backingObjectId
                            $cmmds_queries += $disk_query
                        }
                    }
                }

                # Access VSAN Internal System to issue the Cmmds query
                $vsanIntSys = Get-View ((Get-View $vm.ExtensionData.Runtime.Host -Property Name, ConfigManager.vsanInternalSystem).ConfigManager.vsanInternalSystem)
                $results = $vsanIntSys.QueryCmmds($cmmds_queries)

                $printed = @{}
                $json = $results | ConvertFrom-Json
                foreach ($j in $json.result) {
                    $storagepolicy_id = $j.content.spbmProfileId

                    # If there's no spbmProfileID, it means there's
                    # no VSAN VM Storage Policy assigned
                    # possibly deployed from vSphere C# Client
                    if($storagepolicy_id -eq $null -and $nopolicy -eq $true) {
                        $object_type = $disks_to_uuid_mapping[$j.uuid]
                        $policy = $j.content

                        # quick/dirty way to only print VM name once
                        if($printed[$vm.name] -eq $null -and $thick -eq $false) {
                            $printed[$vm.name] = "1"
                            Write-Host "`n"$vm.Name
                        }

                        if($details -eq $true -and $thick -eq $false) {
                           Write-Host "$object_type `t` $policy"
                        } elseIf($details -eq $false -and $thick -eq $false) {
                           Write-Host "$object_type `t` None"
                        } else {
                            # Ignore VM Home which will always be thick provisioned
                            if($object_type -ne "VM Home") {
                                if($policy.proportionalCapacity -eq 100) {
                                    Write-Host "`n"$vm.Name
                                    if($details -eq $true) {
                                        Write-Host "$object_type `t` $policy"
                                    } else {
                                        Write-Host "$object_type"
                                    }
                                }
                            }
                        }
                    } elseIf($storagepolicy_id -ne $null -and $nopolicy -eq $false) {
                        $object_type = $disks_to_uuid_mapping[$j.uuid]
                        $policy = $j.content

                        # quick/dirty way to only print VM name once
                        if($printed[$vm.name] -eq $null -and $thick -eq $false) {
                            $printed[$vm.name] = "1"
                            Write-Host "`n"$vm.Name
                        }

                        # Convert the VM Storage Policy ID to human readable name
                        $vsan_policy_name = Get-SpbmStoragePolicy -Id $storagepolicy_id

                        if($details -eq $true -and $thick -eq $false) {
                            Write-Host "$object_type `t` $vsan_policy_name `t` $policy"
                        } elseIf($details -eq $false -and $thick -eq $false) {
                            if($vsan_policy_name -eq $null) {
                                Write-Host "$object_type `t` None"
                            } else {
                                Write-Host "$object_type `t` $vsan_policy_name"
                            }
                        } else {
                            # Ignore VM Home which will always be thick provisioned
                            if($object_type -ne "VM Home") {
                                if($policy.proportionalCapacity -eq 100) {
                                    if($printed[$vm.name] -eq $null) {
                                        $printed[$vm.name] = "1"
                                        Write-Host "`n"$vm.Name
                                    }
                                    if($details -eq $true) {
                                        Write-Host "$object_type `t` $policy"
                                    } else {
                                        Write-Host "$object_type"
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

Connect-VIServer -Server 192.168.1.51 -User administrator@vghetto.local -password VMware1! | Out-Null

Get-VM "Photon-Deployed-From-WebClient*" | Get-VSANPolicy -datastore "vsanDatastore" -thick $true -details $true

Disconnect-VIServer * -Confirm:$false
﻿<#
.SYNOPSIS Using the vSphere API in vCenter Server to collect ESXTOP & vscsiStats metrics
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2017/02/using-the-vsphere-api-in-vcenter-server-to-collect-esxtop-vscsistats-metrics.html
.PARAMETER Vmhost
  ESXi host
.EXAMPLE
  PS> Get-VMHost -Name "esxi-1" | Get-VscsiStats
#>

Function Get-VscsiStats {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$VMHost
    )

    $serviceManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.serviceManager) -property "" -ErrorAction SilentlyContinue

    $locationString = "vmware.host." + $VMHost.Name
    $services = $serviceManager.QueryServiceList($null,$locationString)
    foreach ($service in $services) {
        if($service.serviceName -eq "VscsiStats") {
            $serviceView = Get-View $services.Service -Property "entity"
            $serviceView.ExecuteSimpleCommand("FetchAllHistograms")
            break
        }
    }
}

Connect-VIServer -Server 192.168.1.51 -User administrator@vsphere.local -password VMware1! | Out-Null

Get-VMHost -Name "192.168.1.50" | Get-VscsiStats

Disconnect-VIServer * -Confirm:$falseFunction New-GlobalPermission {
<#
    .DESCRIPTION Script to add/remove vSphere Global Permission
    .NOTES  Author:  William Lam
    .NOTES  Site:    www.virtuallyghetto.com
    .NOTES  Reference: http://www.virtuallyghetto.com/2017/02/automating-vsphere-global-permissions-with-powercli.html
    .PARAMETER vc_server
        vCenter Server Hostname or IP Address
    .PARAMETER vc_username
        VC Username
    .PARAMETER vc_password
        VC Password
    .PARAMETER vc_user
        Name of the user to remove global permission on
    .PARAMETER vc_role_id
        The ID of the vSphere Role (retrieved from Get-VIRole)
    .PARAMETER propagate
        Whether or not to propgate the permission assignment (true/false)
#>
    New-GlobalPermission -vc_server "192.168.1.51" -vc_username "administrator@vsphere.local" -vc_password "VMware1!" -vc_user "VGHETTO\lamw" -vc_role_id "-1" -propagate "true"
    param(
        [Parameter(Mandatory=$true)][string]$vc_server,
        [Parameter(Mandatory=$true)][String]$vc_username,
        [Parameter(Mandatory=$true)][String]$vc_password,
        [Parameter(Mandatory=$true)][String]$vc_user,
        [Parameter(Mandatory=$true)][String]$vc_role_id,
        [Parameter(Mandatory=$true)][String]$propagate
    )

    $secpasswd = ConvertTo-SecureString $vc_password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($vc_username, $secpasswd)

    # vSphere MOB URL to private enableMethods
    $mob_url = "https://$vc_server/invsvc/mob3/?moid=authorizationService&method=AuthorizationService.AddGlobalAccessControlList"

# Ingore SSL Warnings
add-type -TypeDefinition  @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

    # Initial login to vSphere MOB using GET and store session using $vmware variable
    $results = Invoke-WebRequest -Uri $mob_url -SessionVariable vmware -Credential $credential -Method GET

    # Extract hidden vmware-session-nonce which must be included in future requests to prevent CSRF error
    # Credit to https://blog.netnerds.net/2013/07/use-powershell-to-keep-a-cookiejar-and-post-to-a-web-form/ for parsing vmware-session-nonce via Powershell
    if($results.StatusCode -eq 200) {
        $null = $results -match 'name="vmware-session-nonce" type="hidden" value="?([^\s^"]+)"'
        $sessionnonce = $matches[1]
    } else {
        Write-host "Failed to login to vSphere MOB"
        exit 1
    }

    # Escape username
    $vc_user_escaped = [uri]::EscapeUriString($vc_user)

    # The POST data payload must include the vmware-session-nonce variable + URL-encoded
    $body = @"
vmware-session-nonce=$sessionnonce&permissions=%3Cpermissions%3E%0D%0A+++%3Cprincipal%3E%0D%0A++++++%3Cname%3E$vc_user_escaped%3C%2Fname%3E%0D%0A++++++%3Cgroup%3Efalse%3C%2Fgroup%3E%0D%0A+++%3C%2Fprincipal%3E%0D%0A+++%3Croles%3E$vc_role_id%3C%2Froles%3E%0D%0A+++%3Cpropagate%3E$propagate%3C%2Fpropagate%3E%0D%0A%3C%2Fpermissions%3E
"@
    # Second request using a POST and specifying our session from initial login + body request
    Write-Host "Adding Global Permission for $vc_user ..."
    $results = Invoke-WebRequest -Uri $mob_url -WebSession $vmware -Method POST -Body $body

    # Logout out of vSphere MOB
    $mob_logout_url = "https://$vc_server/invsvc/mob3/logout"
    $results = Invoke-WebRequest -Uri $mob_logout_url -WebSession $vmware -Method GET
}

Function Remove-GlobalPermission {
<#
    .DESCRIPTION Script to add/remove vSphere Global Permission
    .NOTES  Author:  William Lam
    .NOTES  Site:    www.virtuallyghetto.com
    .NOTES  Reference: http://www.virtuallyghetto.com/2017/02/automating-vsphere-global-permissions-with-powercli.html
    .PARAMETER vc_server
        vCenter Server Hostname or IP Address
    .PARAMETER vc_username
        VC Username
    .PARAMETER vc_password
        VC Password
    .PARAMETER vc_user
        Name of the user to remove global permission on
    .EXAMPLE
        PS> Remove-GlobalPermission -vc_server "192.168.1.51" -vc_username "administrator@vsphere.local" -vc_password "VMware1!" -vc_user "VGHETTO\lamw"
#>
    param(
        [Parameter(Mandatory=$true)][string]$vc_server,
        [Parameter(Mandatory=$true)][String]$vc_username,
        [Parameter(Mandatory=$true)][String]$vc_password,
        [Parameter(Mandatory=$true)][String]$vc_user
    )

    $secpasswd = ConvertTo-SecureString $vc_password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($vc_username, $secpasswd)

    # vSphere MOB URL to private enableMethods
    $mob_url = "https://$vc_server/invsvc/mob3/?moid=authorizationService&method=AuthorizationService.RemoveGlobalAccess"

# Ingore SSL Warnings
add-type -TypeDefinition  @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

    # Initial login to vSphere MOB using GET and store session using $vmware variable
    $results = Invoke-WebRequest -Uri $mob_url -SessionVariable vmware -Credential $credential -Method GET

    # Extract hidden vmware-session-nonce which must be included in future requests to prevent CSRF error
    # Credit to https://blog.netnerds.net/2013/07/use-powershell-to-keep-a-cookiejar-and-post-to-a-web-form/ for parsing vmware-session-nonce via Powershell
    if($results.StatusCode -eq 200) {
        $null = $results -match 'name="vmware-session-nonce" type="hidden" value="?([^\s^"]+)"'
        $sessionnonce = $matches[1]
    } else {
        Write-host "Failed to login to vSphere MOB"
        exit 1
    }

    # Escape username
    $vc_user_escaped = [uri]::EscapeUriString($vc_user)

    # The POST data payload must include the vmware-session-nonce variable + URL-encoded
    $body = @"
vmware-session-nonce=$sessionnonce&principals=%3Cprincipals%3E%0D%0A+++%3Cname%3E$vc_user_escaped%3C%2Fname%3E%0D%0A+++%3Cgroup%3Efalse%3C%2Fgroup%3E%0D%0A%3C%2Fprincipals%3E
"@
    # Second request using a POST and specifying our session from initial login + body request
    Write-Host "Removing Global Permission for $vc_user ..."
    $results = Invoke-WebRequest -Uri $mob_url -WebSession $vmware -Method POST -Body $body

    # Logout out of vSphere MOB
    $mob_logout_url = "https://$vc_server/invsvc/mob3/logout"
    $results = Invoke-WebRequest -Uri $mob_logout_url -WebSession $vmware -Method GET
}

### Sample Usage of Enable/Disable functions ###

$vc_server = "192.168.1.51"
$vc_username = "administrator@vsphere.local"
$vc_password = "VMware1!"
$vc_role_id = "-1"
$vc_user = "VGHETTO\lamw"
$propagate = "true"

# Connect to vCenter Server
$server = Connect-VIServer -Server $vc_server -User $vc_username -Password $vc_password

#New-GlobalPermission -vc_server $vc_server -vc_username $vc_username -vc_password $vc_password -vc_user $vc_user -vc_role_id $vc_role_id -propagate $propagate

#Remove-GlobalPermission -vc_server $vc_server -vc_username $vc_username -vc_password $vc_password -vc_user $vc_user

# Disconnect from vCenter Server
Disconnect-viserver $server -confirm:$false﻿Function Add-VMGuestInfo {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
	===========================================================================
    .SYNOPSIS
        Function to add Guestinfo properties to a VM
    .EXAMPLE
        $newGuestProperties = @{
            "guestinfo.foo1" = "bar1"
            "guestinfo.foo2" = "bar2"
            "guestinfo.foo3" = "bar3"
        }

        Add-VMGuestInfo -vmname DeployVM -vmguestinfo $newGuestProperties
#>
    param(
        [Parameter(Mandatory=$true)][String]$vmname,
        [Parameter(Mandatory=$true)][Hashtable]$vmguestinfo
    )

    $vm = Get-VM -Name $vmname
    $currentVMExtraConfig = $vm.ExtensionData.config.ExtraConfig

    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec

    $vmguestinfo.GetEnumerator() | Foreach-Object {
        $optionValue = New-Object VMware.Vim.OptionValue
        $optionValue.Key = $_.Key
        $optionValue.Value = $_.Value
        $currentVMExtraConfig += $optionValue
    }
    $spec.ExtraConfig = $currentVMExtraConfig
    $vm.ExtensionData.ReconfigVM($spec)
}

Function Remove-VMGuestInfo {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
	===========================================================================
    .SYNOPSIS
        Function to remove Guestinfo properties to a VM
    .EXAMPLE
        $newGuestProperties = @{
            "guestinfo.foo1" = "bar1"
            "guestinfo.foo2" = "bar2"
            "guestinfo.foo3" = "bar3"
        }

        Remove-VMGuestInfo -vmname DeployVM -vmguestinfo $newGuestProperties
#>
    param(
        [Parameter(Mandatory=$true)][String]$vmname,
        [Parameter(Mandatory=$true)][Hashtable]$vmguestinfo
    )

    $vm = Get-VM -Name $vmname
    $currentVMExtraConfig = $vm.ExtensionData.config.ExtraConfig

    $updatedVMExtraConfig = @()
    foreach ($vmExtraConfig in $currentVMExtraConfig) {
       if(-not ($vmguestinfo.ContainsKey($vmExtraConfig.key))) {
            $updatedVMExtraConfig += $vmExtraConfig
       }
    }
    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $spec.ExtraConfig = $updatedVMExtraConfig
    $vm.ExtensionData.ReconfigVM($spec)
}# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script to import vCenter Server 6.x root certificate to Mac OS X or NIX* system
# Reference: http://www.virtuallyghetto.com/2016/07/automating-the-import-of-vcenter-server-6-x-root-certificate.html

Function Import-VCRootCertificate ([string]$VC_HOSTNAME) {
    # Set the default download directory to current users desktop
    # Download will be saved as cert.zip
    $DOWNLOAD_PATH=[Environment]::GetFolderPath("Desktop")
    $DOWNLOAD_FILE_NAME="cert.zip"
    $DOWNLOAD_FILE_PATH="$DOWNLOAD_PATH\$DOWNLOAD_FILE_NAME"
    $EXTRACTED_CERTS_PATH="$DOWNLOAD_PATH\certs"

    # VAMI URL, easy way to verify if we have Windows VC or VCSA
    $URL = "https://"+$VC_HOSTNAME+":5480"
    $FOUND_VCSA = 1
	
	try {
		# Checking to see if we have a Windows VC or VCSA
		# as they have different SSL Certificate download endpoints
		$websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
		try {
			Write-Host "`nTesting vCenter URL $URL"
			$result = Invoke-WebRequest -Uri $URL -TimeoutSec 5
		}
		catch [System.NotSupportedException] {
			Write-Host $_.Exception -ForegroundColor "Red" -BackgroundColor "Black"
			throw
		}
		catch [System.Net.WebException] {
			Write-Host $_.Exception
			$FOUND_VCSA = 0
		}

		if($FOUND_VCSA) {
			$VC_CERT_DOWNLOAD_URL="https://"+$VC_HOSTNAME+"/certs/download"
		} else {
			$VC_CERT_DOWNLOAD_URL="https://"+$VC_HOSTNAME+"/certs/download.zip"
		}

		# Required to ingore SSL Warnings
		if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type)
		{
			add-type -TypeDefinition  @"
				using System.Net;
				using System.Security.Cryptography.X509Certificates;
				public class TrustAllCertsPolicy : ICertificatePolicy {
					public bool CheckValidationResult(
						ServicePoint srvPoint, X509Certificate certificate,
						WebRequest request, int certificateProblem) {
						return true;
					}
				}
"@
		}
		[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
		
		# Download VC's SSL Certificate
		Write-Host "`nDownloading VC SSL Certificate from $VC_CERT_DOWNLOAD_URL to $DOWNLOAD_FILE_PATH"
		$webclient = New-Object System.Net.WebClient
		$webclient.DownloadFile("$VC_CERT_DOWNLOAD_URL","$DOWNLOAD_FILE_PATH")

		# Extracting SSL Certificate zip file
		Add-Type -AssemblyName System.IO.Compression.FileSystem
		[System.IO.Compression.ZipFile]::ExtractToDirectory($DOWNLOAD_FILE_PATH, "$DOWNLOAD_PATH")

		# Find SSL certificates ending with .0
		$Dir = get-childitem $EXTRACTED_CERTS_PATH -recurse
		$List = $Dir | where {$_.extension -eq ".0"}

		# Thanks to https://lennytech.wordpress.com/2013/06/18/powershell-install-sp-root-cert-to-trusted-root/ for snippet of code
		# Retrieve Trusted Root Certification Store
		$certStore = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Store Root, LocalMachine

		# Import VC SSL Certificate(s) into cert store
		Write-Host "Importing to VC SSL Certificate to Certificate Store"
		foreach ($a in $list) {
			$file = "$EXTRACTED_CERTS_PATH\$a"

			# Get the certificate from the location where it was placed by the export process
			$cert = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 $file

			# Open the store with maximum allowed privileges
			$certStore.Open("MaxAllowed")

			# Add the certificate to the store
			$certStore.Add($cert)
		}
		# Close the store
		$certStore.Close()
	}
	catch {
		Write-Host -ForegroundColor "Red" -BackgroundColor "Black" $_.Exception
	}
	finally {
		#clean up
		if (Test-Path $DOWNLOAD_FILE_PATH) {
			Write-Host "Cleaning up, deleting $DOWNLOAD_FILE_PATH"
			Remove-Item $DOWNLOAD_FILE_PATH
		}
		if (Test-Path $EXTRACTED_CERTS_PATH) {
			Write-Host "Cleaning up, deleting $EXTRACTED_CERTS_PATH"
			Remove-Item -Recurse -Force $EXTRACTED_CERTS_PATH
		}
	}
}

Import-VCRootCertificate $Args[0]
# Author: William Lam
# Site: www.virtuallyghetto.com
# Description: Script to automate the installation of vRA 7 IaaS Mgmt Agent
# Reference: http://www.virtuallyghetto.com/2016/02/automating-vrealize-automation-7-simple-minimal-part-2-vra-iaas-agent-deployment.html

# Hostname or IP of vRA Appliance
$VRA_APPLIANCE_HOSTNAME = "vra-appliance.primp-industries.com"
# Username of vRA Appliance
$VRA_APPLIANCE_USERNAME = "root"
# Password of vRA Appliance
$VRA_APPLIANCE_PASSWORD = "VMware1!"
# Path to store vRA Agent on IaaS Mgmt Windows system
$VRA_APPLIANCE_AGENT_DOWNLOAD_PATH = "C:\Windows\Temp\vCAC-IaaSManagementAgent-Setup.msi"
# Path to store vRA Agent installer logs on IaaS Mgmt Windowssystem
$VRA_APPLIANCE_AGENT_INSTALL_LOG = "C:\Windows\Temp\ManagementAgent-Setup.log"

# Credentials to the vRA IaaS Windows System
$VRA_IAAS_SERVICE_USERNAME = "vra-iaas\\Administrator"
$VRA_IAAS_SERVICE_PASSWORD = "!MySuperDuperPassword!"

### DO NOT EDIT BEYOND HERE ###

# URL to vRA Agent on vRA Appliance
$VRA_APPLIANCE_AGENT_URL = "https://" + $VRA_APPLIANCE_HOSTNAME + ":5480/installer/download/vCAC-IaaSManagementAgent-Setup.msi"

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
$webclient = New-Object System.Net.WebClient
$webclient.Credentials = New-Object System.Net.NetworkCredential($VRA_APPLIANCE_USERNAME,$VRA_APPLIANCE_PASSWORD)

Write-Host "Downloading " $VRA_APPLIANCE_AGENT_URL "to" $VRA_APPLIANCE_AGENT_DOWNLOAD_PATH "..."
$webclient.DownloadFile($VRA_APPLIANCE_AGENT_URL,$VRA_APPLIANCE_AGENT_DOWNLOAD_PATH)

# Extracting SSL Thumbprint frmo vRA Appliance
# Thanks to Brian Graf for this snippet!
# I originally used this longer snippet from Alan Renouf (https://communities.vmware.com/thread/501913?start=0&tstart=0)
# Brian 1, Alan 0 ;)
# It's still easier in Linux :D
$VRA_APPLIANCE_ENDPOINT = "https://" + $VRA_APPLIANCE_HOSTNAME + ":5480"

add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;

    public class IDontCarePolicy : ICertificatePolicy {
        public IDontCarePolicy() {}
        public bool CheckValidationResult(
            ServicePoint sPoint, X509Certificate cert,
            WebRequest wRequest, int certProb) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = new-object IDontCarePolicy
$VRA_APPLIANE_VAMI = [System.Net.Webrequest]::Create("$VRA_APPLIANCE_ENDPOINT")
$VRA_APPLIANCE_SSL_THUMBPRINT = $VRA_APPLIANE_VAMI.ServicePoint.Certificate.GetCertHashString()

# Extracting vRA IaaS Windows VM hostname
$VRA_IAAS_HOSTNAME=hostname

# Arguments to silent installer for vRA IaaS Agent
$VRA_INSTALLER_ARGS = "/i $VRA_APPLIANCE_AGENT_DOWNLOAD_PATH /qn /norestart /Lvoicewarmup! `"$VRA_APPLIANCE_AGENT_INSTALL_LOG`" ADDLOCAL=`"ALL`" INSTALLLOCATION=`"C:\\Program Files (x86)\\VMware\\vCAC\\Management Agent`" MANAGEMENT_ENDPOINT_ADDRESS=`"$VRA_APPLIANCE_ENDPOINT`" MANAGEMENT_ENDPOINT_THUMBPRINT=`"$VRA_APPLIANCE_SSL_THUMBPRINT`" SERVICE_USER_NAME=`"$VRA_IAAS_SERVICE_USERNAME`" SERVICE_USER_PASSWORD=`"$VRA_IAAS_SERVICE_PASSWORD`" VA_USER_NAME=`"$VRA_APPLIANCE_USERNAME`" VA_USER_PASSWORD=`"$VRA_APPLIANCE_PASSWORD`" CURRENT_MACHINE_FQDN=`"$VRA_IAAS_HOSTNAME`""

if (Test-Path $VRA_APPLIANCE_AGENT_DOWNLOAD_PATH) {
    Write-Host "Installing vRA 7 Agent ..."
    # Exit code of 0 = success
    $ec = (Start-Process -FilePath msiexec.exe -ArgumentList $VRA_INSTALLER_ARGS -Wait -Passthru).ExitCode
    if ($ec -eq 0) {
        Write-Host "Installation successful!`n"
    } else {
        Write-Host "Installation failed, please have a look at the log!`n"
    }
} else {
    Write-host "Download must have failed as I can not find the file!`n"
}
﻿Function Get-Esxconfig {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function remotely downloads /etc/vmware/config and outputs the content
    .PARAMETER VMHostName
        The name of an individual ESXi host
    .PARAMETER ClusterName
        The name vSphere Cluster
    .EXAMPLE
        Get-Esxconfig
    .EXAMPLE
        Get-Esxconfig -ClusterName cluster-01
    .EXAMPLE
        Get-Esxconfig -VMHostName esxi-01
#>
    param(
        [Parameter(Mandatory=$false)][String]$VMHostName,
        [Parameter(Mandatory=$false)][String]$ClusterName
    )

    if($ClusterName) {
        $cluster = Get-View -ViewType ClusterComputeResource -Property Name,Host -Filter @{"name"=$ClusterName}
        $vmhosts = Get-View $cluster.Host -Property Name
    } elseif($VMHostName) {
        $vmhosts = Get-View -ViewType HostSystem -Property Name -Filter @{"name"=$VMHostName}
    } else {
        $vmhosts = Get-View -ViewType HostSystem -Property Name
    }

    foreach ($vmhost in $vmhosts) {
        $vmhostIp = $vmhost.Name

        $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

        # URL to ESXi's esx.conf configuration file (can use any that show up under https://esxi_ip/host)
        $url = "https://$vmhostIp/host/vmware_config"

        # URL to the ESXi configuration file
        $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
        $spec.Method = "httpGet"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

        # Append the cookie generated from VC
        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie = New-Object System.Net.Cookie
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhost.name
        $websession.Cookies.Add($cookie)

        # Retrieve file
        $result = Invoke-WebRequest -Uri $url -WebSession $websession
        Write-Host "Contents of /etc/vmware/config for $vmhostIp ...`n"
        return $result.content
    }
}

Function Remove-IntelSightingsWorkaround {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function removes the Intel Sightings workaround on an ESXi host as outline by https://kb.vmware.com/s/article/52345
    .PARAMETER AffectedHostList
        Text file containing ESXi Hostnames/IP for hosts you wish to remove remediation
    .EXAMPLE
        Remove-IntelSightingsWorkaround -AffectedHostList hostlist.txt
#>
    param(
        [Parameter(Mandatory=$true)][String]$AffectedHostList
    )

    Function UpdateESXConfig {
        param(
            $VMHost
        )

        $vmhostName = $vmhost.name

        $url = "https://$vmhostName/host/vmware_config"

        $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

        $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
        $spec.Method = "httpGet"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie = New-Object System.Net.Cookie
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhost.name
        $websession.Cookies.Add($cookie)
        $result = Invoke-WebRequest -Uri $url -WebSession $websession
        $esxconfig = $result.content

        # Download the current config file to verify we have not already remediated
        # If not, store existing configuration and append new string
        $remediated = $false
        $newVMwareConfig = ""
        foreach ($line in $esxconfig.Split("`n")) {
            if($line -eq 'cpuid.7.edx = "----:00--:----:----:----:----:----:----"') {
                $remediated = $true
            } else {
                $newVMwareConfig+="$line`n"
            }
        }

        if($remediated -eq $true) {
            Write-Host "`tUpdating /etc/vmware/config ..."

            $newVMwareConfig = $newVMwareConfig.TrimEnd()
            $newVMwareConfig += "`n"

            # Create HTTP PUT spec
            $spec.Method = "httpPut"
            $spec.Url = $url
            $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

            # Upload sfcb.cfg back to ESXi host
            $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
            $cookie.Name = "vmware_cgi_ticket"
            $cookie.Value = $ticket.id
            $cookie.Domain = $vmhost.name
            $websession.Cookies.Add($cookie)
            $result = Invoke-WebRequest -Uri $url -WebSession $websession -Body $newVMwareConfig -Method Put -ContentType "plain/text"
            if($result.StatusCode -eq 200) {
                Write-Host "`tSuccessfully updated VMware config file"
            } else {
                Write-Host "Failed to upload VMware config file"
                break
            }
        } else {
            Write-Host "Remedation not found, skipping host"
        }
    }

    if (Test-Path -Path $AffectedHostList) {
        $affectedHosts = Get-Content -Path $AffectedHostList
        foreach ($affectedHost in $affectedHosts) {
            try {
                $vmhost = Get-View -ViewType HostSystem -Property Name -Filter @{"name"=$affectedHost}
                Write-Host "Processing $affectedHost..."
                UpdateESXConfig -vmhost $vmhost
            } catch {
                Write-Host -ForegroundColor Yellow "Unable to find $affectedHost, skipping ..."
            }
        }
    } else {
        Write-Host -ForegroundColor Red "Can not find $AffectedHostList ..."
    }
}

Function Set-IntelSightingsWorkaround {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function removes the Intel Sightings workaround on an ESXi host as outline by https://kb.vmware.com/s/article/52345
    .PARAMETER AffectedHostList
        Text file containing ESXi Hostnames/IP for hosts you wish to apply remediation
    .EXAMPLE
        Set-IntelSightingsWorkaround -AffectedHostList hostlist.txt
#>
    param(
        [Parameter(Mandatory=$true)][String]$AffectedHostList
    )

    Function UpdateESXConfig {
        param(
            $vmhost
        )

        $vmhostName = $vmhost.name

        $url = "https://$vmhostName/host/vmware_config"

        $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

        $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
        $spec.Method = "httpGet"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie = New-Object System.Net.Cookie
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhostName
        $websession.Cookies.Add($cookie)
        $result = Invoke-WebRequest -Uri $url -WebSession $websession
        $esxconfig = $result.content

        # Download the current config file to verify we have not already remediated
        # If not, store existing configuration and append new string
        $remediated = $false
        $newVMwareConfig = ""
        foreach ($line in $esxconfig.Split("`n")) {
            if($line -eq 'cpuid.7.edx = "----:00--:----:----:----:----:----:----"') {
                $remediated = $true
                break
            } else {
                $newVMwareConfig+="$line`n"
            }
        }

        if($remediated -eq $false) {
            Write-Host "`tUpdating /etc/vmware/config ..."

            $newVMwareConfig = $newVMwareConfig.TrimEnd()
            $newVMwareConfig+="`ncpuid.7.edx = `"----:00--:----:----:----:----:----:----`"`n"

            # Create HTTP PUT spec
            $spec.Method = "httpPut"
            $spec.Url = $url
            $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

            # Upload sfcb.cfg back to ESXi host
            $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
            $cookie.Name = "vmware_cgi_ticket"
            $cookie.Value = $ticket.id
            $cookie.Domain = $vmhostName
            $websession.Cookies.Add($cookie)
            $result = Invoke-WebRequest -Uri $url -WebSession $websession -Body $newVMwareConfig -Method Put -ContentType "plain/text"
            if($result.StatusCode -eq 200) {
                Write-Host "`tSuccessfully updated VMware config file"
            } else {
                Write-Host "Failed to upload VMware config file"
                break
            }
        } else {
            Write-Host "Remedation aleady applied, skipping host"
        }
    }

    if (Test-Path -Path $AffectedHostList) {
        $affectedHosts = Get-Content -Path $AffectedHostList
        foreach ($affectedHost in $affectedHosts) {
            try {
                $vmhost = Get-View -ViewType HostSystem -Property Name -Filter @{"name"=$affectedHost}
                Write-Host "Processing $affectedHost..."
                UpdateESXConfig -vmhost $vmhost
            } catch {
                Write-Host -ForegroundColor Yellow "Unable to find $affectedHost, skipping ..."
            }
        }
    } else {
        Write-Host -ForegroundColor Red "Can not find $AffectedHostList ..."
    }
}Function List-VSANDatastoreFolders {
    # List-DatastoreFolders -DatastoreName WorkloadDatastore
    Param (
        [Parameter(Mandatory=$true)][String]$DatastoreName
    )

    $d = Get-Datastore $DatastoreName
    $br = Get-View $d.ExtensionData.Browser
    $spec = new-object VMware.Vim.HostDatastoreBrowserSearchSpec
    $folderFileQuery= New-Object Vmware.Vim.FolderFileQuery
    $spec.Query = $folderFileQuery
    $fileQueryFlags = New-Object VMware.Vim.FileQueryFlags
    $fileQueryFlags.fileOwner = $false
    $fileQueryFlags.fileSize = $false
    $fileQueryFlags.fileType = $true
    $fileQueryFlags.modification = $false
    $spec.details = $fileQueryFlags
    $spec.sortFoldersFirst = $true
    $results = $br.SearchDatastore("[$($d.Name)]",  $spec)

    $folders = @()
    $files = $results.file
    foreach ($file in $files) {
        if($file.getType().Name -eq "FolderFileInfo") {
            $folderPath = $results.FolderPath + " " + $file.Path

            $tmp = [pscustomobject] @{
                Name = $file.FriendlyName;
                Path = $folderPath;
            }
            $folders+=$tmp
        }
    }
    $folders
}﻿Function Get-MacLearn {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retrieves both the legacy security policies as well as the new
        MAC Learning feature and the new security policies which also live under this
        property which was introduced in vSphere 6.7
    .PARAMETER DVPortgroupName
        The name of Distributed Virtual Portgroup(s)
    .EXAMPLE
        Get-MacLearn -DVPortgroupName @("Nested-01-DVPG")
#>
    param(
        [Parameter(Mandatory=$true)][String[]]$DVPortgroupName
    )

    foreach ($dvpgname in $DVPortgroupName) {
        $dvpg = Get-VDPortgroup -Name $dvpgname -ErrorAction SilentlyContinue
        $switchVersion = ($dvpg | Get-VDSwitch).Version
        if($dvpg -and $switchVersion -eq "6.6.0") {
            $securityPolicy = $dvpg.ExtensionData.Config.DefaultPortConfig.SecurityPolicy
            $macMgmtPolicy = $dvpg.ExtensionData.Config.DefaultPortConfig.MacManagementPolicy

            $securityPolicyResults = [pscustomobject] @{
                DVPortgroup = $dvpgname;
                MacLearning = $macMgmtPolicy.MacLearningPolicy.Enabled;
                NewAllowPromiscuous = $macMgmtPolicy.AllowPromiscuous;
                NewForgedTransmits = $macMgmtPolicy.ForgedTransmits;
                NewMacChanges = $macMgmtPolicy.MacChanges;
                Limit = $macMgmtPolicy.MacLearningPolicy.Limit
                LimitPolicy = $macMgmtPolicy.MacLearningPolicy.limitPolicy
                LegacyAllowPromiscuous = $securityPolicy.AllowPromiscuous.Value;
                LegacyForgedTransmits = $securityPolicy.ForgedTransmits.Value;
                LegacyMacChanges = $securityPolicy.MacChanges.Value;
            }
            $securityPolicyResults
        } else {
            Write-Host -ForegroundColor Red "Unable to find DVPortgroup $dvpgname or VDS is not running 6.6.0"
            break
        }
    }
}

Function Set-MacLearn {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function allows you to manage the new MAC Learning capablitites in
        vSphere 6.7 along with the updated security policies.
    .PARAMETER DVPortgroupName
        The name of Distributed Virtual Portgroup(s)
    .PARAMETER EnableMacLearn
        Boolean to enable/disable MAC Learn
    .PARAMETER EnablePromiscuous
        Boolean to enable/disable the new Prom. Mode property
    .PARAMETER EnableForgedTransmit
        Boolean to enable/disable the Forged Transmit property
    .PARAMETER EnableMacChange
        Boolean to enable/disable the MAC Address change property
    .PARAMETER AllowUnicastFlooding
        Boolean to enable/disable Unicast Flooding (Default $true)
    .PARAMETER Limit
        Define the maximum number of learned MAC Address, maximum is 4096 (default 4096)
    .PARAMETER LimitPolicy
        Define the policy (DROP/ALLOW) when max learned MAC Address limit is reached (default DROP)
    .EXAMPLE
        Set-MacLearn -DVPortgroupName @("Nested-01-DVPG") -EnableMacLearn $true -EnablePromiscuous $false -EnableForgedTransmit $true -EnableMacChange $false
#>
    param(
        [Parameter(Mandatory=$true)][String[]]$DVPortgroupName,
        [Parameter(Mandatory=$true)][Boolean]$EnableMacLearn,
        [Parameter(Mandatory=$true)][Boolean]$EnablePromiscuous,
        [Parameter(Mandatory=$true)][Boolean]$EnableForgedTransmit,
        [Parameter(Mandatory=$true)][Boolean]$EnableMacChange,
        [Parameter(Mandatory=$false)][Boolean]$AllowUnicastFlooding=$true,
        [Parameter(Mandatory=$false)][Int]$Limit=4096,
        [Parameter(Mandatory=$false)][String]$LimitPolicy="DROP"
    )

    foreach ($dvpgname in $DVPortgroupName) {
        $dvpg = Get-VDPortgroup -Name $dvpgname -ErrorAction SilentlyContinue
        $switchVersion = ($dvpg | Get-VDSwitch).Version
        if($dvpg -and $switchVersion -eq "6.6.0") {
            $originalSecurityPolicy = $dvpg.ExtensionData.Config.DefaultPortConfig.SecurityPolicy

            $spec = New-Object VMware.Vim.DVPortgroupConfigSpec
            $dvPortSetting = New-Object VMware.Vim.VMwareDVSPortSetting
            $macMmgtSetting = New-Object VMware.Vim.DVSMacManagementPolicy
            $macLearnSetting = New-Object VMware.Vim.DVSMacLearningPolicy
            $macMmgtSetting.MacLearningPolicy = $macLearnSetting
            $dvPortSetting.MacManagementPolicy = $macMmgtSetting
            $spec.DefaultPortConfig = $dvPortSetting
            $spec.ConfigVersion = $dvpg.ExtensionData.Config.ConfigVersion

            if($EnableMacLearn) {
                $macMmgtSetting.AllowPromiscuous = $EnablePromiscuous
                $macMmgtSetting.ForgedTransmits = $EnableForgedTransmit
                $macMmgtSetting.MacChanges = $EnableMacChange
                $macLearnSetting.Enabled = $EnableMacLearn
                $macLearnSetting.AllowUnicastFlooding = $AllowUnicastFlooding
                $macLearnSetting.LimitPolicy = $LimitPolicy
                $macLearnsetting.Limit = $Limit

                Write-Host "Enabling MAC Learning on DVPortgroup: $dvpgname ..."
                $task = $dvpg.ExtensionData.ReconfigureDVPortgroup_Task($spec)
                $task1 = Get-Task -Id ("Task-$($task.value)")
                $task1 | Wait-Task | Out-Null
            } else {
                $macMmgtSetting.AllowPromiscuous = $false
                $macMmgtSetting.ForgedTransmits = $false
                $macMmgtSetting.MacChanges = $false
                $macLearnSetting.Enabled = $false

                Write-Host "Disabling MAC Learning on DVPortgroup: $dvpgname ..."
                $task = $dvpg.ExtensionData.ReconfigureDVPortgroup_Task($spec)
                $task1 = Get-Task -Id ("Task-$($task.value)")
                $task1 | Wait-Task | Out-Null
            }
        } else {
            Write-Host -ForegroundColor Red "Unable to find DVPortgroup $dvpgname or VDS is not running 6.6.0"
            break
        }
    }
}﻿Function Get-PlaceholderVM {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
    ===========================================================================
    .DESCRIPTION
        This function retrieves all placeholder VMs that are protected by SRM
#>
	$results = @()
	Foreach ($vm in Get-VM) {
		if($vm.ExtensionData.Summary.Config.ManagedBy.Type -eq "placeholderVm") {
			$tmp = [pscustomobject] @{
				Name = $vm.Name;
				ExtKey = $vm.ExtensionData.Summary.Config.ManagedBy.ExtensionKey;
				Type = $vm.ExtensionData.Summary.Config.ManagedBy.Type
			}
			$results+=$tmp
		}
	}
	$results
}# Author: William Lam
# Website: www.virtuallyghetto
# Product: VMware vSphere
# Description: Script to extract ESXi PCI Device details such as Name, Vendor, VID, DID & SVID
# Reference: http://www.virtuallyghetto.com/2015/05/extracting-vid-did-svid-from-pci-devices-in-esxi-using-vsphere-api.html

$server = Connect-VIServer -Server 192.168.1.60 -User administrator@vghetto.local -Password VMware1!

$vihosts = Get-View -Server $server -ViewType HostSystem -Property Name,Hardware.PciDevice

$devices_results = @()

foreach ($vihost in $vihosts) {
	$pciDevices = $vihost.Hardware.PciDevice
	foreach ($pciDevice in $pciDevices) {
		$details = "" | select HOST, DEVICE, VENDOR, VID, DID, SVID
		$vid = [String]::Format("{0:x}", $pciDevice.VendorId)
		$did = [String]::Format("{0:x}", $pciDevice.DeviceId)
		$svid = [String]::Format("{0:x}", $pciDevice.SubVendorId)		

		$details.HOST = $vihost.Name
		$details.DEVICE = $pciDevice.DeviceName
		$details.VENDOR = $pciDevice.VendorName
		$details.VID = $vid
		$details.DID = $did
		$details.SVID = $svid
		$devices_results += $details
	}
}

$devices_results

Disconnect-VIServer $server -Confirm:$false<#
.SYNOPSIS
   This script demonstrates an xVC-vMotion where a live running Virtual Machine 
   is live migrated between two vCenter Servers which are NOT part of the
   same vCenter SSO Domain which is only available using the vSphere 6.0 API
.NOTES
   File Name  : run-cool-xVC-vMotion.ps1
   Author     : William Lam - @lamw
   Version    : 1.0
.LINK
    http://www.virtuallyghetto.com/2015/02/did-you-know-of-an-additional-cool-vmotion-capability-in-vsphere-6-0.html
.LINK
   https://github.com/lamw

.INPUTS
   sourceVC, sourceVCUsername, sourceVCPassword, 
   destVC, destVCUsername, destVCPassword, destVCThumbprint
   datastorename, clustername, vmhostname, vmnetworkname,
   vmname
.OUTPUTS
   Console output

.PARAMETER sourceVC
   The hostname or IP Address of the source vCenter Server
.PARAMETER sourceVCUsername
   The username to connect to source vCenter Server
.PARAMETER sourceVCPassword
   The password to connect to source vCenter Server
.PARAMETER destVC
   The hostname or IP Address of the destination vCenter Server
.PARAMETER destVCUsername
   The username to connect to the destination vCenter Server
.PARAMETER destVCPassword
   The password to connect to the destination vCenter Server
.PARAMETER destVCThumbprint
   The SSL Thumbprint (SHA1) of the destination vCenter Server (Certificate checking is enabled, ensure hostname/IP matches)
.PARAMETER datastorename
   The destination vSphere Datastore where the VM will be migrated to
.PARAMETER clustername
   The destination vSphere Cluster where the VM will be migrated to
.PARAMETER vmhostname
   The destination vSphere ESXi host where the VM will be migrated to
.PARAMETER vmnetworkname
   The destination vSphere VM Portgroup where the VM will be migrated to
.PARAMETER vmname
   The name of the source VM to be migrated
#>
param
(
   [Parameter(Mandatory=$true)]
   [string]
   $sourceVC,
   [Parameter(Mandatory=$true)]
   [string]
   $sourceVCUsername,
   [Parameter(Mandatory=$true)]
   [string]
   $sourceVCPassword,
   [Parameter(Mandatory=$true)]
   [string]
   $destVC,
   [Parameter(Mandatory=$true)]
   [string]
   $destVCUsername,
   [Parameter(Mandatory=$true)]
   [string]
   $destVCPassword,
   [Parameter(Mandatory=$true)]
   [string]
   $destVCThumbprint, 
   [Parameter(Mandatory=$true)]
   [string]
   $datastorename,
   [Parameter(Mandatory=$true)]
   [string]
   $clustername,
   [Parameter(Mandatory=$true)]
   [string]
   $vmhostname,
   [Parameter(Mandatory=$true)]
   [string]
   $vmnetworkname,
   [Parameter(Mandatory=$true)]
   [string]
   $vmname
);

## DEBUGGING
#$source = "LA"
#$vmname = "vMA" 
#
## LA->NY
#if ( $source -eq "LA") {
#  $sourceVC = "vcenter60-4.primp-industries.com"
#  $sourceVCUsername = "administrator@vghetto.local"
#  $sourceVCPassword = "VMware1!"
#  $destVC = "vcenter60-5.primp-industries.com" 
#  $destVCUsername = "administrator@vsphere.local"
#  $destVCpassword = "VMware1!"
#  $datastorename = "vesxi60-8-local-storage"
#  $clustername = "NY-Cluster" 
#  $vmhostname = "vesxi60-8.primp-industries.com"
#  $destVCThumbprint = "82:D0:CF:B5:CC:EA:FE:AE:03:BE:E9:4B:AC:A2:B0:AB:2F:E3:87:49"
#  $vmnetworkname = "NY-VM-Network"
#} else {
## NY->LA
#  $sourceVC = "vcenter60-5.primp-industries.com"
#  $sourceVCUsername = "administrator@vsphere.local"
#  $sourceVCPassword = "VMware1!"
#  $destVC = "vcenter60-4.primp-industries.com" 
#  $destVCUsername = "administrator@vghetto.local"
#  $destVCpassword = "VMware1!" 
#  $datastorename = "vesxi60-7-local-storage"
#  $clustername = "LA-Cluster" 
#  $vmhostname = "vesxi60-7.primp-industries.com"
#  $destVCThumbprint = "B8:46:B9:F3:6C:1D:97:8C:ED:A0:19:92:94:E6:1B:45:15:65:63:96"
#  $vmnetworkname = "LA-VM-Network"
#}

# Connect to Source vCenter Server
$sourceVCConn = Connect-VIServer -Server $sourceVC -user $sourceVCUsername -password $sourceVCPassword
# Connect to Destination vCenter Server
$destVCConn = Connect-VIServer -Server $destVC -user $destVCUsername -password $destVCpassword

# Source VM to migrate
$vm = Get-View (Get-VM -Server $sourceVCConn -Name $vmname) -Property Config.Hardware.Device
# Dest Datastore to migrate VM to
$datastore = (Get-Datastore -Server $destVCConn -Name $datastorename)
# Dest Cluster to migrate VM to
$cluster = (Get-Cluster -Server $destVCConn -Name $clustername)
# Dest ESXi host to migrate VM to
$vmhost = (Get-VMHost -Server $destVCConn -Name $vmhostname)

# Find Ethernet Device on VM to change VM Networks
$devices = $vm.Config.Hardware.Device
foreach ($device in $devices) {
   if($device -is [VMware.Vim.VirtualEthernetCard]) {
      $vmNetworkAdapter = $device
   }
}

# Relocate Spec for Migration
$spec = New-Object VMware.Vim.VirtualMachineRelocateSpec
$spec.datastore = $datastore.Id
$spec.host = $vmhost.Id
$spec.pool = $cluster.ExtensionData.ResourcePool
# New Service Locator required for Destination vCenter Server when not part of same SSO Domain
$service = New-Object VMware.Vim.ServiceLocator
$credential = New-Object VMware.Vim.ServiceLocatorNamePassword
$credential.username = $destVCusername
$credential.password = $destVCpassword
$service.credential = $credential
$service.instanceUuid = $destVCConn.InstanceUuid
$service.sslThumbprint = $destVCThumbprint
$service.url = "https://$destVC"
$spec.service = $service
# Modify VM Network Adapter to new VM Netework (assumption 1 vNIC, but can easily be modified)
$spec.deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec
$spec.deviceChange[0] = New-Object VMware.Vim.VirtualDeviceConfigSpec
$spec.deviceChange[0].Operation = "edit"
$spec.deviceChange[0].Device = $vmNetworkAdapter
$spec.deviceChange[0].Device.backing.deviceName = $vmnetworkname

Write-Host "`nMigrating $vmname from $sourceVC to $destVC ...`n"
# Issue Cross VC-vMotion 
$task = $vm.RelocateVM_Task($spec,"defaultPriority") 
$task1 = Get-Task -Id ("Task-$($task.value)")
$task1 | Wait-Task -Verbose

# Disconnect from Source/Destination VC
Disconnect-VIServer -Server $sourceVCConn -Confirm:$false
Disconnect-VIServer -Server $destVCConn -Confirm:$false# William Lam
# www.virtuallygheto.com
# Using Guest Operations API to invoke command inside of Nested ESXi VM

Function runGuestOpInESXiVM() {
	param(
		$vm_moref,
		$guest_username, 
		$guest_password,
		$guest_command_path,
		$guest_command_args
	)
	
	# Guest Ops Managers
	$guestOpMgr = Get-View $session.ExtensionData.Content.GuestOperationsManager
	$authMgr = Get-View $guestOpMgr.AuthManager
	$procMgr = Get-View $guestOpMgr.processManager
	
	# Create Auth Session Object
	$auth = New-Object VMware.Vim.NamePasswordAuthentication
	$auth.username = $guest_username
	$auth.password = $guest_password
	$auth.InteractiveSession = $false
	
	# Program Spec
	$progSpec = New-Object VMware.Vim.GuestProgramSpec
	# Full path to the command to run inside the guest
	$progSpec.programPath = "$guest_command_path"
	$progSpec.workingDirectory = "/tmp"
	# Arguments to the command path, must include "++goup=host/vim/tmp" as part of the arguments
	$progSpec.arguments = "++group=host/vim/tmp $guest_command_args"
	
	# Issue guest op command
	$cmd_pid = $procMgr.StartProgramInGuest($vm_moref,$auth,$progSpec)
}

$session = Connect-VIServer -Server 192.168.1.60 -User administrator@vghetto.local -Password VMware1!

$esxi_vm = 'Nested-ESXi6'
$esxi_username = 'root'
$esxi_password = 'vmware123'

$vm = Get-VM $esxi_vm

# commands to run inside of Nested ESXi VM
$command_path = '/bin/python'
$command_args = '/bin/esxcli.py system welcomemsg set -m "vGhetto Was Here"'

Write-Host
Write-Host "Invoking command:" $command_path $command_args "to" $esxi_vm
Write-Host
runGuestOpInESXiVM -vm_moref $vm.ExtensionData.MoRef -guest_username $esxi_username -guest_password $esxi_password -guest_command_path $command_path -guest_command_args $command_args

Disconnect-VIServer -Server $session -Confirm:$false﻿Function Get-SecureBoot {
    <#
    .SYNOPSIS Query Seure Boot setting for a VM in vSphere 6.5
    .NOTES  Author:  William Lam
    .NOTES  Site:    www.virtuallyghetto.com
    .PARAMETER Vm
      VM to query Secure Boot setting
    .EXAMPLE
      Get-VM -Name Windows10 | Get-SecureBoot
    #>
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$Vm
     )

     $secureBootSetting = if ($vm.ExtensionData.Config.BootOptions.EfiSecureBootEnabled) { "enabled" } else { "disabled" }
     Write-Host "Secure Boot is" $secureBootSetting
}

Function Set-SecureBoot {
    <#
    .SYNOPSIS Enable/Disable Seure Boot setting for a VM in vSphere 6.5
    .NOTES  Author:  William Lam
    .NOTES  Site:    www.virtuallyghetto.com
    .PARAMETER Vm
      VM to enable/disable Secure Boot
    .EXAMPLE
      Get-VM -Name Windows10 | Set-SecureBoot -Enabled
    .EXAMPLE
      Get-VM -Name Windows10 | Set-SecureBoot -Disabled
    #>
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$Vm,
        [Switch]$Enabled,
        [Switch]$Disabled
     )

    if($Enabled) {
        $secureBootSetting = $true
        $reconfigMessage = "Enabling Secure Boot for $Vm"
    }
    if($Disabled) {
        $secureBootSetting = $false
        $reconfigMessage = "Disabling Secure Boot for $Vm"
    }

    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $bootOptions = New-Object VMware.Vim.VirtualMachineBootOptions
    $bootOptions.EfiSecureBootEnabled = $secureBootSetting
    $spec.BootOptions = $bootOptions
  
    Write-Host "`n$reconfigMessage ..."
    $task = $vm.ExtensionData.ReconfigVM_Task($spec)
    $task1 = Get-Task -Id ("Task-$($task.value)")
    $task1 | Wait-Task | Out-Null
}
﻿# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Configure SMP-FT for a Virtual Machine in vSphere 6.0
# Reference: http://www.virtuallyghetto.com/2016/02/new-vsphere-6-0-api-for-configuring-smp-ft.html

<#
.SYNOPSIS  Configure SMP-FT for a Virtual Machine
.DESCRIPTION The function will allow you to enable/disable SMP-FT for a Virtual Machine
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.PARAMETER Vmname
  Virtual Machine object to perform SMP-FT operation
.PARAMETER Operation
  on/off
.PARAMETER Datastore
  The Datastore to store secondary VM as well as the VM's configuration file (Default assumes same datastore but this can be changed)
.PARAMETER Vmhost
  The ESXi host in which to store the secondary VM
.EXAMPLE
  PS> Set-FT -vmname "SMP-VM" -Operation [on|off] -Datastore "vsanDatastore" -Vmhost "vesxi60-5.primp-industries.com"
#>

Function Set-FT {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    $vmname,
    $operation,
    $datastore,
    $vmhost
    )

    process {
        # Retrieve VM View
        $vmView = Get-View -ViewType VirtualMachine -Property Name,Config.Hardware.Device -Filter @{"name"=$vmname}

        # Retrieve Datastore View
        $datastoreView = Get-View -ViewType Datastore -Property Name -Filter @{"name"=$datastore}

        # Retrieve ESXi View
        $vmhostView = Get-View -ViewType HostSystem -Property Name -Filter @{"name"=$vmhost}

        # VM Devices
        $devices = $vmView.Config.Hardware.Device

        $diskArray = @()
        # Build VM Disk Array to map to datastore
        foreach ($device in $d) {
	        if($device -is [VMware.Vim.VirtualDisk]) {
		        $temp = New-Object Vmware.Vim.FaultToleranceDiskSpec
                $temp.Datastore = $datastoreView.Moref
                $temp.Disk = $device
                $diskArray += $temp
	        }
        }

        # FT Config Spec
        $spec = New-Object VMware.Vim.FaultToleranceConfigSpec
        $metadataSpec = New-Object VMware.Vim.FaultToleranceMetaSpec
        $metadataSpec.metaDataDatastore = $datastoreView.MoRef
        $secondaryVMSepc = New-Object VMware.Vim.FaultToleranceVMConfigSpec
        $secondaryVMSepc.vmConfig = $datastoreView.MoRef
        $secondaryVMSepc.disks = $diskArray
        $spec.metaDataPath = $metadataSpec
        $spec.secondaryVmSpec = $secondaryVMSepc

        if($operation -eq "on") {
            $task = $vmView.CreateSecondaryVMEx_Task($vmhostView.MoRef,$spec)
        } elseif($operation -eq "off") {
            $task = $vmView.TurnOffFaultToleranceForVM_Task()
        } else {
            Write-Host "Invalid Selection"
            exit 1
        }
        $task1 = Get-Task -Id ("Task-$($task.value)")
        $task1 | Wait-Task
    }
}
<#
.SYNOPSIS  Applies a VSAN VM Storage Policy across a list of Virtual Machines
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.EXAMPLE
  PS> Set-VSANPolicy -listofvms $arrayofvmnames -policy $vsanpolicyname
#>

Function Set-VSANPolicy {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [string[]]$listofvms,
    [String]$policy
    )

    $vmstoremediate = @()
    foreach ($vm in $listofvms) {
        $hds = Get-VM $vm | Get-HardDisk
        Write-Host "`nApplying VSAN VM Storage Policy:" $policy "to" $vm "..."
        Set-SpbmEntityConfiguration -Configuration (Get-SpbmEntityConfiguration $hds) -StoragePolicy $policy
    }
}

Connect-VIServer -Server 192.168.1.51 -User administrator@vghetto.local -password VMware1! | Out-Null

# Define list of VMs you wish to remediate and apply VSAN VM Storage Policy
$listofvms = @(
"Photon-Deployed-From-WebClient-Multiple-Disks-1",
"Photon-Deployed-From-WebClient-Multiple-Disks-2"
)

# Name of VSAN VM Storage Policy to apply
$vsanpolicy = "Virtual SAN Default Storage Policy"

Set-VSANPolicy -listofvms $listofvms -policy $vsanpolicy

Disconnect-VIServer * -Confirm:$false
# Author: William Lam
# Website: www.virtuallyghetto
# Product: VMware vCenter Server Apppliance
# Description: PowerCLI script to deploy VCSA directly to ESXi host
# Reference: http://www.virtuallyghetto.com/2014/06/an-alternate-way-to-inject-ovf-properties-when-deploying-virtual-appliances-directly-onto-esxi.html

$esxname = "mini.primp-industries.com"
$esx = Connect-VIServer -Server $esxname

# Name of VM
$vmname = "VCSA"

# Name of the OVF Env VM Adv Setting
$ovfenv_key = “guestinfo.ovfEnv”

# VCSA Example
$ovfvalue = "<?xml version=`"1.0`" encoding=`"UTF-8`"?> 
<Environment 
     xmlns=`"http://schemas.dmtf.org/ovf/environment/1`" 
     xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" 
     xmlns:oe=`"http://schemas.dmtf.org/ovf/environment/1`" 
     xmlns:ve=`"http://www.vmware.com/schema/ovfenv`" 
     oe:id=`"`">
   <PlatformSection> 
      <Kind>VMware ESXi</Kind> 
      <Version>5.5.0</Version> 
      <Vendor>VMware, Inc.</Vendor> 
      <Locale>en</Locale> 
   </PlatformSection> 
   <PropertySection> 
         <Property oe:key=`"vami.DNS.VMware_vCenter_Server_Appliance`" oe:value=`"192.168.1.1`"/> 
         <Property oe:key=`"vami.gateway.VMware_vCenter_Server_Appliance`" oe:value=`"192.168.1.1`"/> 
         <Property oe:key=`"vami.hostname`" oe:value=`"vcsa.primp-industries.com`"/> 
         <Property oe:key=`"vami.ip0.VMware_vCenter_Server_Appliance`" oe:value=`"192.168.1.250`"/> 
         <Property oe:key=`"vami.netmask0.VMware_vCenter_Server_Appliance`" oe:value=`"255.255.255.0`"/>  
         <Property oe:key=`"vm.vmname`" oe:value=`"VMware_vCenter_Server_Appliance`"/>
   </PropertySection>
</Environment>"

# Adds "guestinfo.ovfEnv" VM Adv setting to VM
Get-VM $vmname | New-AdvancedSetting -Name $ovfenv_key -Value $ovfvalue -Confirm:$false -Force:$true

Disconnect-VIServer -Server $esx -Confirm:$false
# Author: William Lam
# Website: www.virtuallyghetto
# Product: VMware vSphere
# Description: Script to issue UNMAP command on specified VMFS datastore
# Reference: http://www.virtuallyghetto.com/2014/09/want-to-issue-a-vaai-unmap-operation-using-the-vsphere-web-client.html

param
(
   [Parameter(Mandatory=$true)]
   [VMware.VimAutomation.ViCore.Types.V1.DatastoreManagement.Datastore]
   $datastore,
   [Parameter(Mandatory=$true)]
   [string]
   $numofvmfsblocks
);

# Retrieve a random ESXi host which has access to the selected Datastore
$esxi = (Get-View (($datastore.ExtensionData.Host | Get-Random).key) -Property Name).name

# Retrieve ESXCLI instance from the selected ESXi host
$esxcli = Get-EsxCli -Server $global:DefaultVIServer -VMHost $esxi

# Reclaim based on the number of blocks specified by user
$esxcli.storage.vmfs.unmap($numofvmfsblocks,$datastore,$null)
﻿Function Get-VCVersion {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function extracts the vCenter Server (Windows or VCSA) build from your env
        and maps it to https://kb.vmware.com/kb/2143838 to retrieve the version and release date
    .EXAMPLE
        Get-VCVersion
#>
    param(
        [Parameter(Mandatory=$false)][VMware.VimAutomation.ViCore.Util10.VersionedObjectImpl]$Server
    )

    # Pulled from https://kb.vmware.com/kb/2143838
    $vcenterBuildVersionMappings = @{
        "5973321"="vCenter 6.5 Update 1,2017-07-27"
        "5705665"="vCenter 6.5 0e Express Patch 3,2017-06-15"
        "5318154"="vCenter 6.5 0d Express Patch 2,2017-04-18"
        "5318200"="vCenter 6.0 Update 3b,2017-04-13"
        "5183549"="vCenter 6.0 Update 3a,2017-03-21"
        "5112527"="vCenter 6.0 Update 3,2017-02-24"
        "4541947"="vCenter 6.0 Update 2a,2016-11-22"
        "3634793"="vCenter 6.0 Update 2,2016-03-16"
        "3339083"="vCenter 6.0 Update 1b,2016-01-07"
        "3018524"="vCenter 6.0 Update 1,2015-09-10"
        "2776511"="vCenter 6.0.0b,2015-07-07"
        "2656760"="vCenter 6.0.0a,2015-04-16"
        "2559268"="vCenter 6.0 GA,2015-03-12"
        "4180647"="vCenter 5.5 Update 3e,2016-08-04"
        "3721164"="vCenter 5.5 Update 3d,2016-04-14"
        "3660016"="vCenter 5.5 Update 3c,2016-03-29"
        "3252642"="vCenter 5.5 Update 3b,2015-12-08"
        "3142196"="vCenter 5.5 Update 3a,2015-10-22"
        "3000241"="vCenter 5.5 Update 3,2015-09-16"
        "2646482"="vCenter 5.5 Update 2e,2015-04-16"
        "2001466"="vCenter 5.5 Update 2,2014-09-09"
        "1945274"="vCenter 5.5 Update 1c,2014-07-22"
        "1891313"="vCenter 5.5 Update 1b,2014-06-12"
        "1750787"="vCenter 5.5 Update 1a,2014-04-19"
        "1750596"="vCenter 5.5.0c,2014-04-19"
        "1623099"="vCenter 5.5 Update 1,2014-03-11"
        "1378903"="vCenter 5.5.0a,2013-10-31"
        "1312299"="vCenter 5.5 GA,2013-09-22"
        "3900744"="vCenter 5.1 Update 3d,2016-05-19"
        "3070521"="vCenter 5.1 Update 3b,2015-10-01"
        "2669725"="vCenter 5.1 Update 3a,2015-04-30"
        "2207772"="vCenter 5.1 Update 2c,2014-10-30"
        "1473063"="vCenter 5.1 Update 2,2014-01-16"
        "1364037"="vCenter 5.1 Update 1c,2013-10-17"
        "1235232"="vCenter 5.1 Update 1b,2013-08-01"
        "1064983"="vCenter 5.1 Update 1,2013-04-25"
        "880146"="vCenter 5.1.0a,2012-10-25"
        "799731"="vCenter 5.1 GA,2012-09-10"
        "3891028"="vCenter 5.0 U3g,2016-06-14"
        "3073236"="vCenter 5.0 U3e,2015-10-01"
        "2656067"="vCenter 5.0 U3d,2015-04-30"
        "1300600"="vCenter 5.0 U3,2013-10-17"
        "913577"="vCenter 5.0 U2,2012-12-20"
        "755629"="vCenter 5.0 U1a,2012-07-12"
        "623373"="vCenter 5.0 U1,2012-03-15"
        "5318112"="vCenter 6.5.0c Express Patch 1b,2017-04-13"
        "5178943"="vCenter 6.5.0b,2017-03-14"
        "4944578"="vCenter 6.5.0a Express Patch 01,2017-02-02"
        "4602587"="vCenter 6.5,2016-11-15"
        "5326079"="vCenter 6.0 Update 3b,2017-04-13"
        "5183552"="vCenter 6.0 Update 3a,2017-03-21"
        "5112529"="vCenter 6.0 Update 3,2017-02-24"
        "4541948"="vCenter 6.0 Update 2a,2016-11-22"
        "4191365"="vCenter 6.0 Update 2m,2016-09-15"
        "3634794"="vCenter 6.0 Update 2,2016-03-15"
        "3339084"="vCenter 6.0 Update 1b,2016-01-07"
        "3018523"="vCenter 6.0 Update 1,2015-09-10"
        "2776510"="vCenter 6.0.0b,2015-07-07"
        "2656761"="vCenter 6.0.0a,2015-04-16"
        "2559267"="vCenter 6.0 GA,2015-03-12"
        "4180648"="vCenter 5.5 Update 3e,2016-08-04"
        "3730881"="vCenter 5.5 Update 3d,2016-04-14"
        "3660015"="vCenter 5.5 Update 3c,2016-03-29"
        "3255668"="vCenter 5.5 Update 3b,2015-12-08"
        "3154314"="vCenter 5.5 Update 3a,2015-10-22"
        "3000347"="vCenter 5.5 Update 3,2015-09-16"
        "2646489"="vCenter 5.5 Update 2e,2015-04-16"
        "2442329"="vCenter 5.5 Update 2d,2015-01-27"
        "2183111"="vCenter 5.5 Update 2b,2014-10-09"
        "2063318"="vCenter 5.5 Update 2,2014-09-09"
        "1623101"="vCenter 5.5 Update 1,2014-03-11"
        "1476327"="vCenter 5.5.0b,2013-12-22"
        "1398495"="vCenter 5.5.0a,2013-10-31"
        "1312298"="vCenter 5.5 GA,2013-09-22"
        "3868380"="vCenter 5.1 Update 3d,2016-05-19"
        "3630963"="vCenter 5.1 Update 3c,2016-03-29"
        "3072314"="vCenter 5.1 Update 3b,2015-10-01"
        "2306353"="vCenter 5.1 Update 3,2014-12-04"
        "1882349"="vCenter 5.1 Update 2a,2014-07-01"
        "1474364"="vCenter 5.1 Update 2,2014-01-16"
        "1364042"="vCenter 5.1 Update 1c,2013-10-17"
        "1123961"="vCenter 5.1 Update 1a,2013-05-22"
        "1065184"="vCenter 5.1 Update 1,2013-04-25"
        "947673"="vCenter 5.1.0b,2012-12-20"
        "880472"="vCenter 5.1.0a,2012-10-25"
        "799730"="vCenter 5.1 GA,2012-08-13"
        "3891027"="vCenter 5.0 U3g,2016-06-14"
        "3073237"="vCenter 5.0 U3e,2015-10-01"
        "2656066"="vCenter 5.0 U3d,2015-04-30"
        "2210222"="vCenter 5.0 U3c,2014-11-20"
        "1917469"="vCenter 5.0 U3a,2014-07-01"
        "1302764"="vCenter 5.0 U3,2013-10-17"
        "920217"="vCenter 5.0 U2,2012-12-20"
        "804277"="vCenter 5.0 U1b,2012-08-16"
        "759855"="vCenter 5.0 U1a,2012-07-12"
        "455964"="vCenter 5.0 GA,2011-08-24"
    }

    if(-not $Server) {
        $Server = $global:DefaultVIServer
    }

    $vcBuildNumber = $Server.Build
    $vcName = $Server.Name
    $vcOS = $Server.ExtensionData.Content.About.OsType
    $vcVersion,$vcRelDate = "Unknown","Unknown"

    if($vcenterBuildVersionMappings.ContainsKey($vcBuildNumber)) {
        ($vcVersion,$vcRelDate) = $vcenterBuildVersionMappings[$vcBuildNumber].split(",")
    }

    $tmp = [pscustomobject] @{
        Name = $vcName;
        Build = $vcBuildNumber;
        Version = $vcVersion;
        OS = $vcOS;
        ReleaseDate = $vcRelDate;
    }
    $tmp
}

Function Get-ESXiVersion {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function extracts the ESXi build from your env and maps it to
        https://kb.vmware.com/kb/2143832 to extract the version and release date
    .PARAMETER ClusterName
        Name of the vSphere Cluster to retrieve ESXi version information
    .EXAMPLE
        Get-ESXiVersion -ClusterName VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$ClusterName
    )

    # Pulled from https://kb.vmware.com/kb/2143832
    $esxiBuildVersionMappings = @{
        "5969303"="ESXi 6.5 U1,2017-07-27"
        "5310538"="ESXi 6.5.0d,2017-04-18"
        "5224529"="ESXi 6.5 Express Patch 1a,2017-03-28"
        "5146846"="ESXi 6.5 Patch 01,2017-03-09"
        "4887370"="ESXi 6.5.0a,2017-02-02"
        "4564106"="ESXi 6.5 GA,2016-11-15"
        "5572656"="ESXi 6.0 Patch 5,2017-06-06"
        "5251623"="ESXi 6.0 Express Patch 7c,2017-03-28"
        "5224934"="ESXi 6.0 Express Patch 7a,2017-03-28"
        "5050593"="ESXi 6.0 Update 3,2017-02-24"
        "4600944"="ESXi 6.0 Patch 4,2016-11-22"
        "4510822"="ESXi 6.0 Express Patch 7,2016-10-17"
        "4192238"="ESXi 6.0 Patch 3,2016-08-04"
        "3825889"="ESXi 6.0 Express Patch 6,2016-05-12"
        "3620759"="ESXi 6.0 Update 2,2016-03-16"
        "3568940"="ESXi 6.0 Express Patch 5,2016-02-23"
        "3380124"="ESXi 6.0 Update 1b,2016-01-07"
        "3247720"="ESXi 6.0 Express Patch 4,2015-11-25"
        "3073146"="ESXi 6.0 U1a Express Patch 3,2015-10-06"
        "3029758"="ESXi 6.0 U1,2015-09-10"
        "2809209"="ESXi 6.0.0b,2015-07-07"
        "2715440"="ESXi 6.0 Express Patch 2,2015-05-14"
        "2615704"="ESXi 6.0 Express Patch 1,2015-04-09"
        "2494585"="ESXi 6.0 GA,2015-03-12"
        "5230635"="ESXi 5.5 Express Patch 11,2017-03-28"
        "4722766"="ESXi 5.5 Patch 10,2016-12-20"
        "4345813"="ESXi 5.5 Patch 9,2016-09-15"
        "4179633"="ESXi 5.5 Patch 8,2016-08-04"
        "3568722"="ESXi 5.5 Express Patch 10,2016-02-22"
        "3343343"="ESXi 5.5 Express Patch 9,2016-01-04"
        "3248547"="ESXi 5.5 Update 3b,2015-12-08"
        "3116895"="ESXi 5.5 Update 3a,2015-10-06"
        "3029944"="ESXi 5.5 Update 3,2015-09-16"
        "2718055"="ESXi 5.5 Patch 5,2015-05-08"
        "2638301"="ESXi 5.5 Express Patch 7,2015-04-07"
        "2456374"="ESXi 5.5 Express Patch 6,2015-02-05"
        "2403361"="ESXi 5.5 Patch 4,2015-01-27"
        "2302651"="ESXi 5.5 Express Patch 5,2014-12-02"
        "2143827"="ESXi 5.5 Patch 3,2014-10-15"
        "2068190"="ESXi 5.5 Update 2,2014-09-09"
        "1892794"="ESXi 5.5 Patch 2,2014-07-01"
        "1881737"="ESXi 5.5 Express Patch 4,2014-06-11"
        "1746018"="ESXi 5.5 Update 1a,2014-04-19"
        "1746974"="ESXi 5.5 Express Patch 3,2014-04-19"
        "1623387"="ESXi 5.5 Update 1,2014-03-11"
        "1474528"="ESXi 5.5 Patch 1,2013-12-22"
        "1331820"="ESXi 5.5 GA,2013-09-22"
        "3872664"="ESXi 5.1 Patch 9,2016-05-24"
        "3070626"="ESXi 5.1 Patch 8,2015-10-01"
        "2583090"="ESXi 5.1 Patch 7,2015-03-26"
        "2323236"="ESXi 5.1 Update 3,2014-12-04"
        "2191751"="ESXi 5.1 Patch 6,2014-10-30"
        "2000251"="ESXi 5.1 Patch 5,2014-07-31"
        "1900470"="ESXi 5.1 Express Patch 5,2014-06-17"
        "1743533"="ESXi 5.1 Patch 4,2014-04-29"
        "1612806"="ESXi 5.1 Express Patch 4,2014-02-27"
        "1483097"="ESXi 5.1 Update 2,2014-01-16"
        "1312873"="ESXi 5.1 Patch 3,2013-10-17"
        "1157734"="ESXi 5.1 Patch 2,2013-07-25"
        "1117900"="ESXi 5.1 Express Patch 3,2013-05-23"
        "1065491"="ESXi 5.1 Update 1,2013-04-25"
        "1021289"="ESXi 5.1 Express Patch 2,2013-03-07"
        "914609"="ESXi 5.1 Patch 1,2012-12-20"
        "838463"="ESXi 5.1.0a,2012-10-25"
        "799733"="ESXi 5.1.0 GA,2012-09-10"
        "3982828"="ESXi 5.0 Patch 13,2016-06-14"
        "3086167"="ESXi 5.0 Patch 12,2015-10-01"
        "2509828"="ESXi 5.0 Patch 11,2015-02-24"
        "2312428"="ESXi 5.0 Patch 10,2014-12-04"
        "2000308"="ESXi 5.0 Patch 9,2014-08-28"
        "1918656"="ESXi 5.0 Express Patch 6,2014-07-01"
        "1851670"="ESXi 5.0 Patch 8,2014-05-29"
        "1489271"="ESXi 5.0 Patch 7,2014-01-23"
        "1311175"="ESXi 5.0 Update 3,2013-10-17"
        "1254542"="ESXi 5.0 Patch 6,2013-08-29"
        "1117897"="ESXi 5.0 Express Patch 5,2013-05-15"
        "1024429"="ESXi 5.0 Patch 5,2013-03-28"
        "914586"="ESXi 5.0 Update 2,2012-12-20"
        "821926"="ESXi 5.0 Patch 4,2012-09-27"
        "768111"="ESXi 5.0 Patch 3,2012-07-12"
        "721882"="ESXi 5.0 Express Patch 4,2012-06-14"
        "702118"="ESXi 5.0 Express Patch 3,2012-05-03"
        "653509"="ESXi 5.0 Express Patch 2,2012-04-12"
        "623860"="ESXi 5.0 Update 1,2012-03-15"
        "515841"="ESXi 5.0 Patch 2,2011-12-15"
        "504890"="ESXi 5.0 Express Patch 1,2011-11-03"
        "474610"="ESXi 5.0 Patch 1,2011-09-13"
        "469512"="ESXi 5.0 GA,2011-08-24"
    }

    $cluster = Get-Cluster -Name $ClusterName -ErrorAction SilentlyContinue
    if($cluster -eq $null) {
        Write-Host -ForegroundColor Red "Error: Unable to find vSAN Cluster $ClusterName ..."
        break
    }

    $results = @()
    foreach ($vmhost in $cluster.ExtensionData.Host) {
        $vmhost_view = Get-View $vmhost -Property Name, Config, ConfigManager.ImageConfigManager

        $esxiName = $vmhost_view.name
        $esxiBuild = $vmhost_view.Config.Product.Build
        $esxiVersionNumber = $vmhost_view.Config.Product.Version
        $esxiVersion,$esxiRelDate,$esxiOrigInstallDate = "Unknown","Unknown","N/A"

        if($esxiBuildVersionMappings.ContainsKey($esxiBuild)) {
            ($esxiVersion,$esxiRelDate) = $esxiBuildVersionMappings[$esxiBuild].split(",")
        }

        # Install Date API was only added in 6.5
        if($esxiVersionNumber -eq "6.5.0") {
            $imageMgr = Get-View $vmhost_view.ConfigManager.ImageConfigManager
            $esxiOrigInstallDate = $imageMgr.installDate()
        }

        $tmp = [pscustomobject] @{
            Name = $esxiName;
            Build = $esxiBuild;
            Version = $esxiVersion;
            ReleaseDate = $esxiRelDate;
            OriginalInstallDate = $esxiOrigInstallDate;
        }
        $results+=$tmp
    }
    $results
}

Function Get-VSANVersion {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function extracts the ESXi build from your env and maps it to
        https://kb.vmware.com/kb/2150753 to extract the vSAN version and release date
    .PARAMETER ClusterName
        Name of a vSAN Cluster to retrieve vSAN version information
    .EXAMPLE
        Get-VSANVersion -ClusterName VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$ClusterName
    )

    # Pulled from https://kb.vmware.com/kb/2150753
    $vsanBuildVersionMappings = @{
        "5969303"="vSAN 6.6.1,ESXi 6.5 Update 1,2017-07-27"
        "5310538"="vSAN 6.6,ESXi 6.5.0d,2017-04-18"
        "5224529"="vSAN 6.5 Express Patch 1a,ESXi 6.5 Express Patch 1a,2017-03-28"
        "5146846"="vSAN 6.5 Patch 01,ESXi 6.5 Patch 01,2017-03-09"
        "4887370"="vSAN 6.5.0a,ESXi 6.5.0a,2017-02-02"
        "4564106"="vSAN 6.5,ESXi 6.5 GA,2016-11-15"
        "5572656"="vSAN 6.2 Patch 5,ESXi 6.0 Patch 5,2017-06-06"
        "5251623"="vSAN 6.2 Express Patch 7c,ESXi 6.0 Express Patch 7c,2017-03-28"
        "5224934"="vSAN 6.2 Express Patch 7a,ESXi 6.0 Express Patch 7a,2017-03-28"
        "5050593"="vSAN 6.2 Update 3,ESXi 6.0 Update 3,2017-02-24"
        "4600944"="vSAN 6.2 Patch 4,ESXi 6.0 Patch 4,2016-11-22"
        "4510822"="vSAN 6.2 Express Patch 7,ESXi 6.0 Express Patch 7,2016-10-17"
        "4192238"="vSAN 6.2 Patch 3,ESXi 6.0 Patch 3,2016-08-04"
        "3825889"="vSAN 6.2 Express Patch 6,ESXi 6.0 Express Patch 6,2016-05-12"
        "3620759"="vSAN 6.2,ESXi 6.0 Update 2,2016-03-16"
        "3568940"="vSAN 6.1 Express Patch 5,ESXi 6.0 Express Patch 5,2016-02-23"
        "3380124"="vSAN 6.1 Update 1b,ESXi 6.0 Update 1b,2016-01-07"
        "3247720"="vSAN 6.1 Express Patch 4,ESXi 6.0 Express Patch 4,2015-11-25"
        "3073146"="vSAN 6.1 U1a (Express Patch 3),ESXi 6.0 U1a (Express Patch 3),2015-10-06"
        "3029758"="vSAN 6.1,ESXi 6.0 U1,2015-09-10"
        "2809209"="vSAN 6.0.0b,ESXi 6.0.0b,2015-07-07"
        "2715440"="vSAN 6.0 Express Patch 2,ESXi 6.0 Express Patch 2,2015-05-14"
        "2615704"="vSAN 6.0 Express Patch 1,ESXi 6.0 Express Patch 1,2015-04-09"
        "2494585"="vSAN 6.0,ESXi 6.0 GA,2015-03-12"
        "5230635"="vSAN 5.5 Express Patch 11,ESXi 5.5 Express Patch 11,2017-03-28"
        "4722766"="vSAN 5.5 Patch 10,ESXi 5.5 Patch 10,2016-12-20"
        "4345813"="vSAN 5.5 Patch 9,ESXi 5.5 Patch 9,2016-09-15"
        "4179633"="vSAN 5.5 Patch 8,ESXi 5.5 Patch 8,2016-08-04"
        "3568722"="vSAN 5.5 Express Patch 10,ESXi 5.5 Express Patch 10,2016-02-22"
        "3343343"="vSAN 5.5 Express Patch 9,ESXi 5.5 Express Patch 9,2016-01-04"
        "3248547"="vSAN 5.5 Update 3b,ESXi 5.5 Update 3b,2015-12-08"
        "3116895"="vSAN 5.5 Update 3a,ESXi 5.5 Update 3a,2015-10-06"
        "3029944"="vSAN 5.5 Update 3,ESXi 5.5 Update 3,2015-09-16"
        "2718055"="vSAN 5.5 Patch 5,ESXi 5.5 Patch 5,2015-05-08"
        "2638301"="vSAN 5.5 Express Patch 7,ESXi 5.5 Express Patch 7,2015-04-07"
        "2456374"="vSAN 5.5 Express Patch 6,ESXi 5.5 Express Patch 6,2015-02-05"
        "2403361"="vSAN 5.5 Patch 4,ESXi 5.5 Patch 4,2015-01-27"
        "2302651"="vSAN 5.5 Express Patch 5,ESXi 5.5 Express Patch 5,2014-12-02"
        "2143827"="vSAN 5.5 Patch 3,ESXi 5.5 Patch 3,2014-10-15"
        "2068190"="vSAN 5.5 Update 2,ESXi 5.5 Update 2,2014-09-09"
        "1892794"="vSAN 5.5 Patch 2,ESXi 5.5 Patch 2,2014-07-01"
        "1881737"="vSAN 5.5 Express Patch 4,ESXi 5.5 Express Patch 4,2014-06-11"
        "1746018"="vSAN 5.5 Update 1a,ESXi 5.5 Update 1a,2014-04-19"
        "1746974"="vSAN 5.5 Express Patch 3,ESXi 5.5 Express Patch 3,2014-04-19"
        "1623387"="vSAN 5.5,ESXi 5.5 Update 1,2014-03-11"
    }

    $cluster = Get-Cluster -Name $ClusterName -ErrorAction SilentlyContinue
    if($cluster -eq $null) {
        Write-Host -ForegroundColor Red "Error: Unable to find vSAN Cluster $ClusterName ..."
        break
    }

    $results = @()
    foreach ($vmhost in $cluster.ExtensionData.Host) {
        $vmhost_view = Get-View $vmhost -Property Name, Config, ConfigManager.ImageConfigManager

        $esxiName = $vmhost_view.name
        $esxiBuild = $vmhost_view.Config.Product.Build
        $esxiVersionNumber = $vmhost_view.Config.Product.Version
        $vsanVersion,$esxiVersion,$esxiRelDate = "Unknown","Unknown","Unknown"

        # Technically as of vSAN 6.2 Mgmt API, this information is already built in natively within
        # the product to retrieve ESXi/VC/vSAN Versions
        # See https://github.com/lamw/vghetto-scripts/blob/master/powershell/VSANVersion.ps1
        if($vsanBuildVersionMappings.ContainsKey($esxiBuild)) {
            ($vsanVersion,$esxiVersion,$esxiRelDate) = $vsanBuildVersionMappings[$esxiBuild].split(",")
        }

        $tmp = [pscustomobject] @{
            Name = $esxiName;
            Build = $esxiBuild;
            VSANVersion = $vsanVersion;
            ESXiVersion = $esxiVersion;
            ReleaseDate = $esxiRelDate;
        }
        $results+=$tmp
    }
    $results
}Function Verify-ESXiMeltdownAccelerationInVM {
<#
    .NOTES
    ===========================================================================
     Created by:    Adam Robinson
     Organization:  University of Michigan
        ===========================================================================
    .DESCRIPTION
        This function helps verify if a virtual machine supports the PCID and INVPCID
        instructions.  These can be passed to guests with hardware version 11+
        and can provide performance improvements to Meltdown mitigation.

        This script can return all VMs or you can specify
        a vSphere Cluster to limit the scope or an individual VM
    .PARAMETER VMName
        The name of an individual Virtual Machine
    .EXAMPLE
        Verify-ESXiMeltdownAccelerationInVM
    .EXAMPLE
        Verify-ESXiMeltdownAccelerationInVM -ClusterName cluster-01
    .EXAMPLE
        Verify-ESXiMeltdownAccelerationInVM -VMName vm-01
#>
    param(
        [Parameter(Mandatory=$false)][String]$VMName,
        [Parameter(Mandatory=$false)][String]$ClusterName
    )

    if($ClusterName) {
        $cluster = Get-View -ViewType ClusterComputeResource -Property Name,ResourcePool -Filter @{"name"=$ClusterName}
        $vms = Get-View ((Get-View $cluster.ResourcePool).VM) -Property Name,Config.Version,Runtime.PowerState,Runtime.FeatureRequirement
    } elseif($VMName) {
        $vms = Get-View -ViewType VirtualMachine -Property Name,Config.Version,Runtime.PowerState,Runtime.FeatureRequirement -Filter @{"name"=$VMName}
    } else {
        $vms = Get-View -ViewType VirtualMachine -Property Name,Config.Version,Runtime.PowerState,Runtime.FeatureRequirement
    }

    $results = @()
    foreach ($vm in $vms | Sort-Object -Property Name) {
        # Only check VMs that are powered on
        if($vm.Runtime.PowerState -eq "poweredOn") {
            $vmDisplayName = $vm.Name
            $vmvHW = $vm.Config.Version

            $PCIDPass = $false
            $INVPCIDPass = $false

            $cpuFeatures = $vm.Runtime.FeatureRequirement
            foreach ($cpuFeature in $cpuFeatures) {
                if($cpuFeature.key -eq "cpuid.PCID") {
                    $PCIDPass = $true
                } elseif($cpuFeature.key -eq "cpuid.INVPCID") {
                    $INVPCIDPass = $true
                }
            }

            $meltdownAcceleration = $false
            if ($PCIDPass -and $INVPCIDPass) {
                $meltdownAcceleration = $true
            }

            $tmp = [pscustomobject] @{
                VM = $vmDisplayName;
                PCID = $PCIDPass;
                INVPCID = $INVPCIDPass;
                vHW = $vmvHW;
                MeltdownAcceleration = $meltdownAcceleration
            }
            $results+=$tmp
        }
    }
    $results | ft
}
Function Verify-ESXiMeltdownAcceleration {
<#
    .NOTES
    ===========================================================================
     Created by:    Adam Robinson
     Organization:  University of Michigan
        ===========================================================================
    .DESCRIPTION
        This function helps verify if the ESXi host supports the PCID and INVPCID
        instructions.  These can be passed to guests with hardware version 11+
        and can provide performance improvements to Meltdown mitigation.

        This script can return all ESXi hosts or you can specify
        a vSphere Cluster to limit the scope or an individual ESXi host
    .PARAMETER VMHostName
        The name of an individual ESXi host
    .PARAMETER ClusterName
        The name vSphere Cluster
    .EXAMPLE
        Verify-ESXiMeltdownAcceleration
    .EXAMPLE
        Verify-ESXiMeltdownAcceleration -ClusterName cluster-01
    .EXAMPLE
        Verify-ESXiMeltdownAcceleration -VMHostName esxi-01
#>
    param(
        [Parameter(Mandatory=$false)][String]$VMHostName,
        [Parameter(Mandatory=$false)][String]$ClusterName
    )

    $accelerationEVCModes = @("intel-broadwell","intel-haswell","Disabled")

    if($ClusterName) {
        $cluster = Get-View -ViewType ClusterComputeResource -Property Name,Host -Filter @{"name"=$ClusterName}
        $vmhosts = Get-View $cluster.Host -Property Name,Config.FeatureCapability,Hardware.CpuFeature,Summary.CurrentEVCModeKey
    } elseif($VMHostName) {
        $vmhosts = Get-View -ViewType HostSystem -Property Name,Config.FeatureCapability,Hardware.CpuFeature,Summary.CurrentEVCModeKey -Filter @{"name"=$VMHostName}
    } else {
        $vmhosts = Get-View -ViewType HostSystem -Property Name,Config.FeatureCapability,Hardware.CpuFeature,Summary.CurrentEVCModeKey
    }

    $results = @()
    foreach ($vmhost in $vmhosts | Sort-Object -Property Name) {
        $vmhostDisplayName = $vmhost.Name

        $evcMode = $vmhost.Summary.CurrentEVCModeKey
        if ($evcMode -eq $null) {
            $evcMode = "Disabled"
        }

        $PCIDPass = $false
        $INVPCIDPass = $false

        #output from $vmhost.Hardware.CpuFeature is a binary string ':' delimited to nibbles
        #the easiest way I could figure out the hex conversion was to make a byte array
        $cpuidEAX = ($vmhost.Hardware.CpuFeature | Where-Object {$_.Level -eq 1}).Eax -Replace ":","" -Split "(?<=\G\d{8})(?=\d{8})"
        $cpuSignature = ($cpuidEAX | Foreach-Object {[System.Convert]::ToByte($_, 2)} | Foreach-Object {$_.ToString("X2")}) -Join ""
        $cpuSignature = "0x" + $cpuSignature

        $cpuFamily = [System.Convert]::ToByte($cpuidEAX[2], 2).ToString("X2")

        $cpuFeatures = $vmhost.Config.FeatureCapability
        foreach ($cpuFeature in $cpuFeatures) {
            if($cpuFeature.key -eq "cpuid.PCID" -and $cpuFeature.value -eq 1) {
                $PCIDPass = $true
            } elseif($cpuFeature.key -eq "cpuid.INVPCID" -and $cpuFeature.value -eq 1) {
                $INVPCIDPass = $true
            }
        }

        $HWv11Acceleration = $false
        if ($cpuFamily -eq "06") {
            if ($PCIDPass -and $INVPCIDPass) {
                if ($accelerationEVCModes -contains $evcMode) {
                    $HWv11Acceleration = $true
                }
                else {
                    $HWv11Acceleration = "EVCTooLow"
                }
            }
        }
        else {
            $HWv11Acceleration = "Unneeded"
        }

        $tmp = [pscustomobject] @{
            VMHost = $vmhostDisplayName;
            PCID = $PCIDPass;
            INVPCID = $INVPCIDPass;
            EVCMode = $evcMode
            "vHW11+Acceleration" = $HWv11Acceleration;
        }
        $results+=$tmp
    }
    $results | ft
}Function Verify-ESXiMicrocodePatchAndVM {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function helps verify both ESXi Patch and Microcode updates have been
        applied as stated per https://kb.vmware.com/s/article/52085

        This script can return all VMs or you can specify
        a vSphere Cluster to limit the scope or an individual VM
    .PARAMETER VMName
        The name of an individual Virtual Machine
    .EXAMPLE
        Verify-ESXiMicrocodePatchAndVM
    .EXAMPLE
        Verify-ESXiMicrocodePatchAndVM -ClusterName cluster-01
    .EXAMPLE
        Verify-ESXiMicrocodePatchAndVM -VMName vm-01
#>
    param(
        [Parameter(Mandatory=$false)][String]$VMName,
        [Parameter(Mandatory=$false)][String]$ClusterName
    )

    if($ClusterName) {
        $cluster = Get-View -ViewType ClusterComputeResource -Property Name,ResourcePool -Filter @{"name"=$ClusterName}
        $vms = Get-View ((Get-View $cluster.ResourcePool).VM) -Property Name,Config.Version,Runtime.PowerState,Runtime.FeatureRequirement
    } elseif($VMName) {
        $vms = Get-View -ViewType VirtualMachine -Property Name,Config.Version,Runtime.PowerState,Runtime.FeatureRequirement -Filter @{"name"=$VMName}
    } else {
        $vms = Get-View -ViewType VirtualMachine -Property Name,Config.Version,Runtime.PowerState,Runtime.FeatureRequirement
    }

    $results = @()
    foreach ($vm in $vms | Sort-Object -Property Name) {
        # Only check VMs that are powered on
        if($vm.Runtime.PowerState -eq "poweredOn") {
            $vmDisplayName = $vm.Name
            $vmvHW = $vm.Config.Version

            $vHWPass = $false
            $IBRSPass = $false
            $IBPBPass = $false
            $STIBPPass = $false
            $vmAffected = $true
            if ($vmvHW -match 'vmx-[0-9]{2}') {
              if ( [int]$vmvHW.Split('-')[-1] -gt 8 ) {
                $vHWPass = $true
              } else {
                $vHWPass = "N/A"
              }

              $cpuFeatures = $vm.Runtime.FeatureRequirement
              foreach ($cpuFeature in $cpuFeatures) {
                  if($cpuFeature.key -eq "cpuid.IBRS") {
                      $IBRSPass = $true
                  } elseif($cpuFeature.key -eq "cpuid.IBPB") {
                      $IBPBPass = $true
                  } elseif($cpuFeature.key -eq "cpuid.STIBP") {
                      $STIBPPass = $true
                  }
              }
              
              if( ($IBRSPass -eq $true -or $IBPBPass -eq $true -or $STIBPPass -eq $true) -and $vHWPass -eq $true) {
                  $vmAffected = $false
              } elseif($vHWPass -eq "N/A") {
                  $vmAffected = $vHWPass
              }
            } else {
              $IBRSPass = "N/A"
              $IBPBPass = "N/A"
              $STIBPPass = "N/A"
              $vmAffected = "N/A"
            }

            $tmp = [pscustomobject] @{
                VM = $vmDisplayName;
                IBRSPresent = $IBRSPass;
                IBPBPresent = $IBPBPass;
                STIBPPresent = $STIBPPass;
                vHW = $vmvHW;
                HypervisorAssistedGuestAffected = $vmAffected;
            }
            $results+=$tmp
        }
    }
    $results | ft
}

Function Verify-ESXiMicrocodePatch {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function helps verify only the ESXi Microcode update has been
        applied as stated per https://kb.vmware.com/s/article/52085

        This script can return all ESXi hosts or you can specify
        a vSphere Cluster to limit the scope or an individual ESXi host
    .PARAMETER VMHostName
        The name of an individual ESXi host
    .PARAMETER ClusterName
        The name vSphere Cluster
    .EXAMPLE
        Verify-ESXiMicrocodePatch
    .EXAMPLE
        Verify-ESXiMicrocodePatch -ClusterName cluster-01
    .EXAMPLE
        Verify-ESXiMicrocodePatch -VMHostName esxi-01
    .EXAMPLE
        Verify-ESXiMicrocodePatch -ClusterName "Virtual SAN Cluster" -IncludeMicrocodeVerCheck $true -PlinkPath "C:\Users\lamw\Desktop\plink.exe" -ESXiUsername "root" -ESXiPassword "foobar"
#>
    param(
        [Parameter(Mandatory=$false)][String]$VMHostName,
        [Parameter(Mandatory=$false)][String]$ClusterName,
        [Parameter(Mandatory=$false)][Boolean]$IncludeMicrocodeVerCheck=$false,
        [Parameter(Mandatory=$false)][String]$PlinkPath,
        [Parameter(Mandatory=$false)][String]$ESXiUsername,
        [Parameter(Mandatory=$false)][String]$ESXiPassword
    )

    if($ClusterName) {
        $cluster = Get-View -ViewType ClusterComputeResource -Property Name,Host -Filter @{"name"=$ClusterName}
        $vmhosts = Get-View $cluster.Host -Property Name,Config.FeatureCapability,Hardware.CpuFeature,Summary.Hardware,ConfigManager.ServiceSystem
    } elseif($VMHostName) {
        $vmhosts = Get-View -ViewType HostSystem -Property Name,Config.FeatureCapability,Hardware.CpuFeature,Summary.Hardware,ConfigManager.ServiceSystem -Filter @{"name"=$VMHostName}
    } else {
        $vmhosts = Get-View -ViewType HostSystem -Property Name,Config.FeatureCapability,Hardware.CpuFeature,Summary.Hardware,ConfigManager.ServiceSystem
    }

    # Merge of tables from https://kb.vmware.com/s/article/52345 and https://kb.vmware.com/s/article/52085
    $procSigUcodeTable = @(
	    [PSCustomObject]@{Name = "Sandy Bridge DT";  procSig = "0x000206a7"; ucodeRevFixed = "0x0000002d"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Sandy Bridge EP";  procSig = "0x000206d7"; ucodeRevFixed = "0x00000713"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Ivy Bridge DT";  procSig = "0x000306a9"; ucodeRevFixed = "0x0000001f"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Ivy Bridge EP";  procSig = "0x000306e4"; ucodeRevFixed = "0x0000042c"; ucodeRevSightings = "0x0000042a"}
	    [PSCustomObject]@{Name = "Ivy Bridge EX";  procSig = "0x000306e7"; ucodeRevFixed = "0x00000713"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Haswell DT";  procSig = "0x000306c3"; ucodeRevFixed = "0x00000024"; ucodeRevSightings = "0x00000023"}
	    [PSCustomObject]@{Name = "Haswell EP";  procSig = "0x000306f2"; ucodeRevFixed = "0x0000003c"; ucodeRevSightings = "0x0000003b"}
	    [PSCustomObject]@{Name = "Haswell EX";  procSig = "0x000306f4"; ucodeRevFixed = "0x00000011"; ucodeRevSightings = "0x00000010"}
	    [PSCustomObject]@{Name = "Broadwell H";  procSig = "0x00040671"; ucodeRevFixed = "0x0000001d"; ucodeRevSightings = "0x0000001b"}
	    [PSCustomObject]@{Name = "Broadwell EP/EX";  procSig = "0x000406f1"; ucodeRevFixed = "0x0b00002a"; ucodeRevSightings = "0x0b000025"}
	    [PSCustomObject]@{Name = "Broadwell DE";  procSig = "0x00050662"; ucodeRevFixed = "0x00000015"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Broadwell DE";  procSig = "0x00050663"; ucodeRevFixed = "0x07000012"; ucodeRevSightings = "0x07000011"}
	    [PSCustomObject]@{Name = "Broadwell DE";  procSig = "0x00050664"; ucodeRevFixed = "0x0f000011"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Broadwell NS";  procSig = "0x00050665"; ucodeRevFixed = "0x0e000009"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Skylake H/S";  procSig = "0x000506e3"; ucodeRevFixed = "0x000000c2"; ucodeRevSightings = ""} # wasn't actually affected by Sightings, ucode just re-released
	    [PSCustomObject]@{Name = "Skylake SP";  procSig = "0x00050654"; ucodeRevFixed = "0x02000043"; ucodeRevSightings = "0x0200003A"}
	    [PSCustomObject]@{Name = "Kaby Lake H/S/X";  procSig = "0x000906e9"; ucodeRevFixed = "0x00000084"; ucodeRevSightings = "0x0000007C"}
	    [PSCustomObject]@{Name = "Zen EPYC";  procSig = "0x00800f12"; ucodeRevFixed = "0x08001227"; ucodeRevSightings = ""}
    )

    # Remote SSH commands for retrieving current ESXi host microcode version
    $plinkoptions = "-ssh -pw $ESXiPassword"
    $cmd = "vsish -e cat /hardware/cpu/cpuList/0 | grep `'Current Revision:`'"
    $remoteCommand = '"' + $cmd + '"'

    $results = @()
    foreach ($vmhost in $vmhosts | Sort-Object -Property Name) {
        $vmhostDisplayName = $vmhost.Name
        $cpuModelName = $($vmhost.Summary.Hardware.CpuModel -replace '\s+', ' ')

        $IBRSPass = $false
        $IBPBPass = $false
        $STIBPPass = $false

        $cpuFeatures = $vmhost.Config.FeatureCapability
        foreach ($cpuFeature in $cpuFeatures) {
            if($cpuFeature.key -eq "cpuid.IBRS" -and $cpuFeature.value -eq 1) {
                $IBRSPass = $true
            } elseif($cpuFeature.key -eq "cpuid.IBPB" -and $cpuFeature.value -eq 1) {
                $IBPBPass = $true
            } elseif($cpuFeature.key -eq "cpuid.STIBP" -and $cpuFeature.value -eq 1) {
                $STIBPPass = $true
            }
        }

        $vmhostAffected = $true
        if($IBRSPass -or $IBPBPass -or $STIBPass) {
           $vmhostAffected = $false
        }

        # Retrieve Microcode version if user specifies which unfortunately requires SSH access
        if($IncludeMicrocodeVerCheck -and $PlinkPath -ne $null -and $ESXiUsername -ne $null -and $ESXiPassword -ne $null) {
            $serviceSystem = Get-View $vmhost.ConfigManager.ServiceSystem
            $services = $serviceSystem.ServiceInfo.Service
            foreach ($service in $services) {
                if($service.Key -eq "TSM-SSH") {
                    $ssh = $service
                    break
                }
            }

            $command = "echo yes | " + $PlinkPath + " " + $plinkoptions + " " + $ESXiUsername + "@" + $vmhost.Name + " " + $remoteCommand

            if($ssh.Running){
                $plinkResults = Invoke-Expression -command $command
                $microcodeVersion = $plinkResults.split(":")[1]
            } else {
                $microcodeVersion = "SSHNeedsToBeEnabled"
            }
        } else {
            $microcodeVersion = "N/A"
        }

        #output from $vmhost.Hardware.CpuFeature is a binary string ':' delimited to nibbles
        #the easiest way I could figure out the hex conversion was to make a byte array
        $cpuidEAX = ($vmhost.Hardware.CpuFeature | Where-Object {$_.Level -eq 1}).Eax -Replace ":",""
        $cpuidEAXbyte = $cpuidEAX -Split "(?<=\G\d{8})(?=\d{8})"
        $cpuidEAXnibble = $cpuidEAX -Split "(?<=\G\d{4})(?=\d{4})"

        $cpuSignature = "0x" + $(($cpuidEAXbyte | Foreach-Object {[System.Convert]::ToByte($_, 2)} | Foreach-Object {$_.ToString("X2")}) -Join "")

        # https://software.intel.com/en-us/articles/intel-architecture-and-processor-identification-with-cpuid-model-and-family-numbers
        $ExtendedFamily = [System.Convert]::ToInt32($($cpuidEAXnibble[1] + $cpuidEAXnibble[2]), 2)
        $Family = [System.Convert]::ToInt32($cpuidEAXnibble[5], 2)

        # output now in decimal, not hex!
        $cpuFamily = $ExtendedFamily + $Family
        $cpuModel = [System.Convert]::ToByte($($cpuidEAXnibble[3] + $cpuidEAXnibble[6]), 2)
        $cpuStepping = [System.Convert]::ToByte($cpuidEAXnibble[7], 2)
               
        
        $intelSighting = "N/A"
        $goodUcode = "N/A"

        # check and compare ucode
        if ($IncludeMicrocodeVerCheck) {
         
            $intelSighting = $false
            $goodUcode = $false
            $matched = $false

            foreach ($cpu in $procSigUcodeTable) {
                if ($cpuSignature -eq $cpu.procSig) {
                    $matched = $true
                    if ($microcodeVersion -eq $cpu.ucodeRevSightings) {
                        $intelSighting = $true
                    } elseif ($microcodeVersion -as [int] -ge $cpu.ucodeRevFixed -as [int]) {
                        $goodUcode = $true
                    }
                }
            } 
            if (!$matched) {
                # CPU is not in procSigUcodeTable, check with BIOS vendor / Intel based procSig or FMS (dec) in output
                $goodUcode = "Unknown"
            }
        }

        $tmp = [pscustomobject] @{
            VMHost = $vmhostDisplayName;
            "CPU Model Name" = $cpuModelName;
            Family = $cpuFamily;
            Model = $cpuModel;
            Stepping = $cpuStepping;
            Microcode = $microcodeVersion;
            procSig = $cpuSignature;
            IBRSPresent = $IBRSPass;
            IBPBPresent = $IBPBPass;
            STIBPPresent = $STIBPPass;
            HypervisorAssistedGuestAffected = $vmhostAffected;
            "Good Microcode" = $goodUcode;
            IntelSighting = $intelSighting;
        }
        $results+=$tmp
    }
    $results | FT *
}
# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script to add a new VMDK w/the MultiWriter Flag enabled in vSphere 6.x
# Reference: http://www.virtuallyghetto.com/2015/10/new-method-of-enabling-multiwriter-vmdk-flag-in-vsphere-6-0-update-1.html

$vcname = "192.168.1.51"
$vcuser = "administrator@vghetto.local"
$vcpass = "VMware1!"

$vmName = "Multi-Writer-VM"
# Syntax: [datastore-name] vm-home-dir/vmdk-name.vmdk
# Use (Get-VM -Name "Multi-Writer-VM").ExtensionData.Layout.Disk to help identify VM-Home-Dir
$vmdkFileNamePath = "[vsanDatastore] f2d16e57-7ecf-bf9f-8a6a-b8aeed7c9e96/Multi-Writer-VM-1.vmdk"
$diskSizeGB = 5
$diskControllerNumber = 0
$diskUnitNumber = 1

#### DO NOT EDIT BEYOND HERE ####

$server = Connect-VIServer -Server $vcname -User $vcuser -Password $vcpass

# Retrieve VM and only its Devices
$vm = Get-View -Server $server -ViewType VirtualMachine -Property Name,Config.Hardware.Device -Filter @{"Name" = $vmName}

# Convert GB to KB
$diskSizeInKB = (($diskSizeGB * 1024 * 1024 * 1024)/1KB)
$diskSizeInKB = [Math]::Round($diskSizeInKB,4,[MidPointRounding]::AwayFromZero)

# Array of Devices on VM
$vmDevices = $vm.Config.Hardware.Device

# Find the SCSI Controller we care about
foreach ($device in $vmDevices) {
	if($device -is [VMware.Vim.VirtualSCSIController] -and $device.BusNumber -eq $diskControllerNumber) {
		$diskControllerKey = $device.key
        break
	}
}

# Create VM Config Spec to add new VMDK & Enable Multi-Writer Flag
$spec = New-Object VMware.Vim.VirtualMachineConfigSpec
$spec.deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec
$spec.deviceChange[0].operation = 'add'
$spec.DeviceChange[0].FileOperation = 'create'
$spec.deviceChange[0].device = New-Object VMware.Vim.VirtualDisk
$spec.deviceChange[0].device.key = -1
$spec.deviceChange[0].device.ControllerKey = $diskControllerKey
$spec.deviceChange[0].device.unitNumber = $diskUnitNumber
$spec.deviceChange[0].device.CapacityInKB = $diskSizeInKB
$spec.DeviceChange[0].device.backing = New-Object VMware.Vim.VirtualDiskFlatVer2BackingInfo
$spec.DeviceChange[0].device.Backing.fileName = $vmdkFileNamePath
$spec.DeviceChange[0].device.Backing.diskMode = "persistent"
$spec.DeviceChange[0].device.Backing.eagerlyScrub = $True
$spec.DeviceChange[0].device.Backing.Sharing = "sharingMultiWriter"

Write-Host "`nAdding new VMDK w/capacity $diskSizeGB GB to VM: $vmname"
$task = $vm.ReconfigVM_Task($spec)
$task1 = Get-Task -Id ("Task-$($task.value)")
$task1 | Wait-Task

Disconnect-VIServer $server -Confirm:$false
# William Lam
# www.virtualyghetto.com

$vcname = "192.168.1.150"
$vcuser = "administrator@vghetto.local"
$vcpass = "VMware1!"
$esxhosts = @("192.168.1.190", "192.168.1.191", "192.168.1.192")
$esxuser = "root"
$esxpass = "VMware1!"
$cluster = "VSAN-Cluster"

#### DO NOT EDIT BEYOND HERE ####

$vcenter = Connect-VIServer $vcname -User $vcuser -Password $vcpass -WarningAction SilentlyContinue

$cluster_ref = Get-Cluster $cluster

$tasks = @()
foreach($esxhost in $esxhosts) {
    Write-Host "Adding $esxhost to $cluster ..."
    Add-VMHost -Name $esxhost -Location $cluster_ref -User $esxuser -Password $esxpass -Force | out-null
}

$spec = New-Object VMware.Vim.ClusterConfigSpecEx
$vsanconfig = New-Object VMware.Vim.VsanClusterConfigInfo
$defaultconfig = New-Object VMware.Vim.VsanClusterConfigInfoHostDefaultInfo
$defaultconfig.AutoClaimStorage = $true
$vsanconfig.DefaultConfig = $defaultconfig
$vsanconfig.enabled = $true
$spec.VsanConfig = $vsanconfig

Write-Host "Enabling VSAN Cluster on $cluster ..."
$task = $cluster_ref.ExtensionData.ReconfigureComputeResource_Task($spec,$true)
$task1 = Get-Task -Id ("Task-$($task.value)")
$task1 | Wait-Task | out-null

Disconnect-VIServer $vcenter -Confirm:$false
# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script demonstrating vSphere MOB Automation using PowerShell
# Reference: http://www.virtuallyghetto.com/2016/07/how-to-automate-vsphere-mob-operations-using-powershell.html

$vc_server = "192.168.1.51"
$vc_username = "administrator@vghetto.local"
$vc_password = "VMware1!"
$mob_url = "https://$vc_server/mob/?moid=VpxSettings&method=queryView"

$secpasswd = ConvertTo-SecureString $vc_password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($vc_username, $secpasswd)

# Ingore SSL Warnings
add-type -TypeDefinition  @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

# Initial login to vSphere MOB using GET and store session using $vmware variable
$results = Invoke-WebRequest -Uri $mob_url -SessionVariable vmware -Credential $credential -Method GET

# Extract hidden vmware-session-nonce which must be included in future requests to prevent CSRF error
# Credit to https://blog.netnerds.net/2013/07/use-powershell-to-keep-a-cookiejar-and-post-to-a-web-form/ for parsing vmware-session-nonce via Powershell
if($results.StatusCode -eq 200) {
    $null = $results -match 'name="vmware-session-nonce" type="hidden" value="?([^\s^"]+)"'
    $sessionnonce = $matches[1]
} else {
    $results
    Write-host "Failed to login to vSphere MOB"
    exit 1
}

# The POST data payload must include the vmware-session-nonce varaible + URL-encoded
$body = @"
vmware-session-nonce=$sessionnonce&name=VirtualCenter.InstanceName
"@

# Second request using a POST and specifying our session from initial login + body request
$results = Invoke-WebRequest -Uri $mob_url -WebSession $vmware -Method POST -Body $body

# Logout out of vSphere MOB
$mob_logout_url = "https://$vc_server/mob/logout"
Invoke-WebRequest -Uri $mob_logout_url -WebSession $vmware -Method GET

# Clean up the results for further processing
# Extract InnerText, split into string array & remove empty lines
$cleanedUpResults = $results.ParsedHtml.body.innertext.split("`n").replace("`"","") | ? {$_.trim() -ne ""}

# Loop through results looking for valuestring which contains the data we want
foreach ($parsedResults in $cleanedUpResults) {
    if($parsedResults -like "valuestring*") {
        $parsedResults.replace("valuestring","")
    }
}
﻿<#
.SYNOPSIS Retrieve the current VMFS Unmap priority for VMFS 6 datastore
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/10/configure-new-automatic-space-reclamation-vmfs-unmap-using-vsphere-6-5-apis.html
.PARAMETER Datastore
  VMFS 6 Datastore to enable or disable VMFS Unamp
.EXAMPLE
  Get-Datastore "mini-local-datastore-hdd" | Get-VMFSUnmap
#>

Function Get-VMFSUnmap {
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [VMware.VimAutomation.ViCore.Impl.V1.DatastoreManagement.DatastoreImpl]$Datastore
     )

     $datastoreInfo = $Datastore.ExtensionData.Info

     if($datastoreInfo -is [VMware.Vim.VmfsDatastoreInfo] -and $datastoreInfo.Vmfs.MajorVersion -eq 6) {
        $datastoreInfo.Vmfs | select Name, UnmapPriority, UnmapGranularity
     } else {
        Write-Host "Not a VMFS Datastore and/or VMFS version is not 6.0"
     }
}

<#
.SYNOPSIS Configure the VMFS Unmap priority for VMFS 6 datastore
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/10/configure-new-automatic-space-reclamation-vmfs-unmap-using-vsphere-6-5-apis.html
.PARAMETER Datastore
  VMFS 6 Datastore to enable or disable VMFS Unamp
.EXAMPLE
  Get-Datastore "mini-local-datastore-hdd" | Set-VMFSUnmap -Enabled $true
.EXAMPLE
  Get-Datastore "mini-local-datastore-hdd" | Set-VMFSUnmap -Enabled $false
#>

Function Set-VMFSUnmap {
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [VMware.VimAutomation.ViCore.Impl.V1.DatastoreManagement.DatastoreImpl]$Datastore,
        [String]$Enabled
     )

    $vmhostView = ($Datastore | Get-VMHost).ExtensionData
    $storageSystem = Get-View $vmhostView.ConfigManager.StorageSystem

    if($Enabled -eq $true) {
        $enableUNMAP = "low"
        $reconfigMessage = "Enabling Automatic VMFS Unmap for $Datastore"
    } else {
        $enableUNMAP = "none"
        $reconfigMessage = "Disabling Automatic VMFS Unmap for $Datastore"
    }

    $uuid = $datastore.ExtensionData.Info.Vmfs.Uuid

    Write-Host "$reconfigMessage ..."
    $storageSystem.UpdateVmfsUnmapPriority($uuid,$enableUNMAP)
}
<#
.SYNOPSIS
This script accepts the name of a VM and the credentials to its
source vCenter Server as well as destination vCenter Server and its
credentials to check if there are any MAC Address conflicts prior to 
issuing a xVC-vMotion of VM (applicable to same and differnet SSO Domain)
.NOTES
File Name : check-vm-mac-conflict.ps1
Author : William Lam - @lamw
Version : 1.0
.LINK
http://www.virtuallyghetto.com/2015/03/duplicate-mac-address-concerns-with-xvc-vmotion-in-vsphere-6-0.html
.LINK
https://github.com/lamw
.INPUTS
sourceVC, sourceVCUsername, sourceVCPassword,destVC, destVCUsername, destVCPassword, vmname
.OUTPUTS
Console output
.PARAMETER sourceVC
The hostname or IP Address of the source vCenter Server
.PARAMETER sourceVCUsername
The username to connect to source vCenter Server
.PARAMETER sourceVCPassword
The password to connect to source vCenter Server
.PARAMETER destVC
The hostname or IP Address of the destination vCenter Server
.PARAMETER destVCUsername
The username to connect to the destination vCenter Server
.PARAMETER destVCPassword
The password to connect to the destination vCenter Server
.PARAMETER vmname
The name of the source VM to check for duplicated MAC Addresses
#>
param
(
   [Parameter(Mandatory=$true)]
   [string]
   $sourceVC,
   [Parameter(Mandatory=$true)]
   [string]
   $sourceVCUsername,
   [Parameter(Mandatory=$true)]
   [string]
   $sourceVCPassword,
   [Parameter(Mandatory=$true)]
   [string]
   $destVC,
   [Parameter(Mandatory=$true)]
   [string]
   $destVCUsername,
   [Parameter(Mandatory=$true)]
   [string]
   $destVCPassword,
   [Parameter(Mandatory=$true)]
   [string]
   $vmname
);

# Debug
#$sourceVC = "vcenter60-1.primp-industries.com"
#$sourceVCUsername = "administrator@vghetto.local"
#$sourceVCPassword = "VMware1!"
#
#$destVC = "vcenter60-2.primp-industries.com"
#$destVCUsername = "administrator@vghetto.local"
#$destVCPassword = "VMware1!"
#
#$vmname = "VM1"

# Connect to Source vCenter Server
$sourceVCConn = Connect-VIServer -Server $sourceVC -user $sourceVCUsername -password $sourceVCPassword

# Connect to Destination vCenter Server
$destVCConn = Connect-VIServer -Server $destVC -user $destVCUsername -password $destVCpassword

# Retrieve Source VM MAC Addresses
$sourceVMMACs = (Get-NetworkAdapter -Server $sourceVCConn -VM $vmname).MacAddress

# Retrieve ALL VM Mac Addresses from Destination vCenter Server
$allVMMacs = @{}
$vms = Get-View -Server $destVCConn -ViewType VirtualMachine -Property Name,Config.Hardware.Device -Filter @{"Config.Template" = "False"}
foreach ($vm in $vms) {
	$devices = $vm.Config.Hardware.Device
	foreach ($device in $devices) {
		if($device -is  [VMware.Vim.VirtualEthernetCard]) {
			# Store hash of Mac to VM to be used later for later comparison
			$allVMMacs.add($device.MacAddress,$vm.Name)
		}
	}
}

# Disconnect from Source/Dest vCenter Servers as it is no longer needed
Disconnect-VIServer -Server $sourceVCConn -Confirm:$false
Disconnect-VIServer -Server $destVCConn -Confirm:$false

# Check for duplicated MAC Addresses in destionation vCenter Server
Write-Host "`nChecking to see if there are MAC Address conflicts with" $vmname "at destination vCenter Server...`n"

foreach ($mac in $sourceVMMACs) {
	if($allVMMacs[$mac]) {
		Write-Host $allVMMacs[$mac] "also has MAC Address: $mac"
	}
}
# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script to enable MultiWriter Flag for existing VMDK in vSphere 6.x
# Reference: http://www.virtuallyghetto.com/2015/10/new-method-of-enabling-multiwriter-vmdk-flag-in-vsphere-6-0-update-1.html

$vcname = "192.168.1.150"
$vcuser = "administrator@vghetto.local"
$vcpass = "VMware1!"

$vmName = "vm-1"
$diskName = "Hard disk 2"

#### DO NOT EDIT BEYOND HERE ####

$server = Connect-VIServer -Server $vcname -User $vcuser -Password $vcpass

# Retrieve VM and only its Devices
$vm = Get-View -Server $server -ViewType VirtualMachine -Property Name,Config.Hardware.Device -Filter @{"Name" = $vmName}

# Array of Devices on VM
$vmDevices = $vm.Config.Hardware.Device

# Find the Virtual Disk that we care about
foreach ($device in $vmDevices) {
	if($device -is  [VMware.Vim.VirtualDisk] -and $device.deviceInfo.Label -eq $diskName) {
		$diskDevice = $device
		$diskDeviceBaking = $device.backing
		break
	}
}

# Create VM Config Spec to Edit existing VMDK & Enable Multi-Writer Flag
$spec = New-Object VMware.Vim.VirtualMachineConfigSpec
$spec.deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec
$spec.deviceChange[0].operation = 'edit'
$spec.deviceChange[0].device = New-Object VMware.Vim.VirtualDisk
$spec.deviceChange[0].device = $diskDevice
$spec.DeviceChange[0].device.backing = New-Object VMware.Vim.VirtualDiskFlatVer2BackingInfo
$spec.DeviceChange[0].device.backing = $diskDeviceBaking
$spec.DeviceChange[0].device.Backing.Sharing = "sharingMultiWriter"

Write-Host "`nEnabling Multiwriter flag on on VMDK:" $diskName "for VM:" $vmname
$task = $vm.ReconfigVM_Task($spec)
$task1 = Get-Task -Id ("Task-$($task.value)")
$task1 | Wait-Task

Disconnect-VIServer $server -Confirm:$false
# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script to configure per-VMDK IOPS Reservations on a VM in vSphere 6.0
# Reference: http://www.virtuallyghetto.com/2015/05/configuring-per-vmdk-iops-reservations-in-vsphere-6-0

$server = Connect-VIServer -Server 192.168.1.60 -User administrator@vghetto.local -Password VMware1!

# Fill out with your VM Name, Disk Label & IOPS Reservation
$vmName = "Photon"
$diskName = "Hard disk 1"
$iopsReservation = "2000"

### DO NOT EDIT BEYOND HERE ###

# Retrieve VM and only its Devices
$vm = Get-View -Server $server -ViewType VirtualMachine -Property Name,Config.Hardware.Device -Filter @{"Name" = $vmName}

# Array of Devices on VM
$vmDevices = $vm.Config.Hardware.Device

# Find the Virtual Disk that we care about
foreach ($device in $vmDevices) {
	if($device -is  [VMware.Vim.VirtualDisk] -and $device.deviceInfo.Label -eq $diskName) {
			$diskDevice = $device
			break
	}
}

# Create VM Config Spec
$spec = New-Object VMware.Vim.VirtualMachineConfigSpec
$spec.deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec
$spec.deviceChange[0].operation = 'edit'
$spec.deviceChange[0].device = New-Object VMware.Vim.VirtualDisk
$spec.deviceChange[0].device = $diskDevice
$spec.deviceChange[0].device.storageIOAllocation.reservation = $iopsReservation
Write-Host "Configuring IOPS Reservation:" $iopsReservation "on VMDK:" $diskName "for VM:" $vmname
$vm.ReconfigVM($spec)

# Uncomment the following snippet if you wish to verify as part of the reconfiguration operation 

#$vm.UpdateViewData()
#$vmDevices = $vm.Config.Hardware.Device
#foreach ($device in $vmDevices) {
#	if($device -is  [VMware.Vim.VirtualDisk] -and $device.deviceInfo.Label -eq $diskName) {
#			$device.storageIOAllocation
#	}
#}

Disconnect-VIServer $server -Confirm:$false$ActivationKey = "<FILL ME>"
$HCXServer = "mgmt-hcxm-02.cpbu.corp"
$VAMIUsername = "admin"
$VAMIPassword = "VMware1!"
$VIServer = "mgmt-vcsa-01.cpbu.corp"
$VIUsername = "administrator@vsphere.local"
$VIPassword = "VMware1!"
$NSXServer = "mgmt-nsxm-01.cpbu.corp"
$NSXUsername = "admin"
$NSXPassword = "VMware1!"

Connect-HcxVAMI -Server $HCXServer -Username $VAMIUsername -Password $VAMIPassword

Set-HcxLicense -LicenseKey $ActivationKey

Set-HcxVCConfig -VIServer $VIServer -VIUsername $VIUsername -VIPassword $VIPassword -PSCServer $VIServer

Set-HcxNSXConfig -NSXServer $NSXServer -NSXUsername $NSXUsername -NSXPassword $NSXPassword

Set-HcxLocation -City "Santa Barbara" -Country "United States of America"

Set-HcxRoleMapping -SystemAdminGroup @("vsphere.local\Administrators","cpbu.corp\Administrators") -EnterpriseAdminGroup @("vsphere.local\Administrators","cpbu.corp\Administrators")
# William Lam
# www.virtuallyghetto.com

$vcname = "192.168.1.150"
$vcuser = "administrator@vghetto.local"
$vcpass = "VMware1!"

$ovffile = "Z:\Desktop\Nested_ESXi_Appliance.ovf"

$cluster = "MacMini-Cluster"
$vmnetwork = "VM Network"
$datastore = "mini-local-datastore-1"
$iprange = "192.168.1"
$netmask = "255.255.255.0"
$gateway = "192.168.1.1"
$dns = "192.168.1.1"
$dnsdomain = "primp-industries.com"
$ntp = "192.168.1.1"
$syslog = "192.168.1.150"
$password = "VMware1!"
$ssh = "True"

#### DO NOT EDIT BEYOND HERE ####

$vcenter = Connect-VIServer $vcname -User $vcuser -Password $vcpass -WarningAction SilentlyContinue

$datastore_ref = Get-Datastore -Name $datastore
$network_ref = Get-VirtualPortGroup -Name $vmnetwork
$cluster_ref = Get-Cluster -Name $cluster
$vmhost_ref = $cluster_ref | Get-VMHost | Select -First 1

$ovfconfig = Get-OvfConfiguration $ovffile
$ovfconfig.NetworkMapping.VM_Network.value = $network_ref

190..192 | Foreach {
    $ipaddress = "$iprange.$_"
    # Try to perform DNS lookup
    try {
        $vmname = ([System.Net.Dns]::GetHostEntry($ipaddress).HostName).split(".")[0]
    }
    Catch [system.exception]
    {
        $vmname = "vesxi-vsan-$ipaddress"
    }
    $ovfconfig.common.guestinfo.hostname.value = $vmname
    $ovfconfig.common.guestinfo.ipaddress.value = $ipaddress
    $ovfconfig.common.guestinfo.netmask.value = $netmask
    $ovfconfig.common.guestinfo.gateway.value = $gateway
    $ovfconfig.common.guestinfo.dns.value = $dns
    $ovfconfig.common.guestinfo.domain.value = $dnsdomain
    $ovfconfig.common.guestinfo.ntp.value = $ntp
    $ovfconfig.common.guestinfo.syslog.value = $syslog
    $ovfconfig.common.guestinfo.password.value = $password
    $ovfconfig.common.guestinfo.ssh.value = $ssh

    # Deploy the OVF/OVA with the config parameters
    Write-Host "Deploying $vmname ..."
    $vm = Import-VApp -Source $ovffile -OvfConfiguration $ovfconfig -Name $vmname -Location $cluster_ref -VMHost $vmhost_ref -Datastore $datastore_ref -DiskStorageFormat thin
    $vm | Start-Vm -RunAsync | Out-Null
}

Disconnect-VIServer $vcenter -Confirm:$false
﻿# Load OVF/OVA configuration into a variable
$ovffile = "C:\Users\william\Desktop\VMware-HCX-Enterprise-3.5.1-10027070.ova"
$ovfconfig = Get-OvfConfiguration $ovffile

# vSphere Cluster + VM Network configurations
$Cluster = "Cluster-01"
$VMName = "MGMT-HCXM-02"
$VMNetwork = "SJC-CORP-MGMT-EP"
$HCXAddressToVerify = "mgmt-hcxm-02.cpbu.corp"

$VMHost = Get-Cluster $Cluster | Get-VMHost | Sort MemoryGB | Select -first 1
$Datastore = $VMHost | Get-datastore | Sort FreeSpaceGB -Descending | Select -first 1
$Network = Get-VDPortGroup -Name $VMNetwork

# Fill out the OVF/OVA configuration parameters

# vSphere Portgroup Network Mapping
$ovfconfig.NetworkMapping.VSMgmt.value = $Network

# IP Address
$ovfConfig.common.mgr_ip_0.value = "172.17.31.50"

# Netmask
$ovfConfig.common.mgr_prefix_ip_0.value = "24"

# Gateway
$ovfConfig.common.mgr_gateway_0.value = "172.17.31.253"

# DNS Server
$ovfConfig.common.mgr_dns_list.value = "172.17.31.5"

# DNS Domain
$ovfConfig.common.mgr_domain_search_list.value  = "cpbu.corp"

# Hostname
$ovfconfig.Common.hostname.Value = "mgmt-hcxm-02.cpbu.corp"

# NTP
$ovfconfig.Common.mgr_ntp_list.Value = "172.17.31.5"

# SSH
$ovfconfig.Common.mgr_isSSHEnabled.Value = $true

# Password
$ovfconfig.Common.mgr_cli_passwd.Value = "VMware1!"
$ovfconfig.Common.mgr_root_passwd.Value = "VMware1!"

# Deploy the OVF/OVA with the config parameters
Write-Host -ForegroundColor Green "Deploying HCX Manager OVA ..."
$vm = Import-VApp -Source $ovffile -OvfConfiguration $ovfconfig -Name $VMName -VMHost $vmhost -Datastore $datastore -DiskStorageFormat thin

# Power On the HCX Manager VM after deployment
Write-Host -ForegroundColor Green "Powering on HCX Manager ..."
$vm | Start-VM -Confirm:$false | Out-Null

# Waiting for HCX Manager to initialize
while(1) {
    try {
        if($PSVersionTable.PSEdition -eq "Core") {
            $requests = Invoke-WebRequest -Uri "https://$($HCXAddressToVerify):9443" -Method GET -SkipCertificateCheck -TimeoutSec 5
        } else {
            $requests = Invoke-WebRequest -Uri "https://$($HCXAddressToVerify):9443" -Method GET -TimeoutSec 5
        }
        if($requests.StatusCode -eq 200) {
            Write-Host -ForegroundColor Green "HCX Manager is now ready to be configured!"
            break
        }
    }
    catch {
        Write-Host -ForegroundColor Yellow "HCX Manager is not ready yet, sleeping for 120 seconds ..."
        sleep 120
    }
}# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script to disable/enable vMotion capability for a specific VM
# Reference: http://www.virtuallyghetto.com/2016/07/how-to-easily-disable-vmotion-cross-vcenter-vmotion-for-a-particular-virtual-machine.html

Function Enable-vSphereMethod {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [String]$vmmoref,
    [string]$vc_server,
    [String]$vc_username,
    [String]$vc_password,
    [String]$enable_method
    )

    $secpasswd = ConvertTo-SecureString $vc_password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($vc_username, $secpasswd)

    # vSphere MOB URL to private enableMethods
    $mob_url = "https://$vc_server/mob/?moid=AuthorizationManager&method=enableMethods"

# Ingore SSL Warnings
add-type -TypeDefinition  @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

    # Initial login to vSphere MOB using GET and store session using $vmware variable
    $results = Invoke-WebRequest -Uri $mob_url -SessionVariable vmware -Credential $credential -Method GET

    # Extract hidden vmware-session-nonce which must be included in future requests to prevent CSRF error
    # Credit to https://blog.netnerds.net/2013/07/use-powershell-to-keep-a-cookiejar-and-post-to-a-web-form/ for parsing vmware-session-nonce via Powershell
    if($results.StatusCode -eq 200) {
        $null = $results -match 'name="vmware-session-nonce" type="hidden" value="?([^\s^"]+)"'
        $sessionnonce = $matches[1]
    } else {
        Write-host "Failed to login to vSphere MOB"
        exit 1
    }

    # The POST data payload must include the vmware-session-nonce variable + URL-encoded
    $body = @"
vmware-session-nonce=$sessionnonce&entity=%3Centity+type%3D%22ManagedEntity%22+xsi%3Atype%3D%22ManagedObjectReference%22%3E$vmmoref%3C%2Fentity%3E%0D%0A&method=%3Cmethod%3E$enable_method%3C%2Fmethod%3E
"@

    # Second request using a POST and specifying our session from initial login + body request
    $results = Invoke-WebRequest -Uri $mob_url -WebSession $vmware -Method POST -Body $body

    # Logout out of vSphere MOB
    $mob_logout_url = "https://$vc_server/mob/logout"
    Invoke-WebRequest -Uri $mob_logout_url -WebSession $vmware -Method GET    
}

Function Disable-vSphereMethod {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [String]$vmmoref,
    [string]$vc_server,
    [String]$vc_username,
    [String]$vc_password,
    [String]$disable_method
    )

    $secpasswd = ConvertTo-SecureString $vc_password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($vc_username, $secpasswd)

    # vSphere MOB URL to private disableMethods
    $mob_url = "https://$vc_server/mob/?moid=AuthorizationManager&method=disableMethods"

# Ingore SSL Warnings
add-type -TypeDefinition  @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

    # Initial login to vSphere MOB using GET and store session using $vmware variable
    $results = Invoke-WebRequest -Uri $mob_url -SessionVariable vmware -Credential $credential -Method GET

    # Extract hidden vmware-session-nonce which must be included in future requests to prevent CSRF error
    # Credit to https://blog.netnerds.net/2013/07/use-powershell-to-keep-a-cookiejar-and-post-to-a-web-form/ for parsing vmware-session-nonce via Powershell
    if($results.StatusCode -eq 200) {
        $null = $results -match 'name="vmware-session-nonce" type="hidden" value="?([^\s^"]+)"'
        $sessionnonce = $matches[1]
    } else {
        Write-host "Failed to login to vSphere MOB"
        exit 1
    }

    # The POST data payload must include the vmware-session-nonce variable + URL-encoded
    $body = @"
vmware-session-nonce=$sessionnonce&entity=%3Centity+type%3D%22ManagedEntity%22+xsi%3Atype%3D%22ManagedObjectReference%22%3E$vmmoref%3C%2Fentity%3E%0D%0A%0D%0A&method=%3CDisabledMethodRequest%3E%0D%0A+++%3Cmethod%3E$disable_method%3C%2Fmethod%3E%0D%0A%3C%2FDisabledMethodRequest%3E%0D%0A%0D%0A&sourceId=self
"@

    # Second request using a POST and specifying our session from initial login + body request
    $results = Invoke-WebRequest -Uri $mob_url -WebSession $vmware -Method POST -Body $body
}

### Sample Usage of Enable/Disable functions ###

$vc_server = "192.168.1.51"
$vc_username = "administrator@vghetto.local"
$vc_password = "VMware1!"
$vm_name = "TestVM-1"
$method_name = "MigrateVM_Task"

# Connect to vCenter Server
$server = Connect-VIServer -Server $vc_server -User $vc_username -Password $vc_password

$vm = Get-VM -Name $vm_name
$vm_moref = (Get-View $vm).MoRef.Value

#Disable-vSphereMethod -vc_server $vc_server -vc_username $vc_username -vc_password $vc_password -vmmoref $vm_moref -disable_method $method_name

#Enable-vSphereMethod -vc_server $vc_server -vc_username $vc_username -vc_password $vc_password -vmmoref $vm_moref -enable_method $method_name

# Disconnect from vCenter Server
Disconnect-viserver $server -confirm:$false
﻿Function Get-ESXiBootDevice {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function identifies how an ESXi host was booted up along with its boot
        device (if applicable). This supports both local installation to Auto Deploy as
        well as Boot from SAN.
    .PARAMETER VMHostname
        The name of an individual ESXi host managed by vCenter Server
    .EXAMPLE
        Get-ESXiBootDevice
    .EXAMPLE
        Get-ESXiBootDevice -VMHostname esxi-01
#>
    param(
        [Parameter(Mandatory=$false)][String]$VMHostname
    )

    if($VMHostname) {
        $vmhosts = Get-VMhost -Name $VMHostname
    } else {
        $vmhosts = Get-VMHost
    }

    $results = @()
    foreach ($vmhost in ($vmhosts | Sort-Object -Property Name)) {
        $esxcli = Get-EsxCli -V2 -VMHost $vmhost
        $bootDetails = $esxcli.system.boot.device.get.Invoke()

        # Check to see if ESXi booted over the network
        $networkBoot = $false
        if($bootDetails.BootNIC) {
            $networkBoot = $true
            $bootDevice = $bootDetails.BootNIC
        } elseif ($bootDetails.StatelessBootNIC) {
            $networkBoot = $true
            $bootDevice = $bootDetails.StatelessBootNIC
        }

        # If ESXi booted over network, check to see if deployment
        # is Stateless, Stateless w/Caching or Stateful
        if($networkBoot) {
            $option = $esxcli.system.settings.advanced.list.CreateArgs()
            $option.option = "/UserVars/ImageCachedSystem"
            try {
                $optionValue = $esxcli.system.settings.advanced.list.Invoke($option)
            } catch {
                $bootType = "stateless"
            }
            $bootType = $optionValue.StringValue
        }

        # Loop through all storage devices to identify boot device
        $devices = $esxcli.storage.core.device.list.Invoke()
        $foundBootDevice = $false
        foreach ($device in $devices) {
            if($device.IsBootDevice -eq $true) {
                $foundBootDevice = $true

                if($device.IsLocal -eq $true -and $networkBoot -and $bootType -ne "stateful") {
                    $bootType = "stateless caching"
                } elseif($device.IsLocal -eq $true -and $networkBoot -eq $false) {
                    $bootType = "local"
                } elseif($device.IsLocal -eq $false -and $networkBoot -eq $false) {
                    $bootType = "remote"
                }

                $bootDevice = $device.Device
                $bootModel = $device.Model
                $bootVendor = $device.VEndor
                $bootSize = $device.Size
                $bootIsSAS = $device.IsSAS
                $bootIsSSD = $device.IsSSD
                $bootIsUSB = $device.IsUSB
            }
        }

        # Pure Stateless (e.g. No USB or Disk for boot)
        if($networkBoot-and $foundBootDevice -eq $false) {
            $bootModel = "N/A"
            $bootVendor = "N/A"
            $bootSize = "N/A"
            $bootIsSAS = "N/A"
            $bootIsSSD = "N/A"
            $bootIsUSB = "N/A"
        }

        $tmp = [pscustomobject] @{
            Host = $vmhost.Name;
            Device = $bootDevice;
            BootType = $bootType;
            Vendor = $bootVendor;
            Model = $bootModel;
            SizeMB = $bootSize;
            IsSAS = $bootIsSAS;
            IsSSD = $bootIsSSD;
            IsUSB = $bootIsUSB;
        }
        $results+=$tmp
    }
    $results | FT -AutoSize
}﻿Function Get-ESXiDPC {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retreives the current disabled TLS protocols for all ESXi
        hosts within a vSphere Cluster
    .SYNOPSIS
        Returns current disabled TLS protocols for Hostd, Authd, sfcbd & VSANVP/IOFilter 
    .PARAMETER Cluster
        The name of the vSphere Cluster
    .EXAMPLE
        Get-ESXiDPC -Cluster VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$Cluster
    )

    $debug = $false

    Function Get-SFCBDConf {
        param(
            [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$VMHost
        )

        $url = "https://$vmhost/host/sfcb.cfg"

        $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

        $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
        $spec.Method = "httpGet"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie = New-Object System.Net.Cookie
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhost.name
        $websession.Cookies.Add($cookie)
        $result = Invoke-WebRequest -Uri $url -WebSession $websession
        $sfcbConf = $result.content
        
        # Extract the TLS fields if they exists
        $sfcbResults = @()
        $usingDefault = $true
        foreach ($line in $sfcbConf.Split("`n")) {
            if($line -match "enableTLSv1:") {
                ($key,$value) = $line.Split(":")
                if($value -match "false") {
                    $sfcbResults+="tlsv1"
                }
                $usingDefault = $false
            }
            if($line -match "enableTLSv1_1:") {
                ($key,$value) = $line.Split(":")
                if($value -match "false") {
                    $sfcbResults+="tlsv1.1"
                }
                $usingDefault = $false
            }
            if($line -match "enableTLSv1_2:") {
                ($key,$value) = $line.Split(":")
                if($value -match "false") {
                    $sfcbResults+="tlsv1.2"
                }
                $usingDefault = $false
            }
        }
        if($usingDefault -or ($sfcbResults.Length -eq 0)) {
            $sfcbResults = "tlsv1,tlsv1.1,sslv3"
            return $sfcbResults
        } else {
            $sfcbResults+="sslv3"
            return $sfcbResults -join ","
        }
    }

    $results = @()
    foreach ($vmhost in (Get-Cluster -Name $Cluster | Get-VMHost)) {
        if( ($vmhost.ApiVersion -eq "6.0" -and (Get-AdvancedSetting -Entity $vmhost -Name "Misc.HostAgentUpdateLevel").value -eq "3") -or ($vmhost.ApiVersion -eq "6.5") ) {
            $esxiVersion = ($vmhost.ApiVersion) + " Update " + (Get-AdvancedSetting -Entity $vmhost -Name "Misc.HostAgentUpdateLevel").value
            
            $vps = (Get-AdvancedSetting -Entity $vmhost -Name "UserVars.ESXiVPsDisabledProtocols" -ErrorAction SilentlyContinue).value
            # ESXi 6.5 - UserVars.ESXiVPsDisabledProtocols covers both VPs+rHTTP
            if($vmhost.ApiVersion -eq "6.5") {
                $rhttpProxy = $vps
                # Only TLS 1.2 is enabled 
                $vmauth = "tlsv1,tlsv1.1,sslv3"
            } else {
                $rhttpProxy = (Get-AdvancedSetting -Entity $vmhost -Name "UserVars.ESXiRhttpproxyDisabledProtocols" -ErrorAction SilentlyContinue).value
                $vmauth = (Get-AdvancedSetting -Entity $vmhost -Name "UserVars.VMAuthdDisabledProtocols" -ErrorAction SilentlyContinue).value
            }
            $sfcbd = Get-SFCBDConf -vmhost $vmhost

            $hostTLSSettings = [pscustomobject] @{
                vmhost = $vmhost.name;
                version = $esxiVersion;
                hostd = $rhttpProxy;
                authd = $vmauth;
                sfcbd = $sfcbd
                ioFilterVSANVP = $vps
            }
            $results+=$hostTLSSettings
        }
    }
    Write-Host -NoNewline "`nDisabled Protocols on all ESXi hosts:"
    $results
}

Function Set-ESXiDPC {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function configures the TLS protocols to disable for all 
        ESXi hosts within a vSphere Cluster
    .SYNOPSIS
        Configures the disabled TLS protocols for Hostd, Authd, sfcbd & VSANVP/IOFilter 
    .PARAMETER Cluster
        The name of the vSphere Cluster
    .EXAMPLE
        Set-ESXiDPC -Cluster VSAN-Cluster -TLS1 $true -TLS1_1 $true -TLS1_2 $false -SSLV3 $true
#>
    param(
        [Parameter(Mandatory=$true)][String]$Cluster,
        [Parameter(Mandatory=$true)][Boolean]$TLS1,
        [Parameter(Mandatory=$true)][Boolean]$TLS1_1,
        [Parameter(Mandatory=$true)][Boolean]$TLS1_2,
        [Parameter(Mandatory=$true)][Boolean]$SSLV3
    )

    Function UpdateSFCBConfig {
        param(
            [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$VMHost
        )

        $url = "https://$vmhost/host/sfcb.cfg"

        $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

        $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
        $spec.Method = "httpGet"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie = New-Object System.Net.Cookie
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhost.name
        $websession.Cookies.Add($cookie)
        $result = Invoke-WebRequest -Uri $url -WebSession $websession
        $sfcbConf = $result.content
        
        # Download the current sfcb.cfg and ignore existing TLS configuration
        $sfcbResults = ""
        foreach ($line in $sfcbConf.Split("`n")) {
            if($line -notmatch "enableTLSv1:" -and $line -notmatch "enableTLSv1_1:" -and $line -notmatch "enableTLSv1_2:" -and $line -ne "") {
                $sfcbResults+="$line`n"
            }
        }
        
        # Append the TLS protocols based on user input to the configuration file
        $sfcbResults+="enableTLSv1: " + (!$TLS1).ToString().ToLower() + "`n"
        $sfcbResults+="enableTLSv1_1: " + (!$TLS1_1).ToString().ToLower() + "`n"
        $sfcbResults+="enableTLSv1_2: " + (!$TLS1_2).ToString().ToLower() +"`n"

        # Create HTTP PUT spec
        $spec.Method = "httpPut"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

        # Upload sfcb.cfg back to ESXi host
        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhost.name
        $websession.Cookies.Add($cookie)
        $result = Invoke-WebRequest -Uri $url -WebSession $websession -Body $sfcbResults -Method Put -ContentType "plain/text"
        if($result.StatusCode -eq 200) {
            Write-Host "`tSuccessfully updated sfcb.cfg file"
        } else {
            Write-Host "Failed to upload sfcb.cfg file"
            break
        }
    }

    # Build TLS string based on user input for setting ESXi Advanced Settings
    if($TLS1 -and $TLS1_1 -and $TLS1_2 -and $SSLV3) {
        Write-Host -ForegroundColor Red "Error: You must at least enable one of the TLS protocols"
        break
    }

    $tlsString = @()
    if($TLS1) { $tlsString += "tlsv1" }
    if($TLS1_1) { $tlsString += "tlsv1.1" }
    if($TLS1_2) { $tlsString += "tlsv1.2" }
    if($SSLV3) { $tlsString += "sslv3" }
    $tlsString = $tlsString -join ","

    Write-Host "`nDisabling the following TLS protocols: $tlsString on ESXi hosts ...`n"
    foreach ($vmhost in (Get-Cluster -Name $Cluster | Get-VMHost)) {
        if( ($vmhost.ApiVersion -eq "6.0" -and (Get-AdvancedSetting -Entity $vmhost -Name "Misc.HostAgentUpdateLevel").value -eq "3") -or ($vmhost.ApiVersion -eq "6.5") ) {
            Write-Host "Updating $vmhost ..."

            Write-Host "`tUpdating sfcb.cfg ..."
            UpdateSFCBConfig -vmhost $vmhost

            if($vmhost.ApiVersion -eq "6.0") {
                Write-Host "`tUpdating UserVars.ESXiRhttpproxyDisabledProtocols ..."
                Get-AdvancedSetting -Entity $vmhost -Name "UserVars.ESXiRhttpproxyDisabledProtocols" | Set-AdvancedSetting -Value $tlsString -Confirm:$false | Out-Null

                Write-Host "`tUpdating UserVars.VMAuthdDisabledProtocols ..."
                Get-AdvancedSetting -Entity $vmhost -Name "UserVars.VMAuthdDisabledProtocols" | Set-AdvancedSetting -Value $tlsString -Confirm:$false | Out-Null
            }
            Write-Host "`tUpdating UserVars.ESXiVPsDisabledProtocols ..."
            Get-AdvancedSetting -Entity $vmhost -Name "UserVars.ESXiVPsDisabledProtocols" | Set-AdvancedSetting -Value $tlsString -Confirm:$false | Out-Null
        }
    }
}
﻿<#
.SYNOPSIS Retrieve the installation date of an ESXi host
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/10/super-easy-way-of-getting-esxi-installation-date-in-vsphere-6-5.html
.PARAMETER Vmhost
  ESXi host to query installation date
.EXAMPLE
  Get-Vmhost "mini" | Get-ESXInstallDate
#>

Function Get-ESXInstallDate {
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$Vmhost
     )

     if($Vmhost.Version -eq "6.5.0") {
        $imageManager = Get-View ($Vmhost.ExtensionData.ConfigManager.ImageConfigManager)
        $installDate = $imageManager.installDate()

        Write-Host "$Vmhost was installed on $installDate"
     } else {
        Write-Host "ESXi must be running 6.5"
     }
}
﻿<#
.SYNOPSIS Retrieve the installation date of an ESXi host
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.PARAMETER Vmhost
  ESXi host to query installed ESXi VIBs
.EXAMPLE
  Get-ESXInstalledVibs -Vmhost (Get-Vmhost "mini")
.EXAMPLE
  Get-ESXInstalledVibs -Vmhost (Get-Vmhost "mini") -vibname vsan
#>

Function Get-ESXInstalledVibs {
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$Vmhost,
        [Parameter(
            Mandatory=$false,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [String]$vibname=""
     )

     $imageManager = Get-View ($Vmhost.ExtensionData.ConfigManager.ImageConfigManager)
     $vibs = $imageManager.fetchSoftwarePackages()

     foreach ($vib in $vibs) {
        if($vibname -ne "") {
            if($vib.name -eq $vibname) {
                return $vib | select Name, Version, Vendor, CreationDate, Summary
            }
        } else {
            $vib | select Name, Version, Vendor, CreationDate, Summary
        }
     }
}
﻿$esxiVersions = @("5.1.0", "5.5.0", "6.0.0", "6.5.0", "6.7.0")
$pathToStoreMetdataFile = $env:TMP

Add-Type -Assembly System.IO.Compression.FileSystem

Write-Host "Downloading ESXi Metadata Files ..."
foreach ($esxiVersion in $esxiVersions) {
    $metadataUrl = "https://hostupdate.vmware.com/software/VUM/PRODUCTION/main/esx/vmw/vmw-ESXi-$esxiVersion-metadata.zip"
    $metadataDownloadPath = $pathToStoreMetdataFile + "\" + $esxiVersion + ".zip"
    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile($metadataUrl,$metadataDownloadPath)

    #https://stackoverflow.com/a/41575369
    $zip = [IO.Compression.ZipFile]::OpenRead($metadataDownloadPath)
    $metadataFileExtractionPath = $pathToStoreMetdataFile + "\$esxiVersion.xml"
    $zip.Entries | where {$_.Name -like 'vmware.xml'} | foreach {[System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, $metadataFileExtractionPath, $true)}
    $zip.Dispose()
    Remove-Item -Path $metadataDownloadPath -Force
}

Write-Host "Processing ESXi Metadata Files ..."
$esxiBulletinCVEesults = @()
foreach ($esxiVersion in $esxiVersions) {
    $metadataFileExtractionPath = $pathToStoreMetdataFile + "\$esxiVersion.xml"
    [xml]$XmlDocument = Get-Content -Path $metadataFileExtractionPath

    Write-Host "Extracting KB Information & CVE URLs for $esxiVersion ..." 
    foreach ($bulletin in $XmlDocument.metadataResponse.bulletin) {
        if($bulletin.category -eq "security") {
            $bulletinId = $bulletin.id
            $kbId = ($bulletin.kbUrl).Replace("http://kb.vmware.com/kb/","")

            $results = Invoke-WebRequest -Uri https://kb.vmware.com/articleview?docid=$kbId -UseBasicParsing

            $cveIds = @()
            foreach ($link in $results.Links) {
                if($link.href -match "CVE") {
                    $cveIds += ($link.href).Replace("http://cve.mitre.org/cgi-bin/cvename.cgi?name=","")
                }
            }

            if($cveIds) {
                foreach ($cveId in $cveIds) {
                    # CVE API to retrieve CVE details
                    $results = Invoke-WebRequest -Uri  http://cve.circl.lu/api/cve/$cveId -UseBasicParsing
                    $jsonResults = $results.Content | ConvertFrom-Json
                    $cvssScore = $jsonResults.cvss
                    $cvssComplexity = $jsonResults.access.complexity

                    if($cvssScore -eq $null) {
                        $cvssScore = "N/A"
                    }
                    if($cvssComplexity -eq $null) {
                        $cvssComplexity = "N/A"
                    }

                    $tmp = [PSCustomObject] @{
                        Bulletin = $bulletinId;
                        CVEId = $cveId;
                        CVSSScore = $cvssScore;
                        CVSSComplexity = $cvssComplexity;
                    }
                    $esxiBulletinCVEesults += $tmp
                }
            }
        }
    }
}

$esxiBulletinCVEesults﻿Function Get-VMConsoleURL {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function generates the HTML5 VM Console URL (default) as seen when using the
        vSphere Web Client connected to a vCenter Server 6.5 environment. You must
        already be logged in for the URL to be valid or else you will be prompted
        to login before being re-directed to VM Console. You have option of also generating
        either the Standalone VMRC or WebMKS URL using additional parameters
    .PARAMETER VMName
        The name of the VM
    .PARAMETER webmksUrl
        Set to true to generate the WebMKS URL instead (e.g. wss://<host>/ticket/<ticket>)
    .PARAMETER vmrcUrl
        Set to true to generate the VMRC URL instead (e.g. vmrc://...)
    .EXAMPLE
        Get-VMConsoleURL -VMName "Embedded-VCSA1"
    .EXAMPLE
        Get-VMConsoleURL -VMName "Embedded-VCSA1" -vmrcUrl $true
    .EXAMPLE
        Get-VMConsoleURL -VMName "Embedded-VCSA1" -webmksUrl $true
        #>
    param(
        [Parameter(Mandatory=$true)][String]$VMName,
        [Parameter(Mandatory=$false)][Boolean]$vmrcUrl,
        [Parameter(Mandatory=$false)][Boolean]$webmksUrl
    )

    Function Get-SSLThumbprint {
        param(
        [Parameter(
            Position=0,
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [Alias('FullName')]
        [String]$URL
        )

        add-type @"
            using System.Net;
            using System.Security.Cryptography.X509Certificates;

                public class IDontCarePolicy : ICertificatePolicy {
                public IDontCarePolicy() {}
                public bool CheckValidationResult(
                    ServicePoint sPoint, X509Certificate cert,
                    WebRequest wRequest, int certProb) {
                    return true;
                }
                }
"@
        [System.Net.ServicePointManager]::CertificatePolicy = new-object IDontCarePolicy

        # Need to connect using simple GET operation for this to work
        Invoke-RestMethod -Uri $URL -Method Get | Out-Null

        $ENDPOINT_REQUEST = [System.Net.Webrequest]::Create("$URL")
        $SSL_THUMBPRINT = $ENDPOINT_REQUEST.ServicePoint.Certificate.GetCertHashString()

        return $SSL_THUMBPRINT -replace '(..(?!$))','$1:'
    }

    $VM = Get-VM -Name $VMName
    $VMMoref = $VM.ExtensionData.MoRef.Value

    if($webmksUrl) {
        $WebMKSTicket = $VM.ExtensionData.AcquireTicket("webmks")
        $VMHostName = $WebMKSTicket.host
        $Ticket = $WebMKSTicket.Ticket
        $URL = "wss://$VMHostName`:443/ticket/$Ticket"
    } elseif($vmrcUrl) {
        $VCName = $global:DefaultVIServer.Name
        $SessionMgr = Get-View $DefaultViserver.ExtensionData.Content.SessionManager
        $Ticket = $SessionMgr.AcquireCloneTicket()
        $URL = "vmrc://clone`:$Ticket@$VCName`:443/?moid=$VMMoref"
    } else {
        $VCInstasnceUUID = $global:DefaultVIServer.InstanceUuid
        $VCName = $global:DefaultVIServer.Name
        $SessionMgr = Get-View $DefaultViserver.ExtensionData.Content.SessionManager
        $Ticket = $SessionMgr.AcquireCloneTicket()
        $VCSSLThumbprint = Get-SSLThumbprint "https://$VCname"
        $URL = "https://$VCName`:9443/vsphere-client/webconsole.html?vmId=$VMMoref&vmName=$VMname&serverGuid=$VCInstasnceUUID&locale=en_US&host=$VCName`:443&sessionTicket=$Ticket&thumbprint=$VCSSLThumbprint”
    }
    $URL
}
﻿<#
.SYNOPSIS Remoting collecting esxcfg-info from an ESXi host using vCenter Server
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/06/using-the-vsphere-api-to-remotely-collect-esxi-esxcfg-info.html
.PARAMETER Vmhost
  ESXi host
.EXAMPLE
  PS> Get-VMHost -Name "esxi-1" | Get-Esxcfginfo
#>

Function Get-Esxcfginfo {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl[]]$VMHost
    )

    $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

    # URL to the ESXi esxcfg-info info
    $url = "https://" + $vmhost.Name + "/cgi-bin/esxcfg-info.cgi?xml"

    $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
    $spec.Method = "httpGet"
    $spec.Url = $url
    $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

    # Append the cookie generated from VC
    $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $cookie = New-Object System.Net.Cookie
    $cookie.Name = "vmware_cgi_ticket"
    $cookie.Value = $ticket.id
    $cookie.Domain = $vmhost.name
    $websession.Cookies.Add($cookie)

    # Retrieve file
    $result = Invoke-WebRequest -Uri $url -WebSession $websession -ContentType "application/xml"
    
    # cast output as an XML object
    return [ xml]$result.content
}

Connect-VIServer -Server 192.168.1.51 -User administrator@vghetto.local -password VMware1! | Out-Null

$xmlResult = Get-VMHost -Name "192.168.1.190" | Get-Esxcfginfo

# Extracting device-name, vendor-name & vendor-id as an example
foreach ($childnodes in $xmlResult.host.'hardware-info'.'pci-info'.'all-pci-devices'.'pci-device') {
   foreach ($childnode in $childnodes | select -ExpandProperty childnodes) {
    if($childnode.name -eq 'device-name') {
        $deviceName = $childnode.'#text'
    } elseif($childnode.name -eq 'vendor-name') {
        $vendorName = $childnode.'#text'
    } elseif($childnode.name -eq 'vendor-id') {
        $vendorId = $childnode.'#text'
    }
   }
   $deviceName
   $vendorName
   $vendorId
   Write-Host ""
}

Disconnect-VIServer * -Confirm:$false
﻿<#
.SYNOPSIS Remoting collecting ESXi configuration files using vCenter Server
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/06/using-the-vsphere-api-to-remotely-collect-esxi-configuration-files.html
.PARAMETER Vmhost
  ESXi host
.EXAMPLE
  PS> Get-VMHost -Name "esxi-1" | Get-Esxconf
#>

Function Get-Esxconf {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl[]]$VMHost
    )

    $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

    # URL to ESXi's esx.conf configuration file (can use any that show up under https://esxi_ip/host)
    $url = "https://192.168.1.190/host/esx.conf"

    # URL to the ESXi configuration file
    $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
    $spec.Method = "httpGet"
    $spec.Url = $url
    $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

    # Append the cookie generated from VC
    $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $cookie = New-Object System.Net.Cookie
    $cookie.Name = "vmware_cgi_ticket"
    $cookie.Value = $ticket.id
    $cookie.Domain = $vmhost.name
    $websession.Cookies.Add($cookie)

    # Retrieve file
    $result = Invoke-WebRequest -Uri $url -WebSession $websession
    return $result.content
}

Connect-VIServer -Server 192.168.1.51 -User administrator@vghetto.local -password VMware1! | Out-Null

$esxConf = Get-VMHost -Name "192.168.1.190" | Get-Esxconf

$esxConf

Disconnect-VIServer * -Confirm:$false﻿<#
.SYNOPSIS Using the vSphere API in vCenter Server to collect ESXTOP & vscsiStats metrics
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2017/02/using-the-vsphere-api-in-vcenter-server-to-collect-esxtop-vscsistats-metrics.html
.PARAMETER Vmhost
  ESXi host
.EXAMPLE
  PS> Get-VMHost -Name "esxi-1" | Get-EsxtopAPI
#>

Function Get-EsxtopAPI {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$VMHost
    )

    $serviceManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.serviceManager) -property "" -ErrorAction SilentlyContinue

    $locationString = "vmware.host." + $VMHost.Name
    $services = $serviceManager.QueryServiceList($null,$locationString)
    foreach ($service in $services) {
        if($service.serviceName -eq "Esxtop") {
            $serviceView = Get-View $services.Service -Property "entity"
            $serviceView.ExecuteSimpleCommand("CounterInfo")
            break
        }
    }
}

Connect-VIServer -Server 192.168.1.51 -User administrator@vsphere.local -password VMware1! | Out-Null

Get-VMHost -Name "192.168.1.50" | Get-EsxtopAPI

Disconnect-VIServer * -Confirm:$false# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script querying remote ESXi host without adding to vCenter Server
# Reference: http://www.virtuallyghetto.com/2016/07/remotely-query-an-esxi-host-without-adding-to-vcenter-server.html

Function Get-RemoteESXi {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [String]$hostname,
    [string]$username,
    [String]$password,
    [string]$port = 443
    )

    # Function to retrieve SSL Thumbprint of a host
    # https://gist.github.com/lamw/988e4599c0f88d9fc25c9f2af8b72c92
    Function Get-SSLThumbprint {
        param(
        [Parameter(
            Position=0,
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [Alias('FullName')]
        [String]$URL
        )

    add-type @"
            using System.Net;
            using System.Security.Cryptography.X509Certificates;
                public class IDontCarePolicy : ICertificatePolicy {
                public IDontCarePolicy() {}
                public bool CheckValidationResult(
                    ServicePoint sPoint, X509Certificate cert,
                    WebRequest wRequest, int certProb) {
                    return true;
                }
            }
"@
        [System.Net.ServicePointManager]::CertificatePolicy = new-object IDontCarePolicy

        # Need to connect using simple GET operation for this to work
        Invoke-RestMethod -Uri $URL -Method Get | Out-Null

        $ENDPOINT_REQUEST = [System.Net.Webrequest]::Create("$URL")
        $SSL_THUMBPRINT = $ENDPOINT_REQUEST.ServicePoint.Certificate.GetCertHashString()

        return $SSL_THUMBPRINT -replace '(..(?!$))','$1:'
    }

    # Host Connection Spec
    $spec = New-Object VMware.Vim.HostConnectSpec
    $spec.Force = $False
    $spec.HostName = $hostname
    $spec.UserName = $username
    $spec.Password = $password
    $spec.Port = $port
    # Retrieve the SSL Thumbprint from ESXi host
    $spec.SslThumbprint = Get-SSLThumbprint "https://$hostname"

    # Using first available Datacenter object to query remote ESXi host 
    return (Get-Datacenter)[0].ExtensionData.QueryConnectionInfoViaSpec($spec)
}

# vCenter Server credentials
$vc_server = "192.168.1.51"
$vc_username = "administrator@vghetto.local"
$vc_password = "VMware1!"

# Remote ESXi credentials to connect
$remote_esxi_hostname = "192.168.1.190"
$remote_esxi_username = "root"
$remote_esxi_password = "vmware123"

$server = Connect-VIServer -Server $vc_server -User $vc_username -Password $vc_password

$result = Get-RemoteESXi -hostname $remote_esxi_hostname -username $remote_esxi_username -password $remote_esxi_password

$result

Disconnect-VIServer $server -Confirm:$false
﻿<#
.SYNOPSIS  Query vCenter Server Database (VCDB) for its
           current usage of the Core, Alarm, Events & Stats table
.DESCRIPTION Script that performs SQL Query against a VCDB running either
             MSSQL & Oracle and collects current usage data for the
             following tables Core, Alarm, Events & Stats table. In
             Addition, if you wish to use the VCSA Migration Tool, the script
             can also calculate the estimated downtime required for either
             migration Option 1 or 2.
.NOTES  Author:    William Lam - @lamw
.NOTES  Site:      www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/09/how-to-check-the-size-of-your-config-stats-events-alarms-tasks-seat-data-in-the-vcdb.html
.PARAMETER dbType
  mssql or oracle
.PARAMETER connectionType
  local (mssql) and remote (mssql or oracle)
.PARAMETER dbServer
   VCDB Server
.PARAMETER dbPort
   VCDB Server Port
.PARAMETER dbUsername
   VCDB Username
.PARAMETER dbPassword
   VCDB Password
.PARAMETER dbInstance
   VCDB Instance Name
.PARAMETER estimate_migration_type
   option1 or option2 for those looking to calculate Windows VC to VCSA Migration (vSphere 6.0 U2m only)
.EXAMPLE
  Run the script locally on the Microsoft SQL Server hosting the vCenter Server Database
  Get-VCDBUsage -dbType mssql -connectionType local -dbServer sql.primp-industries.com
.EXAMPLE
  Run the script remotely on the Microsoft SQL Server hosting the vCenter Server Database
  Get-VCDBUsage -dbType mssql -connectionType local -dbServer sql.primp-industries.com -dbPort 1433 -dbInstance VCDB -dbUsername sa -dbPassword VMware1!
.EXAMPLE
  Run the script remotely on the Microsoft SQL Server hosting the vCenter Server Database & calculate VCSA migration downtime w/option1
  Get-VCDBUsage -dbType mssql -connectionType local -dbServer sql.primp-industries.com -dbPort 1433 -dbInstance VCDB -dbUsername sa -dbPassword VMware1! -migration_type option1
.EXAMPLE
  Run the script remotely to connect to Oracle Sever hosting the vCenter Server Database
  Get-VCDBUsage -dbType oracle -connectionType remote -dbServer oracle.primp-industries.com -dbPort 1521 -dbInstance VCDB -dbUsername vpxuser -dbPassword VMware1!
#>

function UpdateGitHubStats ([string] $csv_stats)
{
    #
    # github token test in psh
    #

    $encoded_token = "YmUxMzZlZWI4ZGI1ZTY3NmJjMGQ1ZmI1MDhjOTYzZGExZDEyNDkzZA=="
    $github_token = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($encoded_token))
    $github_repository = "https://api.github.com/repos/migrate2vcsa/stats/contents/vcsadb.csv?access_token=$github_token"

    $HttpRes = ""

    # Fetch the current file content/commit data (GET)
    try {
        $HttpRes = Invoke-RestMethod -Uri $github_repository -Method "GET" -ContentType "application/json"
    }
    catch {
        Write-Host -ForegroundColor Red "Error connecting to $github_repository"
        Write-Host -ForegroundColor Red $_.Exception.Message
    }


    # Decode base64 text
    $content = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($HttpRes.content))

    # Append any new stuff to the current text file
    $newcontent = $content + "$csv_stats`n"

    # Encode back to base64
    $encoded_content = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($newcontent))

    # Fetch commit sha
    $sha = $HttpRes.sha

    # Generate json response
    $json = @"
    {
        "sha": "$sha",
        "content": "$encoded_content",
        "message": "Updated file",
        "committer": {
            "name" : "vS0ciety",
            "email" : "migratetovcsa@gmail.com"
        }
    }
"@

    # Create the commit request (PUT)
    try {
        $HttpRes = Invoke-RestMethod -Uri $github_repository -Method "PUT" -Body $json -ContentType "application/json"
    }
    catch {
        Write-Host -ForegroundColor Red "Error connecting to $github_repository"
        Write-Host -ForegroundColor Red $_.Exception.Message
    }
}

Function Get-VCDBMigrationTime {
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][Double]$alarmData,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][Double]$coreData,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][Double]$eventData,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][Double]$statData,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$migration_type
    )

    # Sum up total size of the selected migration option
    switch($migration_type) {
        option1 {$VCDBSize=[math]::Round(($coreData + $alarmData),2);break}
        option2 {$VCDBSize=[math]::Round(($coreData + $alarmData + $eventData + $statData),2);break}
    }

    # Formulas extracted from excel spreadsheet from https://kb.vmware.com/kb/2146420
    $H5 = [math]::round(1.62*[math]::pow(2.5, [math]::Log($VCDBSize/75,2)) + (5.47-1.62)/75*$VCDBSize,2)
    $H7 = [math]::round(1.62*[math]::pow(2.5, [math]::Log($VCDBSize/75,2)) + (3.93-1.62)/75*$VCDBSize,2)
    $H6 = $H5 - $H7

    # Calculate timings
    $totalTimeHours = [math]::floor($H5)
    $totalTimeMinutes = [math]::round($H5 - $totalTimeHours,2)*60
    $exportTimeHours = [math]::floor($H6)
    $exportTimeMinutes = [math]::round($H6 - $exportTimeHours,2)*60
    $importTimeHours = [math]::floor($H7)
    $importtTimeminutes = [math]::round($H7 - $importTimeHours,2)*60

    # Return nice description string of selected migration option
    switch($migration_type) {
        option1 { $migrationDescription = "(Core + Alarm = " + $VCDBSize + " GB)";break}
        option2 { $migrationDescription = "(Core + Alarm + Event + Stat = " + $VCDBSize + " GB)";break}
    }

    Write-Host -ForegroundColor Yellow "`nvCenter Server Migration Estimates for"$migration_type $migrationDescription"`n"
    Write-Host "Total  Time :" $totalTimeHours "Hours" $totalTimeMinutes "Minutes"
    Write-Host "Export Time :" $exportTimeHours "Hours" $exportTimeMinutes "Minutes"
    Write-Host "Import Time :" $importTimeHours "Hours" $importtTimeminutes "Minutes`n"
}

Function Get-VCDBUsage {
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$dbType,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$connectionType,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$dbServer,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][int]$dbPort,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$dbUsername,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$dbPassword,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$dbInstance,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)][string]$estimate_migration_type
    )

    $mssql_vcdb_usage_query = @"
use $dbInstance;
select tabletype, sum(rowcounts) as rowcounts,
       sum(spc.usedspaceKB)/1024.0 as usedspaceMB
from
    (select
        s.name as schemaname,
        t.name as tablename,
        p.rows as rowcounts,
        sum(a.used_pages) * 8 as usedspaceKB,
        case
            when t.name like 'VPX_ALARM%' then 'Alarm'
            when t.name like 'VPX_EVENT%' then 'ET'
            when t.name like 'VPX_TASK%' then 'ET'
            when t.name like 'VPX_HIST_STAT%' then 'Stats'
            when t.name = 'VPX_STAT_COUNTER' then 'Stats'
            when t.name = 'VPX_TOPN%' then 'Stats'
            else 'Core'
        end as tabletype
    from
        sys.tables t
    inner join
        sys.schemas s on s.schema_id = t.schema_id
    inner join
        sys.indexes i on t.object_id = i.object_id
    inner join
        sys.partitions p on i.object_id = p.object_id and i.index_id = p.index_id
    inner join
        sys.allocation_units a on p.partition_id = a.container_id
    where
        t.name not like 'dt%'
        and t.is_ms_shipped = 0
        and i.object_id >= 255
    group by
        t.name, s.name, p.rows) as spc

group by tabletype;
"@

    $oracle_vcdb_usage_query = @"
SELECT tabletype,
       SUM(CASE rn WHEN 1 THEN row_cnt ELSE 0 END) AS rowcount,
       ROUND(SUM(sized)/(1024*1024)) usedspaceMB
 FROM (
      SELECT
            CASE
               WHEN segment_name LIKE '%ALARM%' THEN 'Alarm'
               WHEN segment_name LIKE '%EVENT%' THEN 'ET'
               WHEN segment_name LIKE '%TASK%' THEN 'ET'
               WHEN segment_name LIKE '%HIST_STAT%' THEN 'Stats'
               WHEN segment_name LIKE 'VPX_TOPN%' THEN 'Stats'
               ELSE 'Core'
            END AS tabletype,
            row_cnt,
            sized ,
            ROW_NUMBER () OVER (PARTITION BY table_name ORDER BY segment_name) AS rn
       FROM (
            SELECT
                  t.table_name, t.table_name segment_name,
                  t.NUM_ROWS AS row_cnt, s.bytes AS sized
             FROM user_segments s
             JOIN user_tables t ON s.segment_name = t.table_name AND s.segment_type = 'TABLE'
             UNION ALL
            SELECT
                  ti.table_name,i.index_name, ti.NUM_ROWS,s.bytes
             FROM user_segments s
             JOIN user_indexes i ON s.segment_name = i.index_name AND s.segment_type = 'INDEX'
             JOIN user_tables ti ON i.table_name = ti.table_name) table_index ) type_cnt_size
GROUP BY tabletype
"@

    $oracle_odbc_dll_path = "C:\Oracle\odp.net\managed\common\Oracle.ManagedDataAccess.dll"

    Function Run-VCDBMSSQLQuery {

        Function Run-LocalMSSQLQuery {
            Write-Host -ForegroundColor Green "`nRunning Local MSSQL VCDB Usage Query"

            # Check whether Invoke-Sqlcmd cmdlet exists
            if( (Get-Command "Invoke-Sqlcmd" -errorAction SilentlyContinue -CommandType Cmdlet) -eq $null) {
               Write-Host -ForegroundColor Red "Invoke-Sqlcmd cmdlet does not exists on this system, you will need to install SQL Tools or run remotely with DB credentials`n"
               exit
            }

            try {
                $results = Invoke-Sqlcmd -Query $mssql_vcdb_usage_query -ServerInstance $dbServer
            }
            catch { Write-Host -ForegroundColor Red "Unable to connect to the SQL Server. Its possible the SQL Server is not configured to allow remote connections`n"; exit }

            foreach ($result in $results) {
                switch($result.tabletype) {
                    Alarm { $alarm_usage=$result.usedspaceMB; $alarm_rows=$result.rowcounts; break}
                    Core { $core_usage=$result.usedspaceMB; $core_rows=$result.rowcounts; break}
                    ET { $event_usage=$result.usedspaceMB; $event_rows=$result.rowcounts; break}
                    Stats { $stat_usage=$result.usedspaceMB; $stat_rows=$result.rowcounts; break}
                }
            }

            return ($alarm_usage,$core_usage,$event_usage,$stat_usage,$alarm_rows,$core_rows,$event_rows,$stat_rows)
        }

        Function Run-RemoteMSSQLQuery {
            if($dbServer -eq $null -eq $null -or $dbInstance -eq $null -or $dbUsername -eq $null -or $dbPassword -eq $null) {
                Write-host -ForegroundColor Red "One or more parameters is missing for the remote MSSQL Query option`n"
                exit
            }

            Write-Host -ForegroundColor Green "`nRunning Remote MSSQL VCDB Usage Query"

            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
            if($dbPort -eq 0) {
              $SqlConnection.ConnectionString = "Server = $dbServer; Database = $dbInstance; User ID = $dbUsername; Password = $dbPassword;"
            } else {
              $SqlConnection.ConnectionString = "Server = $dbServer, $dbPort; Database = $dbInstance; User ID = $dbUsername; Password = $dbPassword;"
            }

            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
            $SqlCmd.CommandText = $mssql_vcdb_usage_query
            $SqlCmd.Connection = $SqlConnection

            $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
            $SqlAdapter.SelectCommand = $SqlCmd

            try {
                $DataSet = New-Object System.Data.DataSet
                $numRecords = $SqlAdapter.Fill($DataSet)
            } catch { Write-Host -ForegroundColor Red "Unable to connect and execute query on the SQL Server. Its possible the SQL Server is not configured to allow remote connections`n"; exit }

            $SqlConnection.Close()

            foreach ($result in $DataSet.Tables[0]) {
                switch($result.tabletype) {
                    Alarm { $alarm_usage=$result.usedspaceMB; $alarm_rows=$result.rowcounts; break}
                    Core { $core_usage=$result.usedspaceMB; $core_rows=$result.rowcounts; break}
                    ET { $event_usage=$result.usedspaceMB; $event_rows=$result.rowcounts; break}
                    Stats { $stat_usage=$result.usedspaceMB; $stat_rows=$result.rowcounts; break}
                }
            }

            return ($alarm_usage,$core_usage,$event_usage,$stat_usage,$alarm_rows,$core_rows,$event_rows,$stat_rows)
        }

        switch($connectionType) {
            local { Run-LocalMSSQLQuery;break}
            remote { Run-RemoteMSSQLQuery;break}
        }
    }

    Function Run-VCDBOracleQuery {
        if($dbServer -eq $null -or $dbPort -eq $null -or $dbInstance -eq $null -or $dbUsername -eq $null -or $dbPassword -eq $null) {
            Write-host -ForegroundColor Red "One or more parameters is missing for the remote Oracle Query option`n"
            exit
        }

        Write-Host -ForegroundColor Green "`nRunning Remote Oracle VCDB Usage Query"

        if(Test-Path "$oracle_odbc_dll_path") {
            Add-Type -Path "$oracle_odbc_dll_path"

            $connectionString="Data Source = (DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$dbserver)(PORT=$dbport))(CONNECT_DATA=(SERVICE_NAME=$dbinstance)));User Id=$dbusername;Password=$dbpassword;"

            $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connectionString)

            try {
                $connection.open()
                $command=$connection.CreateCommand()
                $command.CommandText=$query
                $reader=$command.ExecuteReader()
            } catch { Write-Host -ForegroundColor Red "Unable to connect to Oracle DB. Ensure your connection info is correct and system allows for remote connections`n"; exit }

            while ($reader.Read()) {
                $table_name = $reader.getValue(0)
                $table_rows = $reader.getValue(1)
                $table_size = $reader.getValue(2)
                switch($table_name) {
                    Alarm { $alarm_usage=$table_size; $alarm_rows=$table_rows; break}
                    Core { $core_usage=$table_size; $core_rows=$table_rows; break}
                    ET { $event_usage=$table_size; $event_rows=$table_rows; break}
                    Stats { $stat_usage=$table_size; $stat_rows=$table_rows; break}
                }
            }
            $connection.Close()

        } else {
            Write-Host -ForegroundColor Red "Unable to find Oracle ODBC DLL which has been defined in the following path: $oracle_odbc_dll_path"
            exit
        }
        return ($alarm_usage,$core_usage,$event_usage,$stat_usage,$alarm_rows,$core_rows,$event_rows,$stat_rows)
    }

    # Run selected DB query and return 4 expected tables from VCDB
    ($alarmData,$coreData,$eventData,$statData,$alarm_rows,$core_rows,$event_rows,$stat_rows) = (0,0,0,0,0,0,0,0)
    switch($dbType) {
        mssql { ($alarmData,$coreData,$eventData,$statData,$alarm_rows,$core_rows,$event_rows,$stat_rows) = Run-VCDBMSSQLQuery; break }
        oracle { ($alarmData,$coreData,$eventData,$statData,$alarm_rows,$core_rows,$event_rows,$stat_rows) = Run-VCDBOracleQuery; break }
        default { Write-Host "mssql or oracle are the only valid dbType options" }
    }

    # Convert data from MB to GB
    $coreData = [math]::Round(($coreData*1024*1024)/1GB,2)
    $alarmData = [math]::Round(($alarmData*1024*1024)/1GB,2)
    $eventData = [math]::Round(($eventData*1024*1024)/1GB,2)
    $statData = [math]::Round(($statData*1024*1024)/1GB,2)

    Write-Host "`nCore Data :"$coreData" GB (rows:"$core_rows")"
    Write-Host "Alarm Data:"$alarmData" GB (rows:"$alarm_rows")"
    Write-Host "Event Data:"$eventData" GB (rows:"$event_rows")"
    Write-Host "Stat Data :"$statData" GB (rows:"$stat_rows")"

    # If user wants VCSA migration estimates, run the additional calculation
    if($estimate_migration_type -eq "option1" -or $estimate_migration_type -eq "option2") {
        Get-VCDBMigrationTime -alarmData $alarmData -coreData $coreData -eventData $eventData -statData $statData -migration_type $estimate_migration_type
    }

    Write-Host -ForegroundColor Magenta `
    "`nWould you like to be able to compare your VCDB Stats with others? `
If so, when prompted, type yes and only the size & # of rows will `
be sent to https://github.com/migrate2vcsa for further processing`n"
    $answer = Read-Host -Prompt "Do you accept (Y or N)"
    if($answer -eq "Y" -or $answer -eq "y") {
        UpdateGitHubStats("$dbType,$alarmData,$coreData,$eventData,$statData,$alarm_rows,$core_rows,$event_rows,$stat_rows")
    }
}

# Please replace variables your own VCDB details
$dbType = "mssql"
$connectionType = "remote"
$dbServer = "sql.primp-industries.com"
$dbPort = "1433"
$dbInstance = "VCDB"
$dbUsername = "sa"
$dbPassword = "VMware1!"

Get-VCDBUsage -connectionType $connectionType -dbType $dbType -dbServer $dbServer -dbPort $dbPort -dbInstance $dbInstance -dbUsername $dbUsername -dbPassword $dbPassword
<#
.SYNOPSIS  Query vCenter Server Database (VCDB) for its
           current usage of the Core, Alarm, Events & Stats table
.DESCRIPTION Script that performs SQL Query against a VCDB running
            on a vPostgres DB
.NOTES  Author:    William Lam - @lamw
.NOTES  Site:      www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/10/how-to-check-the-size-of-your-config-seat-data-in-the-vcdb-in-vpostgres.html
.PARAMETER dbServer
   VCDB Server
.PARAMETER dbName
   VCDB Instance Name
.PARAMETER dbUsername
   VCDB Username
.PARAMETER dbPassword
   VCDB Password
.EXAMPLE
  Get-VCDBUsagevPostgres -dbServer vcenter60-1.primp-industries.com -dbName VCDB -dbUser vc -dbPass "VMware1!"
#>

Function Get-VCDBUsagevPostgres{
    param(
          [string]$dbServer,
          [string]$dbName,
          [string]$dbUser,
          [string]$dbPass
         )

         $query = @"
         SELECT   tabletype,
         sum(reltuples) as rowcount,
         ceil(sum(pg_total_relation_size(oid)) / (1024*1024)) as usedspaceMB
FROM  (
      SELECT   CASE
                  WHEN c.relname LIKE 'vpx_alarm%' THEN 'Alarm'
                  WHEN c.relname LIKE 'vpx_event%' THEN 'ET'
                  WHEN c.relname LIKE 'vpx_task%' THEN 'ET'
                  WHEN c.relname LIKE 'vpx_hist_stat%' THEN 'Stats'
                  WHEN c.relname LIKE 'vpx_topn%' THEN 'Stats'
                  ELSE 'Core'
               END AS tabletype,
               c.reltuples, c.oid
        FROM pg_class C
        LEFT JOIN pg_namespace N
          ON N.oid = C.relnamespace
       WHERE nspname IN ('vc', 'vpx') and relkind in ('r', 't')) t
GROUP BY tabletype;
"@

    $conn = New-Object System.Data.Odbc.OdbcConnection
    $conn.ConnectionString = "Driver={PostgreSQL UNICODE(x64)};Server=$dbServer;Port=5432;Database=$dbName;Uid=$dbUser;Pwd=$dbPass;ReadOnly=1"
    $conn.open()
    $cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
    $ds = New-Object system.Data.DataSet
    (New-Object system.Data.odbc.odbcDataAdapter($cmd)).fill($ds) | out-null
    $conn.close()
    $ds.Tables[0]
}

# Please replace variables your own VCDB details
$dbServer = "vcenter60-1.primp-industries.com"
$dbInstance = "VCDB"
$dbUsername = "vc"
$dbPassword = "ezbo3wrMqkJB6{7t"

Get-VCDBUsagevPostgres $dbServer $dbInstance $dbUsername $dbPassword
<#
.SYNOPSIS  Returns configuration changes for a VM
.DESCRIPTION The function will return the list of configuration changes
    for a given Virtual Machine
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Comment: Modified example from Lucd's blog post http://www.lucd.info/2009/12/18/events-part-3-auditing-vm-device-changes/
.PARAMETER Vm
  Virtual Machine object to query configuration changes
.PARAMETER Hour
  The number of hours to to search for configuration changes, default 8hrs
.EXAMPLE
  PS> Get-VMConfigChanges -vm $VM
.EXAMPLE
  PS> Get-VMConfigChanges -vm $VM -hours 8
#>

Function Get-VMConfigChanges {
    param($vm, $hours=8)

    # Modified code from http://powershell.com/cs/blogs/tips/archive/2012/11/28/removing-empty-object-properties.aspx
    Function prettyPrintEventObject($vmChangeSpec,$task) {
    	$hashtable = $vmChangeSpec |
	    Get-Member -MemberType *Property |
    	Select-Object -ExpandProperty Name |
	    Sort-Object |
    	ForEach-Object -Begin {
  	    	[System.Collections.Specialized.OrderedDictionary]$rv=@{}
  	    	} -process {
  		    if ($vmChangeSpec.$_ -ne $null) {
    		    $rv.$_ = $vmChangeSpec.$_
      		}
	    } -end {$rv}

    	# Add in additional info to the return object (Thanks to Luc's Code)
    	$hashtable.Add('VMName',$task.EntityName)
	    $hashtable.Add('Start', $task.StartTime)
    	$hashtable.Add('End', $task.CompleteTime)
	    $hashtable.Add('State', $task.State)
    	$hashtable.Add('User', $task.Reason.UserName)
      $hashtable.Add('ChainID', $task.EventChainId)

    	# Device Change
	    $vmChangeSpec.DeviceChange | % {
		    if($_.Device -ne $null) {
          $hashtable.Add('Device', $_.Device.GetType().Name)
			    $hashtable.Add('Operation', $_.Operation)
        }
	    }
	    $newVMChangeSpec = New-Object PSObject
	    $newVMChangeSpec | Add-Member ($hashtable) -ErrorAction SilentlyContinue
	    return $newVMChangeSpec
    }

    # Modified code from Luc Dekens http://www.lucd.info/2009/12/18/events-part-3-auditing-vm-device-changes/
    $tasknumber = 999 # Windowsize for task collector
    $eventnumber = 100 # Windowsize for event collector

    $report = @()
    $taskMgr = Get-View TaskManager
    $eventMgr = Get-View eventManager

    $tFilter = New-Object VMware.Vim.TaskFilterSpec
    $tFilter.Time = New-Object VMware.Vim.TaskFilterSpecByTime
    $tFilter.Time.beginTime = (Get-Date).AddHours(-$hours)
    $tFilter.Time.timeType = "startedTime"
    $tFilter.Entity = New-Object VMware.Vim.TaskFilterSpecByEntity
    $tFilter.Entity.Entity = $vm.ExtensionData.MoRef
    $tFilter.Entity.Recursion = New-Object VMware.Vim.TaskFilterSpecRecursionOption
    $tFilter.Entity.Recursion = "self"

    $tCollector = Get-View ($taskMgr.CreateCollectorForTasks($tFilter))

    $dummy = $tCollector.RewindCollector
    $tasks = $tCollector.ReadNextTasks($tasknumber)

    while($tasks){
      $tasks | where {$_.Name -eq "ReconfigVM_Task"} | % {
        $task = $_
        $eFilter = New-Object VMware.Vim.EventFilterSpec
        $eFilter.eventChainId = $task.EventChainId

        $eCollector = Get-View ($eventMgr.CreateCollectorForEvents($eFilter))
        $events = $eCollector.ReadNextEvents($eventnumber)
        while($events){
          $events | % {
            $event = $_
            switch($event.GetType().Name){
              "VmReconfiguredEvent" {
                $event.ConfigSpec | % {
				    $report += prettyPrintEventObject $_ $task
                }
              }
              Default {}
            }
          }
          $events = $eCollector.ReadNextEvents($eventnumber)
        }
        $ecollection = $eCollector.ReadNextEvents($eventnumber)
	    # By default 32 event collectors are allowed. Destroy this event collector.
        $eCollector.DestroyCollector()
      }
      $tasks = $tCollector.ReadNextTasks($tasknumber)
    }

    # By default 32 task collectors are allowed. Destroy this task collector.
    $tCollector.DestroyCollector()

    $report
}

$vcserver = "192.168.1.150"
$vcusername = "administrator@vghetto.local"
$vcpassword = "VMware1!"

Connect-VIServer -Server $vcserver -User $vcusername -Password $vcpassword

$vm = Get-VM "Test-VM"

Get-VMConfigChanges -vm $vm -hours 1

Disconnect-VIServer -Server $vcserver -Confirm:$false
<#
.SYNOPSIS  Returns configuration changes for a VM using vCenter Server Alarm
.DESCRIPTION The function will return the list of configuration changes
    for a given Virtual Machine trigged by vCenter Server Alarm based on
    VmReconfigureEvent
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Comment: Modified example from Lucd's blog post http://www.lucd.info/2009/12/18/events-part-3-auditing-vm-device-changes/
.PARAMETER Moref
  Virtual Machine MoRef ID that generated vCenter Server Alarm
.PARAMETER EventId
  The ID correlating to the ReconfigVM operation
.EXAMPLE
  PS> Get-VMConfigChanges -moref vm-125 -eventId 8389
#>

Function Get-VMConfigChangesFromAlarm {
    param($moref, $eventId)

    # Construct VM object from MoRef ID
    $vm = New-Object VMware.Vim.ManagedObjectReference
    $vm.Type = "VirtualMachine"
    $vm.Value = $moref

    # Modified code from http://powershell.com/cs/blogs/tips/archive/2012/11/28/removing-empty-object-properties.aspx
    Function prettyPrintEventObject($vmChangeSpec,$task) {
    	$hashtable = $vmChangeSpec |
	    Get-Member -MemberType *Property |
    	Select-Object -ExpandProperty Name |
	    Sort-Object |
    	ForEach-Object -Begin {
  	    	[System.Collections.Specialized.OrderedDictionary]$rv=@{}
  	    	} -process {
  		    if ($vmChangeSpec.$_ -ne $null) {
    		    $rv.$_ = $vmChangeSpec.$_
      		}
	    } -end {$rv}

    	# Add in additional info to the return object (Thanks to Luc's Code)
    	$hashtable.Add('VMName',$task.EntityName)
	    $hashtable.Add('Start', $task.StartTime)
    	$hashtable.Add('End', $task.CompleteTime)
	    $hashtable.Add('State', $task.State)
    	$hashtable.Add('User', $task.Reason.UserName)

    	# Device Change
	    $vmChangeSpec.DeviceChange | % {
		    if($_.Device -ne $null) {
		        $hashtable.Add('Device', $_.Device.GetType().Name)
			    $hashtable.Add('Operation', $_.Operation)
            }
	    }
	    $newVMChangeSpec = New-Object PSObject
	    $newVMChangeSpec | Add-Member ($hashtable) -ErrorAction SilentlyContinue
	    return $newVMChangeSpec
    }

    # Modified code from Luc Dekens http://www.lucd.info/2009/12/18/events-part-3-auditing-vm-device-changes/
    $tasknumber = 999 # Windowsize for task collector
    $eventnumber = 100 # Windowsize for event collector

    $report = @()
    $taskMgr = Get-View TaskManager
    $eventMgr = Get-View eventManager

    $tFilter = New-Object VMware.Vim.TaskFilterSpec
    # Need to take eventId substract 1 to get real event
    $tFilter.eventChainId = ([int]$eventId - 1)
    $tFilter.Entity = New-Object VMware.Vim.TaskFilterSpecByEntity
    $tFilter.Entity.Entity = $vm
    $tFilter.Entity.Recursion = New-Object VMware.Vim.TaskFilterSpecRecursionOption
    $tFilter.Entity.Recursion = "self"

    $tCollector = Get-View ($taskMgr.CreateCollectorForTasks($tFilter))

    $dummy = $tCollector.RewindCollector
    $tasks = $tCollector.ReadNextTasks($tasknumber)

    while($tasks){
      $tasks | where {$_.Name -eq "ReconfigVM_Task"} | % {
        $task = $_
        $eFilter = New-Object VMware.Vim.EventFilterSpec
        $eFilter.eventChainId = $task.EventChainId

        $eCollector = Get-View ($eventMgr.CreateCollectorForEvents($eFilter))
        $events = $eCollector.ReadNextEvents($eventnumber)
        while($events){
          $events | % {
            $event = $_
            switch($event.GetType().Name){
              "VmReconfiguredEvent" {
                $event.ConfigSpec | % {
				    $report += prettyPrintEventObject $_ $task
                }
              }
              Default {}
            }
          }
          $events = $eCollector.ReadNextEvents($eventnumber)
        }
        $ecollection = $eCollector.ReadNextEvents($eventnumber)
	    # By default 32 event collectors are allowed. Destroy this event collector.
        $eCollector.DestroyCollector()
      }
      $tasks = $tCollector.ReadNextTasks($tasknumber)
    }

    # By default 32 task collectors are allowed. Destroy this task collector.
    $tCollector.DestroyCollector()

    $report | Out-File -filepath C:\Users\primp\Desktop\alarm.txt -Append
}

$vcserver = "172.30.0.112"
$vcusername = "administrator@vghetto.local"
$vcpassword = "VMware1!"

Connect-VIServer -Server $vcserver -User $vcusername -Password $vcpassword

# Parse vCenter Server Alarm environmental variables
$eventid_from_alarm = $env:VMWARE_ALARM_TRIGGERINGSUMMARY
$moref_from_alarm = $env:VMWARE_ALARM_TARGET_ID

# regex for string within paren http://powershell.com/cs/forums/p/7360/11988.aspx
$regex = [regex]"\((.*)\)"
$string = [regex]::match($eventid_from_alarm, $regex).Groups[1]
$eventid = $string.value

Get-VMConfigChangesFromAlarm -moref $moref_from_alarm -eventId $eventid

Disconnect-VIServer -Server $vcserver -Confirm:$false
﻿# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Retrieves the VM memory overhead for given VM
# Reference: http://www.virtuallyghetto.com/2015/12/easily-retrieve-vm-memory-overhead-using-the-vsphere-6-0-api.html

<#
.SYNOPSIS  Returns VM Ovehead a VM
.DESCRIPTION The function will return VM memory overhead
    for a given Virtual Machine
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.PARAMETER Vm
  Virtual Machine object to query VM memory overhead
.EXAMPLE
  PS> Get-VM "vcenter60-2" | Get-VMMemOverhead
#>

Function Get-VMMemOverhead {
    param(  
    [Parameter(
        Position=0, 
        Mandatory=$true, 
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [Alias('FullName')]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl[]]$VM
    ) 

    process {
        # Retrieve VM & ESXi MoRef
        $vmMoref = $VM.ExtensionData.MoRef
        $vmHostMoref = $VM.ExtensionData.Runtime.Host

        # Retrieve Overhead Memory Manager
        $overheadMgr = Get-View ($global:DefaultVIServer.ExtensionData.Content.OverheadMemoryManager)

        # Get VM Memory overhead
        $overhead = $overheadMgr.LookupVmOverheadMemory($vmMoref,$vmHostMoref)
        Write-Host $VM.Name "has overhead of" ([math]::Round($overhead/1MB,2)).ToString() "MB memory`n"
    }
}<#
.SYNOPSIS  Retrieve the VSAN Policy for a given VM(s) which includes filtering
    of VMs that do not contain a policy (None) or policies in which contains
    Thick Provisioning (e.g Object Space Reservation set to 100)
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.PARAMETER Vm
  Virtual Machine(s) object to query for VSAN VM Storage Policies
.EXAMPLE
  PS> Get-VM * | Get-VSANPolicy -datastore "vsanDatastore"
  PS> Get-VM * | Get-VSANPolicy -datastore "vsanDatastore" -nopolicy $false -thick $true -details $true
#>

Function Get-VSANPolicy {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl[]]$vms,
    [String]$details=$false,
    [String]$datastore,
    [String]$nopolicy=$false,
    [String]$thick=$false
    )

    process {
        foreach ($vm in $vms) {
            # Extract the VSAN UUID for VM Home
            $vm_dir,$vm_vmx = ($vm.ExtensionData.Config.Files.vmPathName).split('/').replace('[','').replace(']','')
            $vmdatastore,$vmhome_obj_uuid = ($vm_dir).split(' ')

            # Process only if we have a match on the specified datastore
            if($vmdatastore -eq $datastore) {
                $cmmds_queries = @()
                $disks_to_uuid_mapping = @{}
                $disks_to_uuid_mapping[$vmhome_obj_uuid] = "VM Home"

                # Create query object for VM home
                $vmhome_query = New-Object VMware.vim.HostVsanInternalSystemCmmdsQuery
                $vmhome_query.Type = "POLICY"
                $vmhome_query.Uuid = $vmhome_obj_uuid

                # Add the VM Home query object to overall cmmds query spec
                $cmmds_queries += $vmhome_query

                # Go through all VMDKs & build query object for each disk
                $devices = $vm.ExtensionData.Config.Hardware.Device
                foreach ($device in $devices) {
                    if($device -is [VMware.Vim.VirtualDisk]) {
                        if($device.backing.backingObjectId) {
                            $disks_to_uuid_mapping[$device.backing.backingObjectId] = $device.deviceInfo.label
                            $disk_query = New-Object VMware.vim.HostVsanInternalSystemCmmdsQuery
                            $disk_query.Type = "POLICY"
                            $disk_query.Uuid = $device.backing.backingObjectId
                            $cmmds_queries += $disk_query
                        }
                    }
                }

                # Access VSAN Internal System to issue the Cmmds query
                $vsanIntSys = Get-View ((Get-View $vm.ExtensionData.Runtime.Host -Property Name, ConfigManager.vsanInternalSystem).ConfigManager.vsanInternalSystem)
                $results = $vsanIntSys.QueryCmmds($cmmds_queries)

                $printed = @{}
                $json = $results | ConvertFrom-Json
                foreach ($j in $json.result) {
                    $storagepolicy_id = $j.content.spbmProfileId

                    # If there's no spbmProfileID, it means there's
                    # no VSAN VM Storage Policy assigned
                    # possibly deployed from vSphere C# Client
                    if($storagepolicy_id -eq $null -and $nopolicy -eq $true) {
                        $object_type = $disks_to_uuid_mapping[$j.uuid]
                        $policy = $j.content

                        # quick/dirty way to only print VM name once
                        if($printed[$vm.name] -eq $null -and $thick -eq $false) {
                            $printed[$vm.name] = "1"
                            Write-Host "`n"$vm.Name
                        }

                        if($details -eq $true -and $thick -eq $false) {
                           Write-Host "$object_type `t` $policy"
                        } elseIf($details -eq $false -and $thick -eq $false) {
                           Write-Host "$object_type `t` None"
                        } else {
                            # Ignore VM Home which will always be thick provisioned
                            if($object_type -ne "VM Home") {
                                if($policy.proportionalCapacity -eq 100) {
                                    Write-Host "`n"$vm.Name
                                    if($details -eq $true) {
                                        Write-Host "$object_type `t` $policy"
                                    } else {
                                        Write-Host "$object_type"
                                    }
                                }
                            }
                        }
                    } elseIf($storagepolicy_id -ne $null -and $nopolicy -eq $false) {
                        $object_type = $disks_to_uuid_mapping[$j.uuid]
                        $policy = $j.content

                        # quick/dirty way to only print VM name once
                        if($printed[$vm.name] -eq $null -and $thick -eq $false) {
                            $printed[$vm.name] = "1"
                            Write-Host "`n"$vm.Name
                        }

                        # Convert the VM Storage Policy ID to human readable name
                        $vsan_policy_name = Get-SpbmStoragePolicy -Id $storagepolicy_id

                        if($details -eq $true -and $thick -eq $false) {
                            Write-Host "$object_type `t` $vsan_policy_name `t` $policy"
                        } elseIf($details -eq $false -and $thick -eq $false) {
                            if($vsan_policy_name -eq $null) {
                                Write-Host "$object_type `t` None"
                            } else {
                                Write-Host "$object_type `t` $vsan_policy_name"
                            }
                        } else {
                            # Ignore VM Home which will always be thick provisioned
                            if($object_type -ne "VM Home") {
                                if($policy.proportionalCapacity -eq 100) {
                                    if($printed[$vm.name] -eq $null) {
                                        $printed[$vm.name] = "1"
                                        Write-Host "`n"$vm.Name
                                    }
                                    if($details -eq $true) {
                                        Write-Host "$object_type `t` $policy"
                                    } else {
                                        Write-Host "$object_type"
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

Connect-VIServer -Server 192.168.1.51 -User administrator@vghetto.local -password VMware1! | Out-Null

Get-VM "Photon-Deployed-From-WebClient*" | Get-VSANPolicy -datastore "vsanDatastore" -thick $true -details $true

Disconnect-VIServer * -Confirm:$false
﻿<#
.SYNOPSIS Using the vSphere API in vCenter Server to collect ESXTOP & vscsiStats metrics
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2017/02/using-the-vsphere-api-in-vcenter-server-to-collect-esxtop-vscsistats-metrics.html
.PARAMETER Vmhost
  ESXi host
.EXAMPLE
  PS> Get-VMHost -Name "esxi-1" | Get-VscsiStats
#>

Function Get-VscsiStats {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$VMHost
    )

    $serviceManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.serviceManager) -property "" -ErrorAction SilentlyContinue

    $locationString = "vmware.host." + $VMHost.Name
    $services = $serviceManager.QueryServiceList($null,$locationString)
    foreach ($service in $services) {
        if($service.serviceName -eq "VscsiStats") {
            $serviceView = Get-View $services.Service -Property "entity"
            $serviceView.ExecuteSimpleCommand("FetchAllHistograms")
            break
        }
    }
}

Connect-VIServer -Server 192.168.1.51 -User administrator@vsphere.local -password VMware1! | Out-Null

Get-VMHost -Name "192.168.1.50" | Get-VscsiStats

Disconnect-VIServer * -Confirm:$falseFunction New-GlobalPermission {
<#
    .DESCRIPTION Script to add/remove vSphere Global Permission
    .NOTES  Author:  William Lam
    .NOTES  Site:    www.virtuallyghetto.com
    .NOTES  Reference: http://www.virtuallyghetto.com/2017/02/automating-vsphere-global-permissions-with-powercli.html
    .PARAMETER vc_server
        vCenter Server Hostname or IP Address
    .PARAMETER vc_username
        VC Username
    .PARAMETER vc_password
        VC Password
    .PARAMETER vc_user
        Name of the user to remove global permission on
    .PARAMETER vc_role_id
        The ID of the vSphere Role (retrieved from Get-VIRole)
    .PARAMETER propagate
        Whether or not to propgate the permission assignment (true/false)
#>
    New-GlobalPermission -vc_server "192.168.1.51" -vc_username "administrator@vsphere.local" -vc_password "VMware1!" -vc_user "VGHETTO\lamw" -vc_role_id "-1" -propagate "true"
    param(
        [Parameter(Mandatory=$true)][string]$vc_server,
        [Parameter(Mandatory=$true)][String]$vc_username,
        [Parameter(Mandatory=$true)][String]$vc_password,
        [Parameter(Mandatory=$true)][String]$vc_user,
        [Parameter(Mandatory=$true)][String]$vc_role_id,
        [Parameter(Mandatory=$true)][String]$propagate
    )

    $secpasswd = ConvertTo-SecureString $vc_password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($vc_username, $secpasswd)

    # vSphere MOB URL to private enableMethods
    $mob_url = "https://$vc_server/invsvc/mob3/?moid=authorizationService&method=AuthorizationService.AddGlobalAccessControlList"

# Ingore SSL Warnings
add-type -TypeDefinition  @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

    # Initial login to vSphere MOB using GET and store session using $vmware variable
    $results = Invoke-WebRequest -Uri $mob_url -SessionVariable vmware -Credential $credential -Method GET

    # Extract hidden vmware-session-nonce which must be included in future requests to prevent CSRF error
    # Credit to https://blog.netnerds.net/2013/07/use-powershell-to-keep-a-cookiejar-and-post-to-a-web-form/ for parsing vmware-session-nonce via Powershell
    if($results.StatusCode -eq 200) {
        $null = $results -match 'name="vmware-session-nonce" type="hidden" value="?([^\s^"]+)"'
        $sessionnonce = $matches[1]
    } else {
        Write-host "Failed to login to vSphere MOB"
        exit 1
    }

    # Escape username
    $vc_user_escaped = [uri]::EscapeUriString($vc_user)

    # The POST data payload must include the vmware-session-nonce variable + URL-encoded
    $body = @"
vmware-session-nonce=$sessionnonce&permissions=%3Cpermissions%3E%0D%0A+++%3Cprincipal%3E%0D%0A++++++%3Cname%3E$vc_user_escaped%3C%2Fname%3E%0D%0A++++++%3Cgroup%3Efalse%3C%2Fgroup%3E%0D%0A+++%3C%2Fprincipal%3E%0D%0A+++%3Croles%3E$vc_role_id%3C%2Froles%3E%0D%0A+++%3Cpropagate%3E$propagate%3C%2Fpropagate%3E%0D%0A%3C%2Fpermissions%3E
"@
    # Second request using a POST and specifying our session from initial login + body request
    Write-Host "Adding Global Permission for $vc_user ..."
    $results = Invoke-WebRequest -Uri $mob_url -WebSession $vmware -Method POST -Body $body

    # Logout out of vSphere MOB
    $mob_logout_url = "https://$vc_server/invsvc/mob3/logout"
    $results = Invoke-WebRequest -Uri $mob_logout_url -WebSession $vmware -Method GET
}

Function Remove-GlobalPermission {
<#
    .DESCRIPTION Script to add/remove vSphere Global Permission
    .NOTES  Author:  William Lam
    .NOTES  Site:    www.virtuallyghetto.com
    .NOTES  Reference: http://www.virtuallyghetto.com/2017/02/automating-vsphere-global-permissions-with-powercli.html
    .PARAMETER vc_server
        vCenter Server Hostname or IP Address
    .PARAMETER vc_username
        VC Username
    .PARAMETER vc_password
        VC Password
    .PARAMETER vc_user
        Name of the user to remove global permission on
    .EXAMPLE
        PS> Remove-GlobalPermission -vc_server "192.168.1.51" -vc_username "administrator@vsphere.local" -vc_password "VMware1!" -vc_user "VGHETTO\lamw"
#>
    param(
        [Parameter(Mandatory=$true)][string]$vc_server,
        [Parameter(Mandatory=$true)][String]$vc_username,
        [Parameter(Mandatory=$true)][String]$vc_password,
        [Parameter(Mandatory=$true)][String]$vc_user
    )

    $secpasswd = ConvertTo-SecureString $vc_password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($vc_username, $secpasswd)

    # vSphere MOB URL to private enableMethods
    $mob_url = "https://$vc_server/invsvc/mob3/?moid=authorizationService&method=AuthorizationService.RemoveGlobalAccess"

# Ingore SSL Warnings
add-type -TypeDefinition  @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

    # Initial login to vSphere MOB using GET and store session using $vmware variable
    $results = Invoke-WebRequest -Uri $mob_url -SessionVariable vmware -Credential $credential -Method GET

    # Extract hidden vmware-session-nonce which must be included in future requests to prevent CSRF error
    # Credit to https://blog.netnerds.net/2013/07/use-powershell-to-keep-a-cookiejar-and-post-to-a-web-form/ for parsing vmware-session-nonce via Powershell
    if($results.StatusCode -eq 200) {
        $null = $results -match 'name="vmware-session-nonce" type="hidden" value="?([^\s^"]+)"'
        $sessionnonce = $matches[1]
    } else {
        Write-host "Failed to login to vSphere MOB"
        exit 1
    }

    # Escape username
    $vc_user_escaped = [uri]::EscapeUriString($vc_user)

    # The POST data payload must include the vmware-session-nonce variable + URL-encoded
    $body = @"
vmware-session-nonce=$sessionnonce&principals=%3Cprincipals%3E%0D%0A+++%3Cname%3E$vc_user_escaped%3C%2Fname%3E%0D%0A+++%3Cgroup%3Efalse%3C%2Fgroup%3E%0D%0A%3C%2Fprincipals%3E
"@
    # Second request using a POST and specifying our session from initial login + body request
    Write-Host "Removing Global Permission for $vc_user ..."
    $results = Invoke-WebRequest -Uri $mob_url -WebSession $vmware -Method POST -Body $body

    # Logout out of vSphere MOB
    $mob_logout_url = "https://$vc_server/invsvc/mob3/logout"
    $results = Invoke-WebRequest -Uri $mob_logout_url -WebSession $vmware -Method GET
}

### Sample Usage of Enable/Disable functions ###

$vc_server = "192.168.1.51"
$vc_username = "administrator@vsphere.local"
$vc_password = "VMware1!"
$vc_role_id = "-1"
$vc_user = "VGHETTO\lamw"
$propagate = "true"

# Connect to vCenter Server
$server = Connect-VIServer -Server $vc_server -User $vc_username -Password $vc_password

#New-GlobalPermission -vc_server $vc_server -vc_username $vc_username -vc_password $vc_password -vc_user $vc_user -vc_role_id $vc_role_id -propagate $propagate

#Remove-GlobalPermission -vc_server $vc_server -vc_username $vc_username -vc_password $vc_password -vc_user $vc_user

# Disconnect from vCenter Server
Disconnect-viserver $server -confirm:$false﻿Function Add-VMGuestInfo {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
	===========================================================================
    .SYNOPSIS
        Function to add Guestinfo properties to a VM
    .EXAMPLE
        $newGuestProperties = @{
            "guestinfo.foo1" = "bar1"
            "guestinfo.foo2" = "bar2"
            "guestinfo.foo3" = "bar3"
        }

        Add-VMGuestInfo -vmname DeployVM -vmguestinfo $newGuestProperties
#>
    param(
        [Parameter(Mandatory=$true)][String]$vmname,
        [Parameter(Mandatory=$true)][Hashtable]$vmguestinfo
    )

    $vm = Get-VM -Name $vmname
    $currentVMExtraConfig = $vm.ExtensionData.config.ExtraConfig

    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec

    $vmguestinfo.GetEnumerator() | Foreach-Object {
        $optionValue = New-Object VMware.Vim.OptionValue
        $optionValue.Key = $_.Key
        $optionValue.Value = $_.Value
        $currentVMExtraConfig += $optionValue
    }
    $spec.ExtraConfig = $currentVMExtraConfig
    $vm.ExtensionData.ReconfigVM($spec)
}

Function Remove-VMGuestInfo {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
	===========================================================================
    .SYNOPSIS
        Function to remove Guestinfo properties to a VM
    .EXAMPLE
        $newGuestProperties = @{
            "guestinfo.foo1" = "bar1"
            "guestinfo.foo2" = "bar2"
            "guestinfo.foo3" = "bar3"
        }

        Remove-VMGuestInfo -vmname DeployVM -vmguestinfo $newGuestProperties
#>
    param(
        [Parameter(Mandatory=$true)][String]$vmname,
        [Parameter(Mandatory=$true)][Hashtable]$vmguestinfo
    )

    $vm = Get-VM -Name $vmname
    $currentVMExtraConfig = $vm.ExtensionData.config.ExtraConfig

    $updatedVMExtraConfig = @()
    foreach ($vmExtraConfig in $currentVMExtraConfig) {
       if(-not ($vmguestinfo.ContainsKey($vmExtraConfig.key))) {
            $updatedVMExtraConfig += $vmExtraConfig
       }
    }
    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $spec.ExtraConfig = $updatedVMExtraConfig
    $vm.ExtensionData.ReconfigVM($spec)
}# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script to import vCenter Server 6.x root certificate to Mac OS X or NIX* system
# Reference: http://www.virtuallyghetto.com/2016/07/automating-the-import-of-vcenter-server-6-x-root-certificate.html

Function Import-VCRootCertificate ([string]$VC_HOSTNAME) {
    # Set the default download directory to current users desktop
    # Download will be saved as cert.zip
    $DOWNLOAD_PATH=[Environment]::GetFolderPath("Desktop")
    $DOWNLOAD_FILE_NAME="cert.zip"
    $DOWNLOAD_FILE_PATH="$DOWNLOAD_PATH\$DOWNLOAD_FILE_NAME"
    $EXTRACTED_CERTS_PATH="$DOWNLOAD_PATH\certs"

    # VAMI URL, easy way to verify if we have Windows VC or VCSA
    $URL = "https://"+$VC_HOSTNAME+":5480"
    $FOUND_VCSA = 1
	
	try {
		# Checking to see if we have a Windows VC or VCSA
		# as they have different SSL Certificate download endpoints
		$websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
		try {
			Write-Host "`nTesting vCenter URL $URL"
			$result = Invoke-WebRequest -Uri $URL -TimeoutSec 5
		}
		catch [System.NotSupportedException] {
			Write-Host $_.Exception -ForegroundColor "Red" -BackgroundColor "Black"
			throw
		}
		catch [System.Net.WebException] {
			Write-Host $_.Exception
			$FOUND_VCSA = 0
		}

		if($FOUND_VCSA) {
			$VC_CERT_DOWNLOAD_URL="https://"+$VC_HOSTNAME+"/certs/download"
		} else {
			$VC_CERT_DOWNLOAD_URL="https://"+$VC_HOSTNAME+"/certs/download.zip"
		}

		# Required to ingore SSL Warnings
		if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type)
		{
			add-type -TypeDefinition  @"
				using System.Net;
				using System.Security.Cryptography.X509Certificates;
				public class TrustAllCertsPolicy : ICertificatePolicy {
					public bool CheckValidationResult(
						ServicePoint srvPoint, X509Certificate certificate,
						WebRequest request, int certificateProblem) {
						return true;
					}
				}
"@
		}
		[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
		
		# Download VC's SSL Certificate
		Write-Host "`nDownloading VC SSL Certificate from $VC_CERT_DOWNLOAD_URL to $DOWNLOAD_FILE_PATH"
		$webclient = New-Object System.Net.WebClient
		$webclient.DownloadFile("$VC_CERT_DOWNLOAD_URL","$DOWNLOAD_FILE_PATH")

		# Extracting SSL Certificate zip file
		Add-Type -AssemblyName System.IO.Compression.FileSystem
		[System.IO.Compression.ZipFile]::ExtractToDirectory($DOWNLOAD_FILE_PATH, "$DOWNLOAD_PATH")

		# Find SSL certificates ending with .0
		$Dir = get-childitem $EXTRACTED_CERTS_PATH -recurse
		$List = $Dir | where {$_.extension -eq ".0"}

		# Thanks to https://lennytech.wordpress.com/2013/06/18/powershell-install-sp-root-cert-to-trusted-root/ for snippet of code
		# Retrieve Trusted Root Certification Store
		$certStore = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Store Root, LocalMachine

		# Import VC SSL Certificate(s) into cert store
		Write-Host "Importing to VC SSL Certificate to Certificate Store"
		foreach ($a in $list) {
			$file = "$EXTRACTED_CERTS_PATH\$a"

			# Get the certificate from the location where it was placed by the export process
			$cert = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Certificate2 $file

			# Open the store with maximum allowed privileges
			$certStore.Open("MaxAllowed")

			# Add the certificate to the store
			$certStore.Add($cert)
		}
		# Close the store
		$certStore.Close()
	}
	catch {
		Write-Host -ForegroundColor "Red" -BackgroundColor "Black" $_.Exception
	}
	finally {
		#clean up
		if (Test-Path $DOWNLOAD_FILE_PATH) {
			Write-Host "Cleaning up, deleting $DOWNLOAD_FILE_PATH"
			Remove-Item $DOWNLOAD_FILE_PATH
		}
		if (Test-Path $EXTRACTED_CERTS_PATH) {
			Write-Host "Cleaning up, deleting $EXTRACTED_CERTS_PATH"
			Remove-Item -Recurse -Force $EXTRACTED_CERTS_PATH
		}
	}
}

Import-VCRootCertificate $Args[0]
# Author: William Lam
# Site: www.virtuallyghetto.com
# Description: Script to automate the installation of vRA 7 IaaS Mgmt Agent
# Reference: http://www.virtuallyghetto.com/2016/02/automating-vrealize-automation-7-simple-minimal-part-2-vra-iaas-agent-deployment.html

# Hostname or IP of vRA Appliance
$VRA_APPLIANCE_HOSTNAME = "vra-appliance.primp-industries.com"
# Username of vRA Appliance
$VRA_APPLIANCE_USERNAME = "root"
# Password of vRA Appliance
$VRA_APPLIANCE_PASSWORD = "VMware1!"
# Path to store vRA Agent on IaaS Mgmt Windows system
$VRA_APPLIANCE_AGENT_DOWNLOAD_PATH = "C:\Windows\Temp\vCAC-IaaSManagementAgent-Setup.msi"
# Path to store vRA Agent installer logs on IaaS Mgmt Windowssystem
$VRA_APPLIANCE_AGENT_INSTALL_LOG = "C:\Windows\Temp\ManagementAgent-Setup.log"

# Credentials to the vRA IaaS Windows System
$VRA_IAAS_SERVICE_USERNAME = "vra-iaas\\Administrator"
$VRA_IAAS_SERVICE_PASSWORD = "!MySuperDuperPassword!"

### DO NOT EDIT BEYOND HERE ###

# URL to vRA Agent on vRA Appliance
$VRA_APPLIANCE_AGENT_URL = "https://" + $VRA_APPLIANCE_HOSTNAME + ":5480/installer/download/vCAC-IaaSManagementAgent-Setup.msi"

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
$webclient = New-Object System.Net.WebClient
$webclient.Credentials = New-Object System.Net.NetworkCredential($VRA_APPLIANCE_USERNAME,$VRA_APPLIANCE_PASSWORD)

Write-Host "Downloading " $VRA_APPLIANCE_AGENT_URL "to" $VRA_APPLIANCE_AGENT_DOWNLOAD_PATH "..."
$webclient.DownloadFile($VRA_APPLIANCE_AGENT_URL,$VRA_APPLIANCE_AGENT_DOWNLOAD_PATH)

# Extracting SSL Thumbprint frmo vRA Appliance
# Thanks to Brian Graf for this snippet!
# I originally used this longer snippet from Alan Renouf (https://communities.vmware.com/thread/501913?start=0&tstart=0)
# Brian 1, Alan 0 ;)
# It's still easier in Linux :D
$VRA_APPLIANCE_ENDPOINT = "https://" + $VRA_APPLIANCE_HOSTNAME + ":5480"

add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;

    public class IDontCarePolicy : ICertificatePolicy {
        public IDontCarePolicy() {}
        public bool CheckValidationResult(
            ServicePoint sPoint, X509Certificate cert,
            WebRequest wRequest, int certProb) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = new-object IDontCarePolicy
$VRA_APPLIANE_VAMI = [System.Net.Webrequest]::Create("$VRA_APPLIANCE_ENDPOINT")
$VRA_APPLIANCE_SSL_THUMBPRINT = $VRA_APPLIANE_VAMI.ServicePoint.Certificate.GetCertHashString()

# Extracting vRA IaaS Windows VM hostname
$VRA_IAAS_HOSTNAME=hostname

# Arguments to silent installer for vRA IaaS Agent
$VRA_INSTALLER_ARGS = "/i $VRA_APPLIANCE_AGENT_DOWNLOAD_PATH /qn /norestart /Lvoicewarmup! `"$VRA_APPLIANCE_AGENT_INSTALL_LOG`" ADDLOCAL=`"ALL`" INSTALLLOCATION=`"C:\\Program Files (x86)\\VMware\\vCAC\\Management Agent`" MANAGEMENT_ENDPOINT_ADDRESS=`"$VRA_APPLIANCE_ENDPOINT`" MANAGEMENT_ENDPOINT_THUMBPRINT=`"$VRA_APPLIANCE_SSL_THUMBPRINT`" SERVICE_USER_NAME=`"$VRA_IAAS_SERVICE_USERNAME`" SERVICE_USER_PASSWORD=`"$VRA_IAAS_SERVICE_PASSWORD`" VA_USER_NAME=`"$VRA_APPLIANCE_USERNAME`" VA_USER_PASSWORD=`"$VRA_APPLIANCE_PASSWORD`" CURRENT_MACHINE_FQDN=`"$VRA_IAAS_HOSTNAME`""

if (Test-Path $VRA_APPLIANCE_AGENT_DOWNLOAD_PATH) {
    Write-Host "Installing vRA 7 Agent ..."
    # Exit code of 0 = success
    $ec = (Start-Process -FilePath msiexec.exe -ArgumentList $VRA_INSTALLER_ARGS -Wait -Passthru).ExitCode
    if ($ec -eq 0) {
        Write-Host "Installation successful!`n"
    } else {
        Write-Host "Installation failed, please have a look at the log!`n"
    }
} else {
    Write-host "Download must have failed as I can not find the file!`n"
}
﻿Function Get-Esxconfig {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function remotely downloads /etc/vmware/config and outputs the content
    .PARAMETER VMHostName
        The name of an individual ESXi host
    .PARAMETER ClusterName
        The name vSphere Cluster
    .EXAMPLE
        Get-Esxconfig
    .EXAMPLE
        Get-Esxconfig -ClusterName cluster-01
    .EXAMPLE
        Get-Esxconfig -VMHostName esxi-01
#>
    param(
        [Parameter(Mandatory=$false)][String]$VMHostName,
        [Parameter(Mandatory=$false)][String]$ClusterName
    )

    if($ClusterName) {
        $cluster = Get-View -ViewType ClusterComputeResource -Property Name,Host -Filter @{"name"=$ClusterName}
        $vmhosts = Get-View $cluster.Host -Property Name
    } elseif($VMHostName) {
        $vmhosts = Get-View -ViewType HostSystem -Property Name -Filter @{"name"=$VMHostName}
    } else {
        $vmhosts = Get-View -ViewType HostSystem -Property Name
    }

    foreach ($vmhost in $vmhosts) {
        $vmhostIp = $vmhost.Name

        $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

        # URL to ESXi's esx.conf configuration file (can use any that show up under https://esxi_ip/host)
        $url = "https://$vmhostIp/host/vmware_config"

        # URL to the ESXi configuration file
        $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
        $spec.Method = "httpGet"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

        # Append the cookie generated from VC
        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie = New-Object System.Net.Cookie
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhost.name
        $websession.Cookies.Add($cookie)

        # Retrieve file
        $result = Invoke-WebRequest -Uri $url -WebSession $websession
        Write-Host "Contents of /etc/vmware/config for $vmhostIp ...`n"
        return $result.content
    }
}

Function Remove-IntelSightingsWorkaround {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function removes the Intel Sightings workaround on an ESXi host as outline by https://kb.vmware.com/s/article/52345
    .PARAMETER AffectedHostList
        Text file containing ESXi Hostnames/IP for hosts you wish to remove remediation
    .EXAMPLE
        Remove-IntelSightingsWorkaround -AffectedHostList hostlist.txt
#>
    param(
        [Parameter(Mandatory=$true)][String]$AffectedHostList
    )

    Function UpdateESXConfig {
        param(
            $VMHost
        )

        $vmhostName = $vmhost.name

        $url = "https://$vmhostName/host/vmware_config"

        $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

        $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
        $spec.Method = "httpGet"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie = New-Object System.Net.Cookie
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhost.name
        $websession.Cookies.Add($cookie)
        $result = Invoke-WebRequest -Uri $url -WebSession $websession
        $esxconfig = $result.content

        # Download the current config file to verify we have not already remediated
        # If not, store existing configuration and append new string
        $remediated = $false
        $newVMwareConfig = ""
        foreach ($line in $esxconfig.Split("`n")) {
            if($line -eq 'cpuid.7.edx = "----:00--:----:----:----:----:----:----"') {
                $remediated = $true
            } else {
                $newVMwareConfig+="$line`n"
            }
        }

        if($remediated -eq $true) {
            Write-Host "`tUpdating /etc/vmware/config ..."

            $newVMwareConfig = $newVMwareConfig.TrimEnd()
            $newVMwareConfig += "`n"

            # Create HTTP PUT spec
            $spec.Method = "httpPut"
            $spec.Url = $url
            $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

            # Upload sfcb.cfg back to ESXi host
            $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
            $cookie.Name = "vmware_cgi_ticket"
            $cookie.Value = $ticket.id
            $cookie.Domain = $vmhost.name
            $websession.Cookies.Add($cookie)
            $result = Invoke-WebRequest -Uri $url -WebSession $websession -Body $newVMwareConfig -Method Put -ContentType "plain/text"
            if($result.StatusCode -eq 200) {
                Write-Host "`tSuccessfully updated VMware config file"
            } else {
                Write-Host "Failed to upload VMware config file"
                break
            }
        } else {
            Write-Host "Remedation not found, skipping host"
        }
    }

    if (Test-Path -Path $AffectedHostList) {
        $affectedHosts = Get-Content -Path $AffectedHostList
        foreach ($affectedHost in $affectedHosts) {
            try {
                $vmhost = Get-View -ViewType HostSystem -Property Name -Filter @{"name"=$affectedHost}
                Write-Host "Processing $affectedHost..."
                UpdateESXConfig -vmhost $vmhost
            } catch {
                Write-Host -ForegroundColor Yellow "Unable to find $affectedHost, skipping ..."
            }
        }
    } else {
        Write-Host -ForegroundColor Red "Can not find $AffectedHostList ..."
    }
}

Function Set-IntelSightingsWorkaround {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function removes the Intel Sightings workaround on an ESXi host as outline by https://kb.vmware.com/s/article/52345
    .PARAMETER AffectedHostList
        Text file containing ESXi Hostnames/IP for hosts you wish to apply remediation
    .EXAMPLE
        Set-IntelSightingsWorkaround -AffectedHostList hostlist.txt
#>
    param(
        [Parameter(Mandatory=$true)][String]$AffectedHostList
    )

    Function UpdateESXConfig {
        param(
            $vmhost
        )

        $vmhostName = $vmhost.name

        $url = "https://$vmhostName/host/vmware_config"

        $sessionManager = Get-View ($global:DefaultVIServer.ExtensionData.Content.sessionManager)

        $spec = New-Object VMware.Vim.SessionManagerHttpServiceRequestSpec
        $spec.Method = "httpGet"
        $spec.Url = $url
        $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

        $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $cookie = New-Object System.Net.Cookie
        $cookie.Name = "vmware_cgi_ticket"
        $cookie.Value = $ticket.id
        $cookie.Domain = $vmhostName
        $websession.Cookies.Add($cookie)
        $result = Invoke-WebRequest -Uri $url -WebSession $websession
        $esxconfig = $result.content

        # Download the current config file to verify we have not already remediated
        # If not, store existing configuration and append new string
        $remediated = $false
        $newVMwareConfig = ""
        foreach ($line in $esxconfig.Split("`n")) {
            if($line -eq 'cpuid.7.edx = "----:00--:----:----:----:----:----:----"') {
                $remediated = $true
                break
            } else {
                $newVMwareConfig+="$line`n"
            }
        }

        if($remediated -eq $false) {
            Write-Host "`tUpdating /etc/vmware/config ..."

            $newVMwareConfig = $newVMwareConfig.TrimEnd()
            $newVMwareConfig+="`ncpuid.7.edx = `"----:00--:----:----:----:----:----:----`"`n"

            # Create HTTP PUT spec
            $spec.Method = "httpPut"
            $spec.Url = $url
            $ticket = $sessionManager.AcquireGenericServiceTicket($spec)

            # Upload sfcb.cfg back to ESXi host
            $websession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
            $cookie.Name = "vmware_cgi_ticket"
            $cookie.Value = $ticket.id
            $cookie.Domain = $vmhostName
            $websession.Cookies.Add($cookie)
            $result = Invoke-WebRequest -Uri $url -WebSession $websession -Body $newVMwareConfig -Method Put -ContentType "plain/text"
            if($result.StatusCode -eq 200) {
                Write-Host "`tSuccessfully updated VMware config file"
            } else {
                Write-Host "Failed to upload VMware config file"
                break
            }
        } else {
            Write-Host "Remedation aleady applied, skipping host"
        }
    }

    if (Test-Path -Path $AffectedHostList) {
        $affectedHosts = Get-Content -Path $AffectedHostList
        foreach ($affectedHost in $affectedHosts) {
            try {
                $vmhost = Get-View -ViewType HostSystem -Property Name -Filter @{"name"=$affectedHost}
                Write-Host "Processing $affectedHost..."
                UpdateESXConfig -vmhost $vmhost
            } catch {
                Write-Host -ForegroundColor Yellow "Unable to find $affectedHost, skipping ..."
            }
        }
    } else {
        Write-Host -ForegroundColor Red "Can not find $AffectedHostList ..."
    }
}Function List-VSANDatastoreFolders {
    # List-DatastoreFolders -DatastoreName WorkloadDatastore
    Param (
        [Parameter(Mandatory=$true)][String]$DatastoreName
    )

    $d = Get-Datastore $DatastoreName
    $br = Get-View $d.ExtensionData.Browser
    $spec = new-object VMware.Vim.HostDatastoreBrowserSearchSpec
    $folderFileQuery= New-Object Vmware.Vim.FolderFileQuery
    $spec.Query = $folderFileQuery
    $fileQueryFlags = New-Object VMware.Vim.FileQueryFlags
    $fileQueryFlags.fileOwner = $false
    $fileQueryFlags.fileSize = $false
    $fileQueryFlags.fileType = $true
    $fileQueryFlags.modification = $false
    $spec.details = $fileQueryFlags
    $spec.sortFoldersFirst = $true
    $results = $br.SearchDatastore("[$($d.Name)]",  $spec)

    $folders = @()
    $files = $results.file
    foreach ($file in $files) {
        if($file.getType().Name -eq "FolderFileInfo") {
            $folderPath = $results.FolderPath + " " + $file.Path

            $tmp = [pscustomobject] @{
                Name = $file.FriendlyName;
                Path = $folderPath;
            }
            $folders+=$tmp
        }
    }
    $folders
}﻿Function Get-MacLearn {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retrieves both the legacy security policies as well as the new
        MAC Learning feature and the new security policies which also live under this
        property which was introduced in vSphere 6.7
    .PARAMETER DVPortgroupName
        The name of Distributed Virtual Portgroup(s)
    .EXAMPLE
        Get-MacLearn -DVPortgroupName @("Nested-01-DVPG")
#>
    param(
        [Parameter(Mandatory=$true)][String[]]$DVPortgroupName
    )

    foreach ($dvpgname in $DVPortgroupName) {
        $dvpg = Get-VDPortgroup -Name $dvpgname -ErrorAction SilentlyContinue
        $switchVersion = ($dvpg | Get-VDSwitch).Version
        if($dvpg -and $switchVersion -eq "6.6.0") {
            $securityPolicy = $dvpg.ExtensionData.Config.DefaultPortConfig.SecurityPolicy
            $macMgmtPolicy = $dvpg.ExtensionData.Config.DefaultPortConfig.MacManagementPolicy

            $securityPolicyResults = [pscustomobject] @{
                DVPortgroup = $dvpgname;
                MacLearning = $macMgmtPolicy.MacLearningPolicy.Enabled;
                NewAllowPromiscuous = $macMgmtPolicy.AllowPromiscuous;
                NewForgedTransmits = $macMgmtPolicy.ForgedTransmits;
                NewMacChanges = $macMgmtPolicy.MacChanges;
                Limit = $macMgmtPolicy.MacLearningPolicy.Limit
                LimitPolicy = $macMgmtPolicy.MacLearningPolicy.limitPolicy
                LegacyAllowPromiscuous = $securityPolicy.AllowPromiscuous.Value;
                LegacyForgedTransmits = $securityPolicy.ForgedTransmits.Value;
                LegacyMacChanges = $securityPolicy.MacChanges.Value;
            }
            $securityPolicyResults
        } else {
            Write-Host -ForegroundColor Red "Unable to find DVPortgroup $dvpgname or VDS is not running 6.6.0"
            break
        }
    }
}

Function Set-MacLearn {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function allows you to manage the new MAC Learning capablitites in
        vSphere 6.7 along with the updated security policies.
    .PARAMETER DVPortgroupName
        The name of Distributed Virtual Portgroup(s)
    .PARAMETER EnableMacLearn
        Boolean to enable/disable MAC Learn
    .PARAMETER EnablePromiscuous
        Boolean to enable/disable the new Prom. Mode property
    .PARAMETER EnableForgedTransmit
        Boolean to enable/disable the Forged Transmit property
    .PARAMETER EnableMacChange
        Boolean to enable/disable the MAC Address change property
    .PARAMETER AllowUnicastFlooding
        Boolean to enable/disable Unicast Flooding (Default $true)
    .PARAMETER Limit
        Define the maximum number of learned MAC Address, maximum is 4096 (default 4096)
    .PARAMETER LimitPolicy
        Define the policy (DROP/ALLOW) when max learned MAC Address limit is reached (default DROP)
    .EXAMPLE
        Set-MacLearn -DVPortgroupName @("Nested-01-DVPG") -EnableMacLearn $true -EnablePromiscuous $false -EnableForgedTransmit $true -EnableMacChange $false
#>
    param(
        [Parameter(Mandatory=$true)][String[]]$DVPortgroupName,
        [Parameter(Mandatory=$true)][Boolean]$EnableMacLearn,
        [Parameter(Mandatory=$true)][Boolean]$EnablePromiscuous,
        [Parameter(Mandatory=$true)][Boolean]$EnableForgedTransmit,
        [Parameter(Mandatory=$true)][Boolean]$EnableMacChange,
        [Parameter(Mandatory=$false)][Boolean]$AllowUnicastFlooding=$true,
        [Parameter(Mandatory=$false)][Int]$Limit=4096,
        [Parameter(Mandatory=$false)][String]$LimitPolicy="DROP"
    )

    foreach ($dvpgname in $DVPortgroupName) {
        $dvpg = Get-VDPortgroup -Name $dvpgname -ErrorAction SilentlyContinue
        $switchVersion = ($dvpg | Get-VDSwitch).Version
        if($dvpg -and $switchVersion -eq "6.6.0") {
            $originalSecurityPolicy = $dvpg.ExtensionData.Config.DefaultPortConfig.SecurityPolicy

            $spec = New-Object VMware.Vim.DVPortgroupConfigSpec
            $dvPortSetting = New-Object VMware.Vim.VMwareDVSPortSetting
            $macMmgtSetting = New-Object VMware.Vim.DVSMacManagementPolicy
            $macLearnSetting = New-Object VMware.Vim.DVSMacLearningPolicy
            $macMmgtSetting.MacLearningPolicy = $macLearnSetting
            $dvPortSetting.MacManagementPolicy = $macMmgtSetting
            $spec.DefaultPortConfig = $dvPortSetting
            $spec.ConfigVersion = $dvpg.ExtensionData.Config.ConfigVersion

            if($EnableMacLearn) {
                $macMmgtSetting.AllowPromiscuous = $EnablePromiscuous
                $macMmgtSetting.ForgedTransmits = $EnableForgedTransmit
                $macMmgtSetting.MacChanges = $EnableMacChange
                $macLearnSetting.Enabled = $EnableMacLearn
                $macLearnSetting.AllowUnicastFlooding = $AllowUnicastFlooding
                $macLearnSetting.LimitPolicy = $LimitPolicy
                $macLearnsetting.Limit = $Limit

                Write-Host "Enabling MAC Learning on DVPortgroup: $dvpgname ..."
                $task = $dvpg.ExtensionData.ReconfigureDVPortgroup_Task($spec)
                $task1 = Get-Task -Id ("Task-$($task.value)")
                $task1 | Wait-Task | Out-Null
            } else {
                $macMmgtSetting.AllowPromiscuous = $false
                $macMmgtSetting.ForgedTransmits = $false
                $macMmgtSetting.MacChanges = $false
                $macLearnSetting.Enabled = $false

                Write-Host "Disabling MAC Learning on DVPortgroup: $dvpgname ..."
                $task = $dvpg.ExtensionData.ReconfigureDVPortgroup_Task($spec)
                $task1 = Get-Task -Id ("Task-$($task.value)")
                $task1 | Wait-Task | Out-Null
            }
        } else {
            Write-Host -ForegroundColor Red "Unable to find DVPortgroup $dvpgname or VDS is not running 6.6.0"
            break
        }
    }
}﻿Function Get-PlaceholderVM {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
    ===========================================================================
    .DESCRIPTION
        This function retrieves all placeholder VMs that are protected by SRM
#>
	$results = @()
	Foreach ($vm in Get-VM) {
		if($vm.ExtensionData.Summary.Config.ManagedBy.Type -eq "placeholderVm") {
			$tmp = [pscustomobject] @{
				Name = $vm.Name;
				ExtKey = $vm.ExtensionData.Summary.Config.ManagedBy.ExtensionKey;
				Type = $vm.ExtensionData.Summary.Config.ManagedBy.Type
			}
			$results+=$tmp
		}
	}
	$results
}# Author: William Lam
# Website: www.virtuallyghetto
# Product: VMware vSphere
# Description: Script to extract ESXi PCI Device details such as Name, Vendor, VID, DID & SVID
# Reference: http://www.virtuallyghetto.com/2015/05/extracting-vid-did-svid-from-pci-devices-in-esxi-using-vsphere-api.html

$server = Connect-VIServer -Server 192.168.1.60 -User administrator@vghetto.local -Password VMware1!

$vihosts = Get-View -Server $server -ViewType HostSystem -Property Name,Hardware.PciDevice

$devices_results = @()

foreach ($vihost in $vihosts) {
	$pciDevices = $vihost.Hardware.PciDevice
	foreach ($pciDevice in $pciDevices) {
		$details = "" | select HOST, DEVICE, VENDOR, VID, DID, SVID
		$vid = [String]::Format("{0:x}", $pciDevice.VendorId)
		$did = [String]::Format("{0:x}", $pciDevice.DeviceId)
		$svid = [String]::Format("{0:x}", $pciDevice.SubVendorId)		

		$details.HOST = $vihost.Name
		$details.DEVICE = $pciDevice.DeviceName
		$details.VENDOR = $pciDevice.VendorName
		$details.VID = $vid
		$details.DID = $did
		$details.SVID = $svid
		$devices_results += $details
	}
}

$devices_results

Disconnect-VIServer $server -Confirm:$false<#
.SYNOPSIS
   This script demonstrates an xVC-vMotion where a live running Virtual Machine 
   is live migrated between two vCenter Servers which are NOT part of the
   same vCenter SSO Domain which is only available using the vSphere 6.0 API
.NOTES
   File Name  : run-cool-xVC-vMotion.ps1
   Author     : William Lam - @lamw
   Version    : 1.0
.LINK
    http://www.virtuallyghetto.com/2015/02/did-you-know-of-an-additional-cool-vmotion-capability-in-vsphere-6-0.html
.LINK
   https://github.com/lamw

.INPUTS
   sourceVC, sourceVCUsername, sourceVCPassword, 
   destVC, destVCUsername, destVCPassword, destVCThumbprint
   datastorename, clustername, vmhostname, vmnetworkname,
   vmname
.OUTPUTS
   Console output

.PARAMETER sourceVC
   The hostname or IP Address of the source vCenter Server
.PARAMETER sourceVCUsername
   The username to connect to source vCenter Server
.PARAMETER sourceVCPassword
   The password to connect to source vCenter Server
.PARAMETER destVC
   The hostname or IP Address of the destination vCenter Server
.PARAMETER destVCUsername
   The username to connect to the destination vCenter Server
.PARAMETER destVCPassword
   The password to connect to the destination vCenter Server
.PARAMETER destVCThumbprint
   The SSL Thumbprint (SHA1) of the destination vCenter Server (Certificate checking is enabled, ensure hostname/IP matches)
.PARAMETER datastorename
   The destination vSphere Datastore where the VM will be migrated to
.PARAMETER clustername
   The destination vSphere Cluster where the VM will be migrated to
.PARAMETER vmhostname
   The destination vSphere ESXi host where the VM will be migrated to
.PARAMETER vmnetworkname
   The destination vSphere VM Portgroup where the VM will be migrated to
.PARAMETER vmname
   The name of the source VM to be migrated
#>
param
(
   [Parameter(Mandatory=$true)]
   [string]
   $sourceVC,
   [Parameter(Mandatory=$true)]
   [string]
   $sourceVCUsername,
   [Parameter(Mandatory=$true)]
   [string]
   $sourceVCPassword,
   [Parameter(Mandatory=$true)]
   [string]
   $destVC,
   [Parameter(Mandatory=$true)]
   [string]
   $destVCUsername,
   [Parameter(Mandatory=$true)]
   [string]
   $destVCPassword,
   [Parameter(Mandatory=$true)]
   [string]
   $destVCThumbprint, 
   [Parameter(Mandatory=$true)]
   [string]
   $datastorename,
   [Parameter(Mandatory=$true)]
   [string]
   $clustername,
   [Parameter(Mandatory=$true)]
   [string]
   $vmhostname,
   [Parameter(Mandatory=$true)]
   [string]
   $vmnetworkname,
   [Parameter(Mandatory=$true)]
   [string]
   $vmname
);

## DEBUGGING
#$source = "LA"
#$vmname = "vMA" 
#
## LA->NY
#if ( $source -eq "LA") {
#  $sourceVC = "vcenter60-4.primp-industries.com"
#  $sourceVCUsername = "administrator@vghetto.local"
#  $sourceVCPassword = "VMware1!"
#  $destVC = "vcenter60-5.primp-industries.com" 
#  $destVCUsername = "administrator@vsphere.local"
#  $destVCpassword = "VMware1!"
#  $datastorename = "vesxi60-8-local-storage"
#  $clustername = "NY-Cluster" 
#  $vmhostname = "vesxi60-8.primp-industries.com"
#  $destVCThumbprint = "82:D0:CF:B5:CC:EA:FE:AE:03:BE:E9:4B:AC:A2:B0:AB:2F:E3:87:49"
#  $vmnetworkname = "NY-VM-Network"
#} else {
## NY->LA
#  $sourceVC = "vcenter60-5.primp-industries.com"
#  $sourceVCUsername = "administrator@vsphere.local"
#  $sourceVCPassword = "VMware1!"
#  $destVC = "vcenter60-4.primp-industries.com" 
#  $destVCUsername = "administrator@vghetto.local"
#  $destVCpassword = "VMware1!" 
#  $datastorename = "vesxi60-7-local-storage"
#  $clustername = "LA-Cluster" 
#  $vmhostname = "vesxi60-7.primp-industries.com"
#  $destVCThumbprint = "B8:46:B9:F3:6C:1D:97:8C:ED:A0:19:92:94:E6:1B:45:15:65:63:96"
#  $vmnetworkname = "LA-VM-Network"
#}

# Connect to Source vCenter Server
$sourceVCConn = Connect-VIServer -Server $sourceVC -user $sourceVCUsername -password $sourceVCPassword
# Connect to Destination vCenter Server
$destVCConn = Connect-VIServer -Server $destVC -user $destVCUsername -password $destVCpassword

# Source VM to migrate
$vm = Get-View (Get-VM -Server $sourceVCConn -Name $vmname) -Property Config.Hardware.Device
# Dest Datastore to migrate VM to
$datastore = (Get-Datastore -Server $destVCConn -Name $datastorename)
# Dest Cluster to migrate VM to
$cluster = (Get-Cluster -Server $destVCConn -Name $clustername)
# Dest ESXi host to migrate VM to
$vmhost = (Get-VMHost -Server $destVCConn -Name $vmhostname)

# Find Ethernet Device on VM to change VM Networks
$devices = $vm.Config.Hardware.Device
foreach ($device in $devices) {
   if($device -is [VMware.Vim.VirtualEthernetCard]) {
      $vmNetworkAdapter = $device
   }
}

# Relocate Spec for Migration
$spec = New-Object VMware.Vim.VirtualMachineRelocateSpec
$spec.datastore = $datastore.Id
$spec.host = $vmhost.Id
$spec.pool = $cluster.ExtensionData.ResourcePool
# New Service Locator required for Destination vCenter Server when not part of same SSO Domain
$service = New-Object VMware.Vim.ServiceLocator
$credential = New-Object VMware.Vim.ServiceLocatorNamePassword
$credential.username = $destVCusername
$credential.password = $destVCpassword
$service.credential = $credential
$service.instanceUuid = $destVCConn.InstanceUuid
$service.sslThumbprint = $destVCThumbprint
$service.url = "https://$destVC"
$spec.service = $service
# Modify VM Network Adapter to new VM Netework (assumption 1 vNIC, but can easily be modified)
$spec.deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec
$spec.deviceChange[0] = New-Object VMware.Vim.VirtualDeviceConfigSpec
$spec.deviceChange[0].Operation = "edit"
$spec.deviceChange[0].Device = $vmNetworkAdapter
$spec.deviceChange[0].Device.backing.deviceName = $vmnetworkname

Write-Host "`nMigrating $vmname from $sourceVC to $destVC ...`n"
# Issue Cross VC-vMotion 
$task = $vm.RelocateVM_Task($spec,"defaultPriority") 
$task1 = Get-Task -Id ("Task-$($task.value)")
$task1 | Wait-Task -Verbose

# Disconnect from Source/Destination VC
Disconnect-VIServer -Server $sourceVCConn -Confirm:$false
Disconnect-VIServer -Server $destVCConn -Confirm:$false# William Lam
# www.virtuallygheto.com
# Using Guest Operations API to invoke command inside of Nested ESXi VM

Function runGuestOpInESXiVM() {
	param(
		$vm_moref,
		$guest_username, 
		$guest_password,
		$guest_command_path,
		$guest_command_args
	)
	
	# Guest Ops Managers
	$guestOpMgr = Get-View $session.ExtensionData.Content.GuestOperationsManager
	$authMgr = Get-View $guestOpMgr.AuthManager
	$procMgr = Get-View $guestOpMgr.processManager
	
	# Create Auth Session Object
	$auth = New-Object VMware.Vim.NamePasswordAuthentication
	$auth.username = $guest_username
	$auth.password = $guest_password
	$auth.InteractiveSession = $false
	
	# Program Spec
	$progSpec = New-Object VMware.Vim.GuestProgramSpec
	# Full path to the command to run inside the guest
	$progSpec.programPath = "$guest_command_path"
	$progSpec.workingDirectory = "/tmp"
	# Arguments to the command path, must include "++goup=host/vim/tmp" as part of the arguments
	$progSpec.arguments = "++group=host/vim/tmp $guest_command_args"
	
	# Issue guest op command
	$cmd_pid = $procMgr.StartProgramInGuest($vm_moref,$auth,$progSpec)
}

$session = Connect-VIServer -Server 192.168.1.60 -User administrator@vghetto.local -Password VMware1!

$esxi_vm = 'Nested-ESXi6'
$esxi_username = 'root'
$esxi_password = 'vmware123'

$vm = Get-VM $esxi_vm

# commands to run inside of Nested ESXi VM
$command_path = '/bin/python'
$command_args = '/bin/esxcli.py system welcomemsg set -m "vGhetto Was Here"'

Write-Host
Write-Host "Invoking command:" $command_path $command_args "to" $esxi_vm
Write-Host
runGuestOpInESXiVM -vm_moref $vm.ExtensionData.MoRef -guest_username $esxi_username -guest_password $esxi_password -guest_command_path $command_path -guest_command_args $command_args

Disconnect-VIServer -Server $session -Confirm:$false﻿Function Get-SecureBoot {
    <#
    .SYNOPSIS Query Seure Boot setting for a VM in vSphere 6.5
    .NOTES  Author:  William Lam
    .NOTES  Site:    www.virtuallyghetto.com
    .PARAMETER Vm
      VM to query Secure Boot setting
    .EXAMPLE
      Get-VM -Name Windows10 | Get-SecureBoot
    #>
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$Vm
     )

     $secureBootSetting = if ($vm.ExtensionData.Config.BootOptions.EfiSecureBootEnabled) { "enabled" } else { "disabled" }
     Write-Host "Secure Boot is" $secureBootSetting
}

Function Set-SecureBoot {
    <#
    .SYNOPSIS Enable/Disable Seure Boot setting for a VM in vSphere 6.5
    .NOTES  Author:  William Lam
    .NOTES  Site:    www.virtuallyghetto.com
    .PARAMETER Vm
      VM to enable/disable Secure Boot
    .EXAMPLE
      Get-VM -Name Windows10 | Set-SecureBoot -Enabled
    .EXAMPLE
      Get-VM -Name Windows10 | Set-SecureBoot -Disabled
    #>
    param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true)
        ]
        [VMware.VimAutomation.ViCore.Impl.V1.Inventory.InventoryItemImpl]$Vm,
        [Switch]$Enabled,
        [Switch]$Disabled
     )

    if($Enabled) {
        $secureBootSetting = $true
        $reconfigMessage = "Enabling Secure Boot for $Vm"
    }
    if($Disabled) {
        $secureBootSetting = $false
        $reconfigMessage = "Disabling Secure Boot for $Vm"
    }

    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $bootOptions = New-Object VMware.Vim.VirtualMachineBootOptions
    $bootOptions.EfiSecureBootEnabled = $secureBootSetting
    $spec.BootOptions = $bootOptions
  
    Write-Host "`n$reconfigMessage ..."
    $task = $vm.ExtensionData.ReconfigVM_Task($spec)
    $task1 = Get-Task -Id ("Task-$($task.value)")
    $task1 | Wait-Task | Out-Null
}
﻿# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Configure SMP-FT for a Virtual Machine in vSphere 6.0
# Reference: http://www.virtuallyghetto.com/2016/02/new-vsphere-6-0-api-for-configuring-smp-ft.html

<#
.SYNOPSIS  Configure SMP-FT for a Virtual Machine
.DESCRIPTION The function will allow you to enable/disable SMP-FT for a Virtual Machine
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.PARAMETER Vmname
  Virtual Machine object to perform SMP-FT operation
.PARAMETER Operation
  on/off
.PARAMETER Datastore
  The Datastore to store secondary VM as well as the VM's configuration file (Default assumes same datastore but this can be changed)
.PARAMETER Vmhost
  The ESXi host in which to store the secondary VM
.EXAMPLE
  PS> Set-FT -vmname "SMP-VM" -Operation [on|off] -Datastore "vsanDatastore" -Vmhost "vesxi60-5.primp-industries.com"
#>

Function Set-FT {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    $vmname,
    $operation,
    $datastore,
    $vmhost
    )

    process {
        # Retrieve VM View
        $vmView = Get-View -ViewType VirtualMachine -Property Name,Config.Hardware.Device -Filter @{"name"=$vmname}

        # Retrieve Datastore View
        $datastoreView = Get-View -ViewType Datastore -Property Name -Filter @{"name"=$datastore}

        # Retrieve ESXi View
        $vmhostView = Get-View -ViewType HostSystem -Property Name -Filter @{"name"=$vmhost}

        # VM Devices
        $devices = $vmView.Config.Hardware.Device

        $diskArray = @()
        # Build VM Disk Array to map to datastore
        foreach ($device in $d) {
	        if($device -is [VMware.Vim.VirtualDisk]) {
		        $temp = New-Object Vmware.Vim.FaultToleranceDiskSpec
                $temp.Datastore = $datastoreView.Moref
                $temp.Disk = $device
                $diskArray += $temp
	        }
        }

        # FT Config Spec
        $spec = New-Object VMware.Vim.FaultToleranceConfigSpec
        $metadataSpec = New-Object VMware.Vim.FaultToleranceMetaSpec
        $metadataSpec.metaDataDatastore = $datastoreView.MoRef
        $secondaryVMSepc = New-Object VMware.Vim.FaultToleranceVMConfigSpec
        $secondaryVMSepc.vmConfig = $datastoreView.MoRef
        $secondaryVMSepc.disks = $diskArray
        $spec.metaDataPath = $metadataSpec
        $spec.secondaryVmSpec = $secondaryVMSepc

        if($operation -eq "on") {
            $task = $vmView.CreateSecondaryVMEx_Task($vmhostView.MoRef,$spec)
        } elseif($operation -eq "off") {
            $task = $vmView.TurnOffFaultToleranceForVM_Task()
        } else {
            Write-Host "Invalid Selection"
            exit 1
        }
        $task1 = Get-Task -Id ("Task-$($task.value)")
        $task1 | Wait-Task
    }
}
<#
.SYNOPSIS  Applies a VSAN VM Storage Policy across a list of Virtual Machines
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.EXAMPLE
  PS> Set-VSANPolicy -listofvms $arrayofvmnames -policy $vsanpolicyname
#>

Function Set-VSANPolicy {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [string[]]$listofvms,
    [String]$policy
    )

    $vmstoremediate = @()
    foreach ($vm in $listofvms) {
        $hds = Get-VM $vm | Get-HardDisk
        Write-Host "`nApplying VSAN VM Storage Policy:" $policy "to" $vm "..."
        Set-SpbmEntityConfiguration -Configuration (Get-SpbmEntityConfiguration $hds) -StoragePolicy $policy
    }
}

Connect-VIServer -Server 192.168.1.51 -User administrator@vghetto.local -password VMware1! | Out-Null

# Define list of VMs you wish to remediate and apply VSAN VM Storage Policy
$listofvms = @(
"Photon-Deployed-From-WebClient-Multiple-Disks-1",
"Photon-Deployed-From-WebClient-Multiple-Disks-2"
)

# Name of VSAN VM Storage Policy to apply
$vsanpolicy = "Virtual SAN Default Storage Policy"

Set-VSANPolicy -listofvms $listofvms -policy $vsanpolicy

Disconnect-VIServer * -Confirm:$false
# Author: William Lam
# Website: www.virtuallyghetto
# Product: VMware vCenter Server Apppliance
# Description: PowerCLI script to deploy VCSA directly to ESXi host
# Reference: http://www.virtuallyghetto.com/2014/06/an-alternate-way-to-inject-ovf-properties-when-deploying-virtual-appliances-directly-onto-esxi.html

$esxname = "mini.primp-industries.com"
$esx = Connect-VIServer -Server $esxname

# Name of VM
$vmname = "VCSA"

# Name of the OVF Env VM Adv Setting
$ovfenv_key = “guestinfo.ovfEnv”

# VCSA Example
$ovfvalue = "<?xml version=`"1.0`" encoding=`"UTF-8`"?> 
<Environment 
     xmlns=`"http://schemas.dmtf.org/ovf/environment/1`" 
     xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`" 
     xmlns:oe=`"http://schemas.dmtf.org/ovf/environment/1`" 
     xmlns:ve=`"http://www.vmware.com/schema/ovfenv`" 
     oe:id=`"`">
   <PlatformSection> 
      <Kind>VMware ESXi</Kind> 
      <Version>5.5.0</Version> 
      <Vendor>VMware, Inc.</Vendor> 
      <Locale>en</Locale> 
   </PlatformSection> 
   <PropertySection> 
         <Property oe:key=`"vami.DNS.VMware_vCenter_Server_Appliance`" oe:value=`"192.168.1.1`"/> 
         <Property oe:key=`"vami.gateway.VMware_vCenter_Server_Appliance`" oe:value=`"192.168.1.1`"/> 
         <Property oe:key=`"vami.hostname`" oe:value=`"vcsa.primp-industries.com`"/> 
         <Property oe:key=`"vami.ip0.VMware_vCenter_Server_Appliance`" oe:value=`"192.168.1.250`"/> 
         <Property oe:key=`"vami.netmask0.VMware_vCenter_Server_Appliance`" oe:value=`"255.255.255.0`"/>  
         <Property oe:key=`"vm.vmname`" oe:value=`"VMware_vCenter_Server_Appliance`"/>
   </PropertySection>
</Environment>"

# Adds "guestinfo.ovfEnv" VM Adv setting to VM
Get-VM $vmname | New-AdvancedSetting -Name $ovfenv_key -Value $ovfvalue -Confirm:$false -Force:$true

Disconnect-VIServer -Server $esx -Confirm:$false
# Author: William Lam
# Website: www.virtuallyghetto
# Product: VMware vSphere
# Description: Script to issue UNMAP command on specified VMFS datastore
# Reference: http://www.virtuallyghetto.com/2014/09/want-to-issue-a-vaai-unmap-operation-using-the-vsphere-web-client.html

param
(
   [Parameter(Mandatory=$true)]
   [VMware.VimAutomation.ViCore.Types.V1.DatastoreManagement.Datastore]
   $datastore,
   [Parameter(Mandatory=$true)]
   [string]
   $numofvmfsblocks
);

# Retrieve a random ESXi host which has access to the selected Datastore
$esxi = (Get-View (($datastore.ExtensionData.Host | Get-Random).key) -Property Name).name

# Retrieve ESXCLI instance from the selected ESXi host
$esxcli = Get-EsxCli -Server $global:DefaultVIServer -VMHost $esxi

# Reclaim based on the number of blocks specified by user
$esxcli.storage.vmfs.unmap($numofvmfsblocks,$datastore,$null)
﻿Function Get-VCVersion {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function extracts the vCenter Server (Windows or VCSA) build from your env
        and maps it to https://kb.vmware.com/kb/2143838 to retrieve the version and release date
    .EXAMPLE
        Get-VCVersion
#>
    param(
        [Parameter(Mandatory=$false)][VMware.VimAutomation.ViCore.Util10.VersionedObjectImpl]$Server
    )

    # Pulled from https://kb.vmware.com/kb/2143838
    $vcenterBuildVersionMappings = @{
        "5973321"="vCenter 6.5 Update 1,2017-07-27"
        "5705665"="vCenter 6.5 0e Express Patch 3,2017-06-15"
        "5318154"="vCenter 6.5 0d Express Patch 2,2017-04-18"
        "5318200"="vCenter 6.0 Update 3b,2017-04-13"
        "5183549"="vCenter 6.0 Update 3a,2017-03-21"
        "5112527"="vCenter 6.0 Update 3,2017-02-24"
        "4541947"="vCenter 6.0 Update 2a,2016-11-22"
        "3634793"="vCenter 6.0 Update 2,2016-03-16"
        "3339083"="vCenter 6.0 Update 1b,2016-01-07"
        "3018524"="vCenter 6.0 Update 1,2015-09-10"
        "2776511"="vCenter 6.0.0b,2015-07-07"
        "2656760"="vCenter 6.0.0a,2015-04-16"
        "2559268"="vCenter 6.0 GA,2015-03-12"
        "4180647"="vCenter 5.5 Update 3e,2016-08-04"
        "3721164"="vCenter 5.5 Update 3d,2016-04-14"
        "3660016"="vCenter 5.5 Update 3c,2016-03-29"
        "3252642"="vCenter 5.5 Update 3b,2015-12-08"
        "3142196"="vCenter 5.5 Update 3a,2015-10-22"
        "3000241"="vCenter 5.5 Update 3,2015-09-16"
        "2646482"="vCenter 5.5 Update 2e,2015-04-16"
        "2001466"="vCenter 5.5 Update 2,2014-09-09"
        "1945274"="vCenter 5.5 Update 1c,2014-07-22"
        "1891313"="vCenter 5.5 Update 1b,2014-06-12"
        "1750787"="vCenter 5.5 Update 1a,2014-04-19"
        "1750596"="vCenter 5.5.0c,2014-04-19"
        "1623099"="vCenter 5.5 Update 1,2014-03-11"
        "1378903"="vCenter 5.5.0a,2013-10-31"
        "1312299"="vCenter 5.5 GA,2013-09-22"
        "3900744"="vCenter 5.1 Update 3d,2016-05-19"
        "3070521"="vCenter 5.1 Update 3b,2015-10-01"
        "2669725"="vCenter 5.1 Update 3a,2015-04-30"
        "2207772"="vCenter 5.1 Update 2c,2014-10-30"
        "1473063"="vCenter 5.1 Update 2,2014-01-16"
        "1364037"="vCenter 5.1 Update 1c,2013-10-17"
        "1235232"="vCenter 5.1 Update 1b,2013-08-01"
        "1064983"="vCenter 5.1 Update 1,2013-04-25"
        "880146"="vCenter 5.1.0a,2012-10-25"
        "799731"="vCenter 5.1 GA,2012-09-10"
        "3891028"="vCenter 5.0 U3g,2016-06-14"
        "3073236"="vCenter 5.0 U3e,2015-10-01"
        "2656067"="vCenter 5.0 U3d,2015-04-30"
        "1300600"="vCenter 5.0 U3,2013-10-17"
        "913577"="vCenter 5.0 U2,2012-12-20"
        "755629"="vCenter 5.0 U1a,2012-07-12"
        "623373"="vCenter 5.0 U1,2012-03-15"
        "5318112"="vCenter 6.5.0c Express Patch 1b,2017-04-13"
        "5178943"="vCenter 6.5.0b,2017-03-14"
        "4944578"="vCenter 6.5.0a Express Patch 01,2017-02-02"
        "4602587"="vCenter 6.5,2016-11-15"
        "5326079"="vCenter 6.0 Update 3b,2017-04-13"
        "5183552"="vCenter 6.0 Update 3a,2017-03-21"
        "5112529"="vCenter 6.0 Update 3,2017-02-24"
        "4541948"="vCenter 6.0 Update 2a,2016-11-22"
        "4191365"="vCenter 6.0 Update 2m,2016-09-15"
        "3634794"="vCenter 6.0 Update 2,2016-03-15"
        "3339084"="vCenter 6.0 Update 1b,2016-01-07"
        "3018523"="vCenter 6.0 Update 1,2015-09-10"
        "2776510"="vCenter 6.0.0b,2015-07-07"
        "2656761"="vCenter 6.0.0a,2015-04-16"
        "2559267"="vCenter 6.0 GA,2015-03-12"
        "4180648"="vCenter 5.5 Update 3e,2016-08-04"
        "3730881"="vCenter 5.5 Update 3d,2016-04-14"
        "3660015"="vCenter 5.5 Update 3c,2016-03-29"
        "3255668"="vCenter 5.5 Update 3b,2015-12-08"
        "3154314"="vCenter 5.5 Update 3a,2015-10-22"
        "3000347"="vCenter 5.5 Update 3,2015-09-16"
        "2646489"="vCenter 5.5 Update 2e,2015-04-16"
        "2442329"="vCenter 5.5 Update 2d,2015-01-27"
        "2183111"="vCenter 5.5 Update 2b,2014-10-09"
        "2063318"="vCenter 5.5 Update 2,2014-09-09"
        "1623101"="vCenter 5.5 Update 1,2014-03-11"
        "1476327"="vCenter 5.5.0b,2013-12-22"
        "1398495"="vCenter 5.5.0a,2013-10-31"
        "1312298"="vCenter 5.5 GA,2013-09-22"
        "3868380"="vCenter 5.1 Update 3d,2016-05-19"
        "3630963"="vCenter 5.1 Update 3c,2016-03-29"
        "3072314"="vCenter 5.1 Update 3b,2015-10-01"
        "2306353"="vCenter 5.1 Update 3,2014-12-04"
        "1882349"="vCenter 5.1 Update 2a,2014-07-01"
        "1474364"="vCenter 5.1 Update 2,2014-01-16"
        "1364042"="vCenter 5.1 Update 1c,2013-10-17"
        "1123961"="vCenter 5.1 Update 1a,2013-05-22"
        "1065184"="vCenter 5.1 Update 1,2013-04-25"
        "947673"="vCenter 5.1.0b,2012-12-20"
        "880472"="vCenter 5.1.0a,2012-10-25"
        "799730"="vCenter 5.1 GA,2012-08-13"
        "3891027"="vCenter 5.0 U3g,2016-06-14"
        "3073237"="vCenter 5.0 U3e,2015-10-01"
        "2656066"="vCenter 5.0 U3d,2015-04-30"
        "2210222"="vCenter 5.0 U3c,2014-11-20"
        "1917469"="vCenter 5.0 U3a,2014-07-01"
        "1302764"="vCenter 5.0 U3,2013-10-17"
        "920217"="vCenter 5.0 U2,2012-12-20"
        "804277"="vCenter 5.0 U1b,2012-08-16"
        "759855"="vCenter 5.0 U1a,2012-07-12"
        "455964"="vCenter 5.0 GA,2011-08-24"
    }

    if(-not $Server) {
        $Server = $global:DefaultVIServer
    }

    $vcBuildNumber = $Server.Build
    $vcName = $Server.Name
    $vcOS = $Server.ExtensionData.Content.About.OsType
    $vcVersion,$vcRelDate = "Unknown","Unknown"

    if($vcenterBuildVersionMappings.ContainsKey($vcBuildNumber)) {
        ($vcVersion,$vcRelDate) = $vcenterBuildVersionMappings[$vcBuildNumber].split(",")
    }

    $tmp = [pscustomobject] @{
        Name = $vcName;
        Build = $vcBuildNumber;
        Version = $vcVersion;
        OS = $vcOS;
        ReleaseDate = $vcRelDate;
    }
    $tmp
}

Function Get-ESXiVersion {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function extracts the ESXi build from your env and maps it to
        https://kb.vmware.com/kb/2143832 to extract the version and release date
    .PARAMETER ClusterName
        Name of the vSphere Cluster to retrieve ESXi version information
    .EXAMPLE
        Get-ESXiVersion -ClusterName VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$ClusterName
    )

    # Pulled from https://kb.vmware.com/kb/2143832
    $esxiBuildVersionMappings = @{
        "5969303"="ESXi 6.5 U1,2017-07-27"
        "5310538"="ESXi 6.5.0d,2017-04-18"
        "5224529"="ESXi 6.5 Express Patch 1a,2017-03-28"
        "5146846"="ESXi 6.5 Patch 01,2017-03-09"
        "4887370"="ESXi 6.5.0a,2017-02-02"
        "4564106"="ESXi 6.5 GA,2016-11-15"
        "5572656"="ESXi 6.0 Patch 5,2017-06-06"
        "5251623"="ESXi 6.0 Express Patch 7c,2017-03-28"
        "5224934"="ESXi 6.0 Express Patch 7a,2017-03-28"
        "5050593"="ESXi 6.0 Update 3,2017-02-24"
        "4600944"="ESXi 6.0 Patch 4,2016-11-22"
        "4510822"="ESXi 6.0 Express Patch 7,2016-10-17"
        "4192238"="ESXi 6.0 Patch 3,2016-08-04"
        "3825889"="ESXi 6.0 Express Patch 6,2016-05-12"
        "3620759"="ESXi 6.0 Update 2,2016-03-16"
        "3568940"="ESXi 6.0 Express Patch 5,2016-02-23"
        "3380124"="ESXi 6.0 Update 1b,2016-01-07"
        "3247720"="ESXi 6.0 Express Patch 4,2015-11-25"
        "3073146"="ESXi 6.0 U1a Express Patch 3,2015-10-06"
        "3029758"="ESXi 6.0 U1,2015-09-10"
        "2809209"="ESXi 6.0.0b,2015-07-07"
        "2715440"="ESXi 6.0 Express Patch 2,2015-05-14"
        "2615704"="ESXi 6.0 Express Patch 1,2015-04-09"
        "2494585"="ESXi 6.0 GA,2015-03-12"
        "5230635"="ESXi 5.5 Express Patch 11,2017-03-28"
        "4722766"="ESXi 5.5 Patch 10,2016-12-20"
        "4345813"="ESXi 5.5 Patch 9,2016-09-15"
        "4179633"="ESXi 5.5 Patch 8,2016-08-04"
        "3568722"="ESXi 5.5 Express Patch 10,2016-02-22"
        "3343343"="ESXi 5.5 Express Patch 9,2016-01-04"
        "3248547"="ESXi 5.5 Update 3b,2015-12-08"
        "3116895"="ESXi 5.5 Update 3a,2015-10-06"
        "3029944"="ESXi 5.5 Update 3,2015-09-16"
        "2718055"="ESXi 5.5 Patch 5,2015-05-08"
        "2638301"="ESXi 5.5 Express Patch 7,2015-04-07"
        "2456374"="ESXi 5.5 Express Patch 6,2015-02-05"
        "2403361"="ESXi 5.5 Patch 4,2015-01-27"
        "2302651"="ESXi 5.5 Express Patch 5,2014-12-02"
        "2143827"="ESXi 5.5 Patch 3,2014-10-15"
        "2068190"="ESXi 5.5 Update 2,2014-09-09"
        "1892794"="ESXi 5.5 Patch 2,2014-07-01"
        "1881737"="ESXi 5.5 Express Patch 4,2014-06-11"
        "1746018"="ESXi 5.5 Update 1a,2014-04-19"
        "1746974"="ESXi 5.5 Express Patch 3,2014-04-19"
        "1623387"="ESXi 5.5 Update 1,2014-03-11"
        "1474528"="ESXi 5.5 Patch 1,2013-12-22"
        "1331820"="ESXi 5.5 GA,2013-09-22"
        "3872664"="ESXi 5.1 Patch 9,2016-05-24"
        "3070626"="ESXi 5.1 Patch 8,2015-10-01"
        "2583090"="ESXi 5.1 Patch 7,2015-03-26"
        "2323236"="ESXi 5.1 Update 3,2014-12-04"
        "2191751"="ESXi 5.1 Patch 6,2014-10-30"
        "2000251"="ESXi 5.1 Patch 5,2014-07-31"
        "1900470"="ESXi 5.1 Express Patch 5,2014-06-17"
        "1743533"="ESXi 5.1 Patch 4,2014-04-29"
        "1612806"="ESXi 5.1 Express Patch 4,2014-02-27"
        "1483097"="ESXi 5.1 Update 2,2014-01-16"
        "1312873"="ESXi 5.1 Patch 3,2013-10-17"
        "1157734"="ESXi 5.1 Patch 2,2013-07-25"
        "1117900"="ESXi 5.1 Express Patch 3,2013-05-23"
        "1065491"="ESXi 5.1 Update 1,2013-04-25"
        "1021289"="ESXi 5.1 Express Patch 2,2013-03-07"
        "914609"="ESXi 5.1 Patch 1,2012-12-20"
        "838463"="ESXi 5.1.0a,2012-10-25"
        "799733"="ESXi 5.1.0 GA,2012-09-10"
        "3982828"="ESXi 5.0 Patch 13,2016-06-14"
        "3086167"="ESXi 5.0 Patch 12,2015-10-01"
        "2509828"="ESXi 5.0 Patch 11,2015-02-24"
        "2312428"="ESXi 5.0 Patch 10,2014-12-04"
        "2000308"="ESXi 5.0 Patch 9,2014-08-28"
        "1918656"="ESXi 5.0 Express Patch 6,2014-07-01"
        "1851670"="ESXi 5.0 Patch 8,2014-05-29"
        "1489271"="ESXi 5.0 Patch 7,2014-01-23"
        "1311175"="ESXi 5.0 Update 3,2013-10-17"
        "1254542"="ESXi 5.0 Patch 6,2013-08-29"
        "1117897"="ESXi 5.0 Express Patch 5,2013-05-15"
        "1024429"="ESXi 5.0 Patch 5,2013-03-28"
        "914586"="ESXi 5.0 Update 2,2012-12-20"
        "821926"="ESXi 5.0 Patch 4,2012-09-27"
        "768111"="ESXi 5.0 Patch 3,2012-07-12"
        "721882"="ESXi 5.0 Express Patch 4,2012-06-14"
        "702118"="ESXi 5.0 Express Patch 3,2012-05-03"
        "653509"="ESXi 5.0 Express Patch 2,2012-04-12"
        "623860"="ESXi 5.0 Update 1,2012-03-15"
        "515841"="ESXi 5.0 Patch 2,2011-12-15"
        "504890"="ESXi 5.0 Express Patch 1,2011-11-03"
        "474610"="ESXi 5.0 Patch 1,2011-09-13"
        "469512"="ESXi 5.0 GA,2011-08-24"
    }

    $cluster = Get-Cluster -Name $ClusterName -ErrorAction SilentlyContinue
    if($cluster -eq $null) {
        Write-Host -ForegroundColor Red "Error: Unable to find vSAN Cluster $ClusterName ..."
        break
    }

    $results = @()
    foreach ($vmhost in $cluster.ExtensionData.Host) {
        $vmhost_view = Get-View $vmhost -Property Name, Config, ConfigManager.ImageConfigManager

        $esxiName = $vmhost_view.name
        $esxiBuild = $vmhost_view.Config.Product.Build
        $esxiVersionNumber = $vmhost_view.Config.Product.Version
        $esxiVersion,$esxiRelDate,$esxiOrigInstallDate = "Unknown","Unknown","N/A"

        if($esxiBuildVersionMappings.ContainsKey($esxiBuild)) {
            ($esxiVersion,$esxiRelDate) = $esxiBuildVersionMappings[$esxiBuild].split(",")
        }

        # Install Date API was only added in 6.5
        if($esxiVersionNumber -eq "6.5.0") {
            $imageMgr = Get-View $vmhost_view.ConfigManager.ImageConfigManager
            $esxiOrigInstallDate = $imageMgr.installDate()
        }

        $tmp = [pscustomobject] @{
            Name = $esxiName;
            Build = $esxiBuild;
            Version = $esxiVersion;
            ReleaseDate = $esxiRelDate;
            OriginalInstallDate = $esxiOrigInstallDate;
        }
        $results+=$tmp
    }
    $results
}

Function Get-VSANVersion {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function extracts the ESXi build from your env and maps it to
        https://kb.vmware.com/kb/2150753 to extract the vSAN version and release date
    .PARAMETER ClusterName
        Name of a vSAN Cluster to retrieve vSAN version information
    .EXAMPLE
        Get-VSANVersion -ClusterName VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$ClusterName
    )

    # Pulled from https://kb.vmware.com/kb/2150753
    $vsanBuildVersionMappings = @{
        "5969303"="vSAN 6.6.1,ESXi 6.5 Update 1,2017-07-27"
        "5310538"="vSAN 6.6,ESXi 6.5.0d,2017-04-18"
        "5224529"="vSAN 6.5 Express Patch 1a,ESXi 6.5 Express Patch 1a,2017-03-28"
        "5146846"="vSAN 6.5 Patch 01,ESXi 6.5 Patch 01,2017-03-09"
        "4887370"="vSAN 6.5.0a,ESXi 6.5.0a,2017-02-02"
        "4564106"="vSAN 6.5,ESXi 6.5 GA,2016-11-15"
        "5572656"="vSAN 6.2 Patch 5,ESXi 6.0 Patch 5,2017-06-06"
        "5251623"="vSAN 6.2 Express Patch 7c,ESXi 6.0 Express Patch 7c,2017-03-28"
        "5224934"="vSAN 6.2 Express Patch 7a,ESXi 6.0 Express Patch 7a,2017-03-28"
        "5050593"="vSAN 6.2 Update 3,ESXi 6.0 Update 3,2017-02-24"
        "4600944"="vSAN 6.2 Patch 4,ESXi 6.0 Patch 4,2016-11-22"
        "4510822"="vSAN 6.2 Express Patch 7,ESXi 6.0 Express Patch 7,2016-10-17"
        "4192238"="vSAN 6.2 Patch 3,ESXi 6.0 Patch 3,2016-08-04"
        "3825889"="vSAN 6.2 Express Patch 6,ESXi 6.0 Express Patch 6,2016-05-12"
        "3620759"="vSAN 6.2,ESXi 6.0 Update 2,2016-03-16"
        "3568940"="vSAN 6.1 Express Patch 5,ESXi 6.0 Express Patch 5,2016-02-23"
        "3380124"="vSAN 6.1 Update 1b,ESXi 6.0 Update 1b,2016-01-07"
        "3247720"="vSAN 6.1 Express Patch 4,ESXi 6.0 Express Patch 4,2015-11-25"
        "3073146"="vSAN 6.1 U1a (Express Patch 3),ESXi 6.0 U1a (Express Patch 3),2015-10-06"
        "3029758"="vSAN 6.1,ESXi 6.0 U1,2015-09-10"
        "2809209"="vSAN 6.0.0b,ESXi 6.0.0b,2015-07-07"
        "2715440"="vSAN 6.0 Express Patch 2,ESXi 6.0 Express Patch 2,2015-05-14"
        "2615704"="vSAN 6.0 Express Patch 1,ESXi 6.0 Express Patch 1,2015-04-09"
        "2494585"="vSAN 6.0,ESXi 6.0 GA,2015-03-12"
        "5230635"="vSAN 5.5 Express Patch 11,ESXi 5.5 Express Patch 11,2017-03-28"
        "4722766"="vSAN 5.5 Patch 10,ESXi 5.5 Patch 10,2016-12-20"
        "4345813"="vSAN 5.5 Patch 9,ESXi 5.5 Patch 9,2016-09-15"
        "4179633"="vSAN 5.5 Patch 8,ESXi 5.5 Patch 8,2016-08-04"
        "3568722"="vSAN 5.5 Express Patch 10,ESXi 5.5 Express Patch 10,2016-02-22"
        "3343343"="vSAN 5.5 Express Patch 9,ESXi 5.5 Express Patch 9,2016-01-04"
        "3248547"="vSAN 5.5 Update 3b,ESXi 5.5 Update 3b,2015-12-08"
        "3116895"="vSAN 5.5 Update 3a,ESXi 5.5 Update 3a,2015-10-06"
        "3029944"="vSAN 5.5 Update 3,ESXi 5.5 Update 3,2015-09-16"
        "2718055"="vSAN 5.5 Patch 5,ESXi 5.5 Patch 5,2015-05-08"
        "2638301"="vSAN 5.5 Express Patch 7,ESXi 5.5 Express Patch 7,2015-04-07"
        "2456374"="vSAN 5.5 Express Patch 6,ESXi 5.5 Express Patch 6,2015-02-05"
        "2403361"="vSAN 5.5 Patch 4,ESXi 5.5 Patch 4,2015-01-27"
        "2302651"="vSAN 5.5 Express Patch 5,ESXi 5.5 Express Patch 5,2014-12-02"
        "2143827"="vSAN 5.5 Patch 3,ESXi 5.5 Patch 3,2014-10-15"
        "2068190"="vSAN 5.5 Update 2,ESXi 5.5 Update 2,2014-09-09"
        "1892794"="vSAN 5.5 Patch 2,ESXi 5.5 Patch 2,2014-07-01"
        "1881737"="vSAN 5.5 Express Patch 4,ESXi 5.5 Express Patch 4,2014-06-11"
        "1746018"="vSAN 5.5 Update 1a,ESXi 5.5 Update 1a,2014-04-19"
        "1746974"="vSAN 5.5 Express Patch 3,ESXi 5.5 Express Patch 3,2014-04-19"
        "1623387"="vSAN 5.5,ESXi 5.5 Update 1,2014-03-11"
    }

    $cluster = Get-Cluster -Name $ClusterName -ErrorAction SilentlyContinue
    if($cluster -eq $null) {
        Write-Host -ForegroundColor Red "Error: Unable to find vSAN Cluster $ClusterName ..."
        break
    }

    $results = @()
    foreach ($vmhost in $cluster.ExtensionData.Host) {
        $vmhost_view = Get-View $vmhost -Property Name, Config, ConfigManager.ImageConfigManager

        $esxiName = $vmhost_view.name
        $esxiBuild = $vmhost_view.Config.Product.Build
        $esxiVersionNumber = $vmhost_view.Config.Product.Version
        $vsanVersion,$esxiVersion,$esxiRelDate = "Unknown","Unknown","Unknown"

        # Technically as of vSAN 6.2 Mgmt API, this information is already built in natively within
        # the product to retrieve ESXi/VC/vSAN Versions
        # See https://github.com/lamw/vghetto-scripts/blob/master/powershell/VSANVersion.ps1
        if($vsanBuildVersionMappings.ContainsKey($esxiBuild)) {
            ($vsanVersion,$esxiVersion,$esxiRelDate) = $vsanBuildVersionMappings[$esxiBuild].split(",")
        }

        $tmp = [pscustomobject] @{
            Name = $esxiName;
            Build = $esxiBuild;
            VSANVersion = $vsanVersion;
            ESXiVersion = $esxiVersion;
            ReleaseDate = $esxiRelDate;
        }
        $results+=$tmp
    }
    $results
}Function Verify-ESXiMeltdownAccelerationInVM {
<#
    .NOTES
    ===========================================================================
     Created by:    Adam Robinson
     Organization:  University of Michigan
        ===========================================================================
    .DESCRIPTION
        This function helps verify if a virtual machine supports the PCID and INVPCID
        instructions.  These can be passed to guests with hardware version 11+
        and can provide performance improvements to Meltdown mitigation.

        This script can return all VMs or you can specify
        a vSphere Cluster to limit the scope or an individual VM
    .PARAMETER VMName
        The name of an individual Virtual Machine
    .EXAMPLE
        Verify-ESXiMeltdownAccelerationInVM
    .EXAMPLE
        Verify-ESXiMeltdownAccelerationInVM -ClusterName cluster-01
    .EXAMPLE
        Verify-ESXiMeltdownAccelerationInVM -VMName vm-01
#>
    param(
        [Parameter(Mandatory=$false)][String]$VMName,
        [Parameter(Mandatory=$false)][String]$ClusterName
    )

    if($ClusterName) {
        $cluster = Get-View -ViewType ClusterComputeResource -Property Name,ResourcePool -Filter @{"name"=$ClusterName}
        $vms = Get-View ((Get-View $cluster.ResourcePool).VM) -Property Name,Config.Version,Runtime.PowerState,Runtime.FeatureRequirement
    } elseif($VMName) {
        $vms = Get-View -ViewType VirtualMachine -Property Name,Config.Version,Runtime.PowerState,Runtime.FeatureRequirement -Filter @{"name"=$VMName}
    } else {
        $vms = Get-View -ViewType VirtualMachine -Property Name,Config.Version,Runtime.PowerState,Runtime.FeatureRequirement
    }

    $results = @()
    foreach ($vm in $vms | Sort-Object -Property Name) {
        # Only check VMs that are powered on
        if($vm.Runtime.PowerState -eq "poweredOn") {
            $vmDisplayName = $vm.Name
            $vmvHW = $vm.Config.Version

            $PCIDPass = $false
            $INVPCIDPass = $false

            $cpuFeatures = $vm.Runtime.FeatureRequirement
            foreach ($cpuFeature in $cpuFeatures) {
                if($cpuFeature.key -eq "cpuid.PCID") {
                    $PCIDPass = $true
                } elseif($cpuFeature.key -eq "cpuid.INVPCID") {
                    $INVPCIDPass = $true
                }
            }

            $meltdownAcceleration = $false
            if ($PCIDPass -and $INVPCIDPass) {
                $meltdownAcceleration = $true
            }

            $tmp = [pscustomobject] @{
                VM = $vmDisplayName;
                PCID = $PCIDPass;
                INVPCID = $INVPCIDPass;
                vHW = $vmvHW;
                MeltdownAcceleration = $meltdownAcceleration
            }
            $results+=$tmp
        }
    }
    $results | ft
}
Function Verify-ESXiMeltdownAcceleration {
<#
    .NOTES
    ===========================================================================
     Created by:    Adam Robinson
     Organization:  University of Michigan
        ===========================================================================
    .DESCRIPTION
        This function helps verify if the ESXi host supports the PCID and INVPCID
        instructions.  These can be passed to guests with hardware version 11+
        and can provide performance improvements to Meltdown mitigation.

        This script can return all ESXi hosts or you can specify
        a vSphere Cluster to limit the scope or an individual ESXi host
    .PARAMETER VMHostName
        The name of an individual ESXi host
    .PARAMETER ClusterName
        The name vSphere Cluster
    .EXAMPLE
        Verify-ESXiMeltdownAcceleration
    .EXAMPLE
        Verify-ESXiMeltdownAcceleration -ClusterName cluster-01
    .EXAMPLE
        Verify-ESXiMeltdownAcceleration -VMHostName esxi-01
#>
    param(
        [Parameter(Mandatory=$false)][String]$VMHostName,
        [Parameter(Mandatory=$false)][String]$ClusterName
    )

    $accelerationEVCModes = @("intel-broadwell","intel-haswell","Disabled")

    if($ClusterName) {
        $cluster = Get-View -ViewType ClusterComputeResource -Property Name,Host -Filter @{"name"=$ClusterName}
        $vmhosts = Get-View $cluster.Host -Property Name,Config.FeatureCapability,Hardware.CpuFeature,Summary.CurrentEVCModeKey
    } elseif($VMHostName) {
        $vmhosts = Get-View -ViewType HostSystem -Property Name,Config.FeatureCapability,Hardware.CpuFeature,Summary.CurrentEVCModeKey -Filter @{"name"=$VMHostName}
    } else {
        $vmhosts = Get-View -ViewType HostSystem -Property Name,Config.FeatureCapability,Hardware.CpuFeature,Summary.CurrentEVCModeKey
    }

    $results = @()
    foreach ($vmhost in $vmhosts | Sort-Object -Property Name) {
        $vmhostDisplayName = $vmhost.Name

        $evcMode = $vmhost.Summary.CurrentEVCModeKey
        if ($evcMode -eq $null) {
            $evcMode = "Disabled"
        }

        $PCIDPass = $false
        $INVPCIDPass = $false

        #output from $vmhost.Hardware.CpuFeature is a binary string ':' delimited to nibbles
        #the easiest way I could figure out the hex conversion was to make a byte array
        $cpuidEAX = ($vmhost.Hardware.CpuFeature | Where-Object {$_.Level -eq 1}).Eax -Replace ":","" -Split "(?<=\G\d{8})(?=\d{8})"
        $cpuSignature = ($cpuidEAX | Foreach-Object {[System.Convert]::ToByte($_, 2)} | Foreach-Object {$_.ToString("X2")}) -Join ""
        $cpuSignature = "0x" + $cpuSignature

        $cpuFamily = [System.Convert]::ToByte($cpuidEAX[2], 2).ToString("X2")

        $cpuFeatures = $vmhost.Config.FeatureCapability
        foreach ($cpuFeature in $cpuFeatures) {
            if($cpuFeature.key -eq "cpuid.PCID" -and $cpuFeature.value -eq 1) {
                $PCIDPass = $true
            } elseif($cpuFeature.key -eq "cpuid.INVPCID" -and $cpuFeature.value -eq 1) {
                $INVPCIDPass = $true
            }
        }

        $HWv11Acceleration = $false
        if ($cpuFamily -eq "06") {
            if ($PCIDPass -and $INVPCIDPass) {
                if ($accelerationEVCModes -contains $evcMode) {
                    $HWv11Acceleration = $true
                }
                else {
                    $HWv11Acceleration = "EVCTooLow"
                }
            }
        }
        else {
            $HWv11Acceleration = "Unneeded"
        }

        $tmp = [pscustomobject] @{
            VMHost = $vmhostDisplayName;
            PCID = $PCIDPass;
            INVPCID = $INVPCIDPass;
            EVCMode = $evcMode
            "vHW11+Acceleration" = $HWv11Acceleration;
        }
        $results+=$tmp
    }
    $results | ft
}Function Verify-ESXiMicrocodePatchAndVM {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function helps verify both ESXi Patch and Microcode updates have been
        applied as stated per https://kb.vmware.com/s/article/52085

        This script can return all VMs or you can specify
        a vSphere Cluster to limit the scope or an individual VM
    .PARAMETER VMName
        The name of an individual Virtual Machine
    .EXAMPLE
        Verify-ESXiMicrocodePatchAndVM
    .EXAMPLE
        Verify-ESXiMicrocodePatchAndVM -ClusterName cluster-01
    .EXAMPLE
        Verify-ESXiMicrocodePatchAndVM -VMName vm-01
#>
    param(
        [Parameter(Mandatory=$false)][String]$VMName,
        [Parameter(Mandatory=$false)][String]$ClusterName
    )

    if($ClusterName) {
        $cluster = Get-View -ViewType ClusterComputeResource -Property Name,ResourcePool -Filter @{"name"=$ClusterName}
        $vms = Get-View ((Get-View $cluster.ResourcePool).VM) -Property Name,Config.Version,Runtime.PowerState,Runtime.FeatureRequirement
    } elseif($VMName) {
        $vms = Get-View -ViewType VirtualMachine -Property Name,Config.Version,Runtime.PowerState,Runtime.FeatureRequirement -Filter @{"name"=$VMName}
    } else {
        $vms = Get-View -ViewType VirtualMachine -Property Name,Config.Version,Runtime.PowerState,Runtime.FeatureRequirement
    }

    $results = @()
    foreach ($vm in $vms | Sort-Object -Property Name) {
        # Only check VMs that are powered on
        if($vm.Runtime.PowerState -eq "poweredOn") {
            $vmDisplayName = $vm.Name
            $vmvHW = $vm.Config.Version

            $vHWPass = $false
            $IBRSPass = $false
            $IBPBPass = $false
            $STIBPPass = $false
            $vmAffected = $true
            if ($vmvHW -match 'vmx-[0-9]{2}') {
              if ( [int]$vmvHW.Split('-')[-1] -gt 8 ) {
                $vHWPass = $true
              } else {
                $vHWPass = "N/A"
              }

              $cpuFeatures = $vm.Runtime.FeatureRequirement
              foreach ($cpuFeature in $cpuFeatures) {
                  if($cpuFeature.key -eq "cpuid.IBRS") {
                      $IBRSPass = $true
                  } elseif($cpuFeature.key -eq "cpuid.IBPB") {
                      $IBPBPass = $true
                  } elseif($cpuFeature.key -eq "cpuid.STIBP") {
                      $STIBPPass = $true
                  }
              }
              
              if( ($IBRSPass -eq $true -or $IBPBPass -eq $true -or $STIBPPass -eq $true) -and $vHWPass -eq $true) {
                  $vmAffected = $false
              } elseif($vHWPass -eq "N/A") {
                  $vmAffected = $vHWPass
              }
            } else {
              $IBRSPass = "N/A"
              $IBPBPass = "N/A"
              $STIBPPass = "N/A"
              $vmAffected = "N/A"
            }

            $tmp = [pscustomobject] @{
                VM = $vmDisplayName;
                IBRSPresent = $IBRSPass;
                IBPBPresent = $IBPBPass;
                STIBPPresent = $STIBPPass;
                vHW = $vmvHW;
                HypervisorAssistedGuestAffected = $vmAffected;
            }
            $results+=$tmp
        }
    }
    $results | ft
}

Function Verify-ESXiMicrocodePatch {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function helps verify only the ESXi Microcode update has been
        applied as stated per https://kb.vmware.com/s/article/52085

        This script can return all ESXi hosts or you can specify
        a vSphere Cluster to limit the scope or an individual ESXi host
    .PARAMETER VMHostName
        The name of an individual ESXi host
    .PARAMETER ClusterName
        The name vSphere Cluster
    .EXAMPLE
        Verify-ESXiMicrocodePatch
    .EXAMPLE
        Verify-ESXiMicrocodePatch -ClusterName cluster-01
    .EXAMPLE
        Verify-ESXiMicrocodePatch -VMHostName esxi-01
    .EXAMPLE
        Verify-ESXiMicrocodePatch -ClusterName "Virtual SAN Cluster" -IncludeMicrocodeVerCheck $true -PlinkPath "C:\Users\lamw\Desktop\plink.exe" -ESXiUsername "root" -ESXiPassword "foobar"
#>
    param(
        [Parameter(Mandatory=$false)][String]$VMHostName,
        [Parameter(Mandatory=$false)][String]$ClusterName,
        [Parameter(Mandatory=$false)][Boolean]$IncludeMicrocodeVerCheck=$false,
        [Parameter(Mandatory=$false)][String]$PlinkPath,
        [Parameter(Mandatory=$false)][String]$ESXiUsername,
        [Parameter(Mandatory=$false)][String]$ESXiPassword
    )

    if($ClusterName) {
        $cluster = Get-View -ViewType ClusterComputeResource -Property Name,Host -Filter @{"name"=$ClusterName}
        $vmhosts = Get-View $cluster.Host -Property Name,Config.FeatureCapability,Hardware.CpuFeature,Summary.Hardware,ConfigManager.ServiceSystem
    } elseif($VMHostName) {
        $vmhosts = Get-View -ViewType HostSystem -Property Name,Config.FeatureCapability,Hardware.CpuFeature,Summary.Hardware,ConfigManager.ServiceSystem -Filter @{"name"=$VMHostName}
    } else {
        $vmhosts = Get-View -ViewType HostSystem -Property Name,Config.FeatureCapability,Hardware.CpuFeature,Summary.Hardware,ConfigManager.ServiceSystem
    }

    # Merge of tables from https://kb.vmware.com/s/article/52345 and https://kb.vmware.com/s/article/52085
    $procSigUcodeTable = @(
	    [PSCustomObject]@{Name = "Sandy Bridge DT";  procSig = "0x000206a7"; ucodeRevFixed = "0x0000002d"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Sandy Bridge EP";  procSig = "0x000206d7"; ucodeRevFixed = "0x00000713"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Ivy Bridge DT";  procSig = "0x000306a9"; ucodeRevFixed = "0x0000001f"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Ivy Bridge EP";  procSig = "0x000306e4"; ucodeRevFixed = "0x0000042c"; ucodeRevSightings = "0x0000042a"}
	    [PSCustomObject]@{Name = "Ivy Bridge EX";  procSig = "0x000306e7"; ucodeRevFixed = "0x00000713"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Haswell DT";  procSig = "0x000306c3"; ucodeRevFixed = "0x00000024"; ucodeRevSightings = "0x00000023"}
	    [PSCustomObject]@{Name = "Haswell EP";  procSig = "0x000306f2"; ucodeRevFixed = "0x0000003c"; ucodeRevSightings = "0x0000003b"}
	    [PSCustomObject]@{Name = "Haswell EX";  procSig = "0x000306f4"; ucodeRevFixed = "0x00000011"; ucodeRevSightings = "0x00000010"}
	    [PSCustomObject]@{Name = "Broadwell H";  procSig = "0x00040671"; ucodeRevFixed = "0x0000001d"; ucodeRevSightings = "0x0000001b"}
	    [PSCustomObject]@{Name = "Broadwell EP/EX";  procSig = "0x000406f1"; ucodeRevFixed = "0x0b00002a"; ucodeRevSightings = "0x0b000025"}
	    [PSCustomObject]@{Name = "Broadwell DE";  procSig = "0x00050662"; ucodeRevFixed = "0x00000015"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Broadwell DE";  procSig = "0x00050663"; ucodeRevFixed = "0x07000012"; ucodeRevSightings = "0x07000011"}
	    [PSCustomObject]@{Name = "Broadwell DE";  procSig = "0x00050664"; ucodeRevFixed = "0x0f000011"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Broadwell NS";  procSig = "0x00050665"; ucodeRevFixed = "0x0e000009"; ucodeRevSightings = ""}
	    [PSCustomObject]@{Name = "Skylake H/S";  procSig = "0x000506e3"; ucodeRevFixed = "0x000000c2"; ucodeRevSightings = ""} # wasn't actually affected by Sightings, ucode just re-released
	    [PSCustomObject]@{Name = "Skylake SP";  procSig = "0x00050654"; ucodeRevFixed = "0x02000043"; ucodeRevSightings = "0x0200003A"}
	    [PSCustomObject]@{Name = "Kaby Lake H/S/X";  procSig = "0x000906e9"; ucodeRevFixed = "0x00000084"; ucodeRevSightings = "0x0000007C"}
	    [PSCustomObject]@{Name = "Zen EPYC";  procSig = "0x00800f12"; ucodeRevFixed = "0x08001227"; ucodeRevSightings = ""}
    )

    # Remote SSH commands for retrieving current ESXi host microcode version
    $plinkoptions = "-ssh -pw $ESXiPassword"
    $cmd = "vsish -e cat /hardware/cpu/cpuList/0 | grep `'Current Revision:`'"
    $remoteCommand = '"' + $cmd + '"'

    $results = @()
    foreach ($vmhost in $vmhosts | Sort-Object -Property Name) {
        $vmhostDisplayName = $vmhost.Name
        $cpuModelName = $($vmhost.Summary.Hardware.CpuModel -replace '\s+', ' ')

        $IBRSPass = $false
        $IBPBPass = $false
        $STIBPPass = $false

        $cpuFeatures = $vmhost.Config.FeatureCapability
        foreach ($cpuFeature in $cpuFeatures) {
            if($cpuFeature.key -eq "cpuid.IBRS" -and $cpuFeature.value -eq 1) {
                $IBRSPass = $true
            } elseif($cpuFeature.key -eq "cpuid.IBPB" -and $cpuFeature.value -eq 1) {
                $IBPBPass = $true
            } elseif($cpuFeature.key -eq "cpuid.STIBP" -and $cpuFeature.value -eq 1) {
                $STIBPPass = $true
            }
        }

        $vmhostAffected = $true
        if($IBRSPass -or $IBPBPass -or $STIBPass) {
           $vmhostAffected = $false
        }

        # Retrieve Microcode version if user specifies which unfortunately requires SSH access
        if($IncludeMicrocodeVerCheck -and $PlinkPath -ne $null -and $ESXiUsername -ne $null -and $ESXiPassword -ne $null) {
            $serviceSystem = Get-View $vmhost.ConfigManager.ServiceSystem
            $services = $serviceSystem.ServiceInfo.Service
            foreach ($service in $services) {
                if($service.Key -eq "TSM-SSH") {
                    $ssh = $service
                    break
                }
            }

            $command = "echo yes | " + $PlinkPath + " " + $plinkoptions + " " + $ESXiUsername + "@" + $vmhost.Name + " " + $remoteCommand

            if($ssh.Running){
                $plinkResults = Invoke-Expression -command $command
                $microcodeVersion = $plinkResults.split(":")[1]
            } else {
                $microcodeVersion = "SSHNeedsToBeEnabled"
            }
        } else {
            $microcodeVersion = "N/A"
        }

        #output from $vmhost.Hardware.CpuFeature is a binary string ':' delimited to nibbles
        #the easiest way I could figure out the hex conversion was to make a byte array
        $cpuidEAX = ($vmhost.Hardware.CpuFeature | Where-Object {$_.Level -eq 1}).Eax -Replace ":",""
        $cpuidEAXbyte = $cpuidEAX -Split "(?<=\G\d{8})(?=\d{8})"
        $cpuidEAXnibble = $cpuidEAX -Split "(?<=\G\d{4})(?=\d{4})"

        $cpuSignature = "0x" + $(($cpuidEAXbyte | Foreach-Object {[System.Convert]::ToByte($_, 2)} | Foreach-Object {$_.ToString("X2")}) -Join "")

        # https://software.intel.com/en-us/articles/intel-architecture-and-processor-identification-with-cpuid-model-and-family-numbers
        $ExtendedFamily = [System.Convert]::ToInt32($($cpuidEAXnibble[1] + $cpuidEAXnibble[2]), 2)
        $Family = [System.Convert]::ToInt32($cpuidEAXnibble[5], 2)

        # output now in decimal, not hex!
        $cpuFamily = $ExtendedFamily + $Family
        $cpuModel = [System.Convert]::ToByte($($cpuidEAXnibble[3] + $cpuidEAXnibble[6]), 2)
        $cpuStepping = [System.Convert]::ToByte($cpuidEAXnibble[7], 2)
               
        
        $intelSighting = "N/A"
        $goodUcode = "N/A"

        # check and compare ucode
        if ($IncludeMicrocodeVerCheck) {
         
            $intelSighting = $false
            $goodUcode = $false
            $matched = $false

            foreach ($cpu in $procSigUcodeTable) {
                if ($cpuSignature -eq $cpu.procSig) {
                    $matched = $true
                    if ($microcodeVersion -eq $cpu.ucodeRevSightings) {
                        $intelSighting = $true
                    } elseif ($microcodeVersion -as [int] -ge $cpu.ucodeRevFixed -as [int]) {
                        $goodUcode = $true
                    }
                }
            } 
            if (!$matched) {
                # CPU is not in procSigUcodeTable, check with BIOS vendor / Intel based procSig or FMS (dec) in output
                $goodUcode = "Unknown"
            }
        }

        $tmp = [pscustomobject] @{
            VMHost = $vmhostDisplayName;
            "CPU Model Name" = $cpuModelName;
            Family = $cpuFamily;
            Model = $cpuModel;
            Stepping = $cpuStepping;
            Microcode = $microcodeVersion;
            procSig = $cpuSignature;
            IBRSPresent = $IBRSPass;
            IBPBPresent = $IBPBPass;
            STIBPPresent = $STIBPPass;
            HypervisorAssistedGuestAffected = $vmhostAffected;
            "Good Microcode" = $goodUcode;
            IntelSighting = $intelSighting;
        }
        $results+=$tmp
    }
    $results | FT *
}
# Author: William Lam
# Blog: www.virtuallyghetto.com
# Description: Script to add a new VMDK w/the MultiWriter Flag enabled in vSphere 6.x
# Reference: ht﻿Function Get-VMCreationDate {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
    ===========================================================================
    .DESCRIPTION
        This function extract VM Creation Date using vSphere API (currently only available on VMware Cloud on AWS SDDCs)
        which it does by processing the HTML results found in the vSphere MOB. Once this functionality is available
        via the vSphere SDKs, this will be a simple 1-liner for PowerCLI and other vSphere SDKs
    .PARAMETER VMName
        The name of a VM to extract the creation date
    .PARAMETER vc_server
        The name of the VMWonAWS vCenter Server
    .PARAMETER vc_username
        The username of the VMWonAWS vCenter Server
    .PARAMETER VMName
        The password of the VMWonAWS vCenter Server
    .EXAMPLE
        Connect-VIServer -Server $vc_server -User $vc_username -Password $vc_password
        Get-VMCreationDate -vc_server $vc_server -vc_username $vc_username -vc_password $vc_password -vmname $vmname
#>
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [String]$vmname,
    [string]$vc_server,
    [String]$vc_username,
    [String]$vc_password
    )

    $vm = Get-VM -Name $vmname
    $vm_moref = (Get-View $vm).MoRef.Value
    $vm_moref = $vm_moref -replace "-","%2d"

    $secpasswd = ConvertTo-SecureString $vc_password -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($vc_username, $secpasswd)

    # vSphere MOB URL to private enableMethods
    $mob_url = "https://$vc_server/mob/?moid=$vm_moref&doPath=config"

# Ingore SSL Warnings
add-type -TypeDefinition  @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

    # Initial login to vSphere MOB using GET and store session using $vmware variable
    $results = Invoke-WebRequest -Uri $mob_url -SessionVariable vmware -Credential $credential -Method GET

    if($results.StatusCode -eq 200) {
        # Parsing HTML (ewww) from the vSphere MOB, using the vSphere SDKs once they are enabled for VMware Cloud on AWS will be simple 1-liner 
        $createDate = ($results.ParsedHtml.getElementsByTagName("TR") | where {$_.innerText -match "createDate"} | select innerText | ft -hide | Out-String).replace("createDatedateTime","").Replace("`"","").Trim()

        $creaeDateResults = [pscustomobject] @{
            Name = $vm.Name;
            CreateDate = $createDate;
        }
        return $creaeDateResults
    } else {
        Write-host "Failed to login to vSphere MOB"
        exit 1
    }

    # Logout out of vSphere MOB
    $mob_logout_url = "https://$vc_server/mob/logout"
    $logout = Invoke-WebRequest -Uri $mob_logout_url -WebSession $vmware -Method GET
}﻿Function Set-VMKeystrokes {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
    ===========================================================================
    .DESCRIPTION
        This function sends a series of character keystrokse to a particular VM
    .PARAMETER VMName
		The name of a VM to send keystrokes to
	.PARAMETER StringInput
		The string of characters to send to VM
	.PARAMETER DebugOn
		Enable debugging which will output input charcaters and their mappings
    .EXAMPLE
        Set-VMKeystrokes -VMName $VM -StringInput "root"
    .EXAMPLE
        Set-VMKeystrokes -VMName $VM -StringInput "root" -ReturnCarriage $true
    .EXAMPLE
        Set-VMKeystrokes -VMName $VM -StringInput "root" -DebugOn $true
#>
    param(
        [Parameter(Mandatory=$true)][String]$VMName,
        [Parameter(Mandatory=$true)][String]$StringInput,
        [Parameter(Mandatory=$false)][Boolean]$ReturnCarriage,
        [Parameter(Mandatory=$false)][Boolean]$DebugOn
    )

    # Map subset of USB HID keyboard scancodes
    # https://gist.github.com/MightyPork/6da26e382a7ad91b5496ee55fdc73db2
    $hidCharacterMap = @{
		"a"="0x04";
		"b"="0x05";
		"c"="0x06";
		"d"="0x07";
		"e"="0x08";
		"f"="0x09";
		"g"="0x0a";
		"h"="0x0b";
		"i"="0x0c";
		"j"="0x0d";
		"k"="0x0e";
		"l"="0x0f";
		"m"="0x10";
		"n"="0x11";
		"o"="0x12";
		"p"="0x13";
		"q"="0x14";
		"r"="0x15";
		"s"="0x16";
		"t"="0x17";
		"u"="0x18";
		"v"="0x19";
		"w"="0x1a";
		"x"="0x1b";
		"y"="0x1c";
		"z"="0x1d";
		"1"="0x1e";
		"2"="0x1f";
		"3"="0x20";
		"4"="0x21";
		"5"="0x22";
		"6"="0x23";
		"7"="0x24";
		"8"="0x25";
		"9"="0x26";
		"0"="0x27";
		"!"="0x1e";
		"@"="0x1f";
		"#"="0x20";
		"$"="0x21";
		"%"="0x22";
		"^"="0x23";
		"&"="0x24";
		"*"="0x25";
		"("="0x26";
		")"="0x27";
		"_"="0x2d";
		"+"="0x2e";
		"{"="0x2f";
		"}"="0x30";
		"|"="0x31";
		":"="0x33";
		"`""="0x34";
		"~"="0x35";
		"<"="0x36";
		">"="0x37";
		"?"="0x38";
		"-"="0x2d";
		"="="0x2e";
		"["="0x2f";
		"]"="0x30";
		"\"="0x31";
		"`;"="0x33";
		"`'"="0x34";
		","="0x36";
		"."="0x37";
		"/"="0x38";
		" "="0x2c";
    }

    $vm = Get-View -ViewType VirtualMachine -Filter @{"Name"="^$($VMName)$"}

	# Verify we have a VM or fail
    if(!$vm) {
        Write-host "Unable to find VM $VMName"
        return
    }

    $hidCodesEvents = @()
    foreach($character in $StringInput.ToCharArray()) {
        # Check to see if we've mapped the character to HID code
        if($hidCharacterMap.ContainsKey([string]$character)) {
            $hidCode = $hidCharacterMap[[string]$character]

            $tmp = New-Object VMware.Vim.UsbScanCodeSpecKeyEvent

            # Add leftShift modifer for capital letters and/or special characters
            if( ($character -cmatch "[A-Z]") -or ($character -match "[!|@|#|$|%|^|&|(|)|_|+|{|}|||:|~|<|>|?|*]") ) {
                $modifer = New-Object Vmware.Vim.UsbScanCodeSpecModifierType
                $modifer.LeftShift = $true
                $tmp.Modifiers = $modifer
            }

            # Convert to expected HID code format
            $hidCodeHexToInt = [Convert]::ToInt64($hidCode,"16")
            $hidCodeValue = ($hidCodeHexToInt -shl 16) -bor 0007

            $tmp.UsbHidCode = $hidCodeValue
            $hidCodesEvents+=$tmp

            if($DebugOn) {
                Write-Host "Character: $character -> HIDCode: $hidCode -> HIDCodeValue: $hidCodeValue"
            }
        } else {
            Write-Host "The following character `"$character`" has not been mapped, you will need to manually process this character"
            break
        }
    }

    # Add return carriage to the end of the string input (useful for logins or executing commands)
    if($ReturnCarriage) {
        # Convert return carriage to HID code format
        $hidCodeHexToInt = [Convert]::ToInt64("0x28","16")
        $hidCodeValue = ($hidCodeHexToInt -shl 16) + 7

        $tmp = New-Object VMware.Vim.UsbScanCodeSpecKeyEvent
        $tmp.UsbHidCode = $hidCodeValue
        $hidCodesEvents+=$tmp
    }

    # Call API to send keystrokes to VM
    $spec = New-Object Vmware.Vim.UsbScanCodeSpec
    $spec.KeyEvents = $hidCodesEvents
    Write-Host "Sending keystrokes to $VMName ...`n"
    $results = $vm.PutUsbScanCodes($spec)
}
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
}<#
.SYNOPSIS Script to deploy vRealize Network Insight (vRNI) 3.2 Platform + Proxy VM
.NOTES  Author:  William Lam
.NOTES  Site:    www.virtuallyghetto.com
.NOTES  Reference: http://www.virtuallyghetto.com/2016/12/automated-deployment-and-setup-of-vrealize-network-insight-vrni.html
#>

# Path to vRNI OVAs
﻿$vRNIPlatformOVA = "C:\Users\primp\Desktop\VMWare-vRealize-Networking-insight-3.2.0.1480511973-platform.ova"
$vRNIProxyOVA = "C:\Users\primp\Desktop\VMWare-vRealize-Networking-insight-3.2.0.1480511973-proxy.ova"

# vRNI License Key
$vRNILicenseKey = ""

# vRNI Platform VM Config
$vRNIPlatformVMName = "vRNI-Platform-3.2"
$vRNIPlatformIPAddress = "172.30.0.199"
$vRNIPlatformNetmask = "255.255.255.0"
$vRNIPlatformGateway = "172.30.0.1"

# vRNI Proxy VM Config
$vRNIProxyVMName = "vRNI-Proxy-3.2"
$vRNIProxyIPAddress = "172.30.0.201"
$vRNIProxyNetmask = "255.255.255.0"
$vRNIProxyGateway = "172.30.0.201"

# vRNI Deployment Settings
$DeploymentSize = "medium"
$DNS = "172.30.0.100"
$DNSDomain = "primp-industries.com"
$NTPServer = "172.30.0.100"

$VMCluster = "Primp-Cluster"
$VMDatastore = "himalaya-local-SATA-re4gp4T:storage"
$VMNetwork = "access333"

### DO NOT EDIT BEYOND HERE ###

Function My-Logger {
    param(
    [Parameter(Mandatory=$true)]
    [String]$message
    )

    $timeStamp = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"

    Write-Host -NoNewline -ForegroundColor White "[$timestamp]"
    Write-Host -ForegroundColor Green " $message"
}

$StartTime = Get-Date

$hash = @{licenseKey = $vRNILicenseKey}
$json = $hash | ConvertTo-Json

$location = Get-Cluster $VMCluster
$datastore = Get-Datastore -Name $VMDatastore
$vmhost = Get-VMHost
$network = Get-VirtualPortGroup -Name $VMNetwork -VMHost $vmhost[0]
$vRNIPlatformOVFConfig = Get-OvfConfiguration $vRNIPlatformOVA
$vRNIProxyOVFConfig = Get-OvfConfiguration $vRNIProxyOVA

$vRNIPlatformOVFConfig.DeploymentOption.Value = $DeploymentSize
$vRNIPlatformOVFConfig.NetworkMapping.Vlan256_corp_2.Value = $VMNetwork
$vRNIPlatformOVFConfig.Common.IP_Address.Value = $vRNIPlatformIPAddress
$vRNIPlatformOVFConfig.Common.Netmask.Value = $vRNIPlatformNetmask
$vRNIPlatformOVFConfig.Common.Default_Gateway.Value = $vRNIPlatformGateway
$vRNIPlatformOVFConfig.Common.DNS.Value = $DNS
$vRNIPlatformOVFConfig.Common.Domain_Search.Value = $DNSDomain
$vRNIPlatformOVFConfig.Common.NTP.Value = $NTPServer

$vRNIProxyOVFConfig.DeploymentOption.Value = $DeploymentSize
$vRNIProxyOVFConfig.NetworkMapping.Vlan256_corp_2.Value = $VMNetwork
$vRNIProxyOVFConfig.Common.IP_Address.Value = $vRNIProxyIPAddress
$vRNIProxyOVFConfig.Common.Netmask.Value = $vRNIProxyNetmask
$vRNIProxyOVFConfig.Common.Default_Gateway.Value = $vRNIProxyGateway
$vRNIProxyOVFConfig.Common.DNS.Value = $DNS
$vRNIProxyOVFConfig.Common.Domain_Search.Value = $DNSDomain
$vRNIProxyOVFConfig.Common.NTP.Value = $NTPServer

My-Logger "Deploying vRNI Platform OVA ..."
$vRNIPlatformVM = Import-VApp -OvfConfiguration $vRNIPlatformOVFConfig -Source $vRNIPlatformOVA -Name $vRNIPlatformVMName -VMHost $vmhost -Datastore $datastore -DiskStorageFormat thin -Location $location

My-Logger "Starting vRNI Platform VM ..."
Start-VM -VM $vRNIPlatformVM -Confirm:$false | Out-Null

My-Logger "Checking to see if vRNI Platform VM is ready ..."
while(1) {
    try {
        $results = Invoke-WebRequest -Uri https://$vRNIPlatformIPAddress/#license/step/1 -Method GET
        if($results.StatusCode -eq 200) {
            break
        }
    }
    catch {
        My-Logger "vRNI Platform is not ready, sleeping for 120 seconds ..."
        sleep 120
    }
}

# vRNI URLs for configuration
$validateURL = "https://$vRNIPlatformIPAddress/api/management/licensing/validate"
$activateURL = "https://$vRNIPlatformIPAddress/api/management/licensing/activate"
$proxySecretGenURL = "https://$vRNIPlatformIPAddress/api/management/nodes"

My-Logger "Verifying vRNI License Key ..."
$results = Invoke-WebRequest -Uri $validateURL -SessionVariable vmware -Method POST -ContentType "application/json" -Body $json
if($results.StatusCode -eq 200) {
    My-Logger "Activating vRNI License Key ..."
    $results = Invoke-WebRequest -Uri $activateURL -WebSession $vmware -Method POST -ContentType "application/json" -Body $json
    if($results.StatusCode -eq 200) {
        My-Logger "Generating vRNI Proxy Shared Secret ..."
        $results = Invoke-WebRequest -Uri $proxySecretGenURL -WebSession $vmware -Method POST -ContentType "application/json"
        if($results.StatusCode -eq 200) {
            $cleanedUpResults = $results.ParsedHtml.body.innertext.split("`n").replace("`"","") | ? {$_.trim() -ne ""}
            $lString = $cleanedUpResults.replace("{status:true,statusCode:{code:0,codeStr:OK},message:Proxy Key Generated,data:","")
            $vRNIPlatformSharedSecret = $lString.replace("}","")

            if($vRNIPlatformSharedSecret -ne $null) {
                # Update OVF Property w/shared secret
                $vRNIProxyOVFConfig.Common.Proxy_Shared_Secret.Value = $vRNIPlatformSharedSecret

                My-Logger "Deploying vRNI Proxy OVA w/Platform shared secret  ..."
                $vRNIProxyVM = Import-VApp -OvfConfiguration $vRNIProxyOVFConfig -Source $vRNIProxyOVA -Name $vRNIProxyVMName -VMHost $vmhost -Datastore $datastore -DiskStorageFormat thin -Location $location

                My-Logger "Starting vRNI Proxy VM ..."
                Start-VM -VM $vRNIProxyVM -Confirm:$false | Out-Null

                My-Logger "Waiting for vRNI Proxy VM to be detected by vRNI Platform VM ..."
                $notDectected = $true
                while($notDectected) {
                    $results = Invoke-WebRequest -Uri $proxySecretGenURL -WebSession $vmware -Method GET -ContentType "application/json"
                    $nodes = $results.Content | ConvertFrom-Json
                    if($nodes.Count -eq 2) {
                        foreach ($node in $nodes) {
                            if($node.ipAddress -eq "$vRNIProxyIPAddress" -and $node.healthStatus -eq "HEALTHY") {
                                My-Logger "vRNI Proxy VM detected"
                                $notDectected = $false
                            }
                        }
                    } else {
                        sleep 60
                        My-Logger "Still waiting for vRNI Proxy VM, sleeping for 60 seconds ..."
                    }
                }
            } else {
                Write-Host -ForegroundColor Red "Failed to retrieve vRNI Platform Shared Secret Key ..."
                break
            }
        }
    } else {
        Write-Host -ForegroundColor Red "Failed to activate vRNI License Key ..."
        break
    }
} else {
    Write-Host -ForegroundColor Red "Failed to validate vRNI License Key ... "
}

$EndTime = Get-Date
$duration = [math]::Round((New-TimeSpan -Start $StartTime -End $EndTime).TotalMinutes,2)

My-Logger "vRealize Network Insight Deployment Complete!"
My-Logger "         Login to https://$vRNIPlatformIPAddress using"
My-Logger "            Username: admin@local"
My-Logger "            Password: admin"
My-Logger "StartTime: $StartTime"
My-Logger "  EndTime: $EndTime"
My-Logger " Duration: $duration minutes"
﻿<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function demonstrates the use of vSAN Management API to retrieve
        the exact same information provided by the RVC command "vsan.check_limits"
        Please see http://www.virtuallyghetto.com/2017/06/how-to-convert-vsan-rvc-commands-into-powercli-andor-other-vsphere-sdks.html for more details
    .PARAMETER Cluster
        The name of a vSAN Cluster
    .EXAMPLE
        Get-VsanLimits -Cluster VSAN-Cluster
#>
Function Get-VsanLimits {
    param(
        [Parameter(Mandatory=$true)][String]$Cluster
    )

    $vmhosts = (Get-Cluster -Name $Cluster | Get-VMHost | Sort-Object -Property Name)

    $limitsResults = @()
    foreach($vmhost in $vmhosts) {
        $connectionState = $vmhost.ExtensionData.Runtime.runtime.connectionState
        $vsanEnabled = (Get-View $vmhost.ExtensionData.ConfigManager.vsanSystem).config.enabled

        if($connectionState -ne "Connected" -and $vsanEnabled -ne $true) {
            break
        }

        $vsanInternalSystem = Get-View $vmhost.ExtensionData.ConfigManager.vsanInternalSystem

        # Fetch RDT Information
        $resultsForRdtLsomDom = $vsanInternalSystem.QueryVsanStatistics(@('rdtglobal','lsom-node','lsom','dom','dom-objects-counts'))
        $jsonFroRdtLsomDom = $resultsForRdtLsomDom | ConvertFrom-Json

        # Process RDT Data Start #
        $rdtAssocs = $jsonFroRdtLsomDom.'rdt.globalinfo'.assocCount.ToString() + "/" + $jsonFroRdtLsomDom.'rdt.globalinfo'.maxAssocCount.ToString()
        $rdtSockets = $jsonFroRdtLsomDom.'rdt.globalinfo'.socketCount.ToString() + "/" + $jsonFroRdtLsomDom.'rdt.globalinfo'.maxSocketCount.ToString()
        $rdtClients = 0
        foreach($line in $jsonFroRdtLsomDom.'dom.clients' | Get-Member) {
            # crappy way to iterate through keys ...
            if($($line.Name) -ne "Equals" -and $($line.Name) -ne "GetHashCode" -and $($line.Name) -ne "GetType" -and $($line.Name) -ne "ToString") {
                $rdtClients++
            }
        }
        $rdtOwners = 0
        foreach($line in $jsonFroRdtLsomDom.'dom.owners.count' | Get-Member) {
            # crappy way to iterate through keys ...
            if($($line.Name) -ne "Equals" -and $($line.Name) -ne "GetHashCode" -and $($line.Name) -ne "GetType" -and $($line.Name) -ne "ToString") {
                $rdtOwners++
            }
        }
        # Process RDT Data End #

        # Fetch Component information
        $resultsForComponents = $vsanInternalSystem.QueryPhysicalVsanDisks(@('lsom_objects_count','uuid','isSsd','capacity','capacityUsed'))
        $jsonForComponents = $resultsForComponents | ConvertFrom-Json

        # Process Component Data Start #
        $vsanUUIDs = @{}
        $vsanDiskMgmtSystem = Get-VsanView -Id VimClusterVsanVcDiskManagementSystem-vsan-disk-management-system
        $diskGroups = $vsanDiskMgmtSystem.QueryDiskMappings($vmhost.ExtensionData.Moref)
        foreach($diskGroup in $diskGroups) {
            $mappings = $diskGroup.mapping
            foreach($mapping in $mappings ) {
                $ssds = $mapping.ssd
                $nonSsds = $mapping.nonSsd

                foreach($ssd in $ssds ) {
                    $vsanUUIDs.add($ssd.vsanDiskInfo.vsanUuid,$ssd)
                }

                foreach($nonSsd in $nonSsds ) {
                    $vsanUUIDs.add($nonSsd.vsanDiskInfo.vsanUuid,$nonSsd)
                }
            }
        }
        $maxComponents = $jsonFroRdtLsomDom.'lsom.node'.numMaxComponents

        $diskString = ""
        $hostComponents = 0
        foreach($line in $jsonForComponents | Get-Member) {
            # crappy way to iterate through keys ...
            if($($line.Name) -ne "Equals" -and $($line.Name) -ne "GetHashCode" -and $($line.Name) -ne "GetType" -and $($line.Name) -ne "ToString") {
                if($vsanUUIDs.ContainsKey($line.Name)) {
                    $numComponents = ($jsonFroRdtLsomDom.'lsom.disks'.$($line.Name).info.numComp).toString()
                    $maxCoponents = ($jsonFroRdtLsomDom.'lsom.disks'.$($line.Name).info.maxComp).toString()
                    $hostComponents += $jsonForComponents.$($line.Name).lsom_objects_count
                    $usage = ($jsonFroRdtLsomDom.'lsom.disks'.$($line.Name).info.capacityUsed * 100) / $jsonFroRdtLsomDom.'lsom.disks'.$($line.Name).info.capacity
                    $usage = [math]::ceiling($usage)

                    $diskString+=$vsanUUIDs.$($line.Name).CanonicalName + ": " + $usage + "% Components: " + $numComponents + "/" + $maxCoponents + "`n"
                }
            }
        }
        # Process Component Data End #

        # Store output into an object
        $hostLimitsResult = [pscustomobject] @{
            Host = $vmhost.Name
            RDT = "Assocs: " + $rdtAssocs + "`nSockets: " + $rdtSockets + "`nClients: " + $rdtClients + "`nOwners: " + $rdtOwners
            Disks = "Components: " + $hostComponents + "/" + $maxComponents + "`n" + $diskString
        }
        $limitsResults+=$hostLimitsResult
    }
    # Display output
    $limitsResults | Format-Table -Wrap
}﻿# Author: William Lam
# Website: www.virtuallyghetto.com
# Product: VMware vSphere / VSAN
# Description: VSAN Flash/MD capacity report
# Reference: http://www.virtuallyghetto.com/2014/04/vsan-flashmd-capacity-reporting.html

$vcName = ""
$vcenter = Connect-VIServer $vcname -WarningAction SilentlyContinue

$vsanMaxConfigInfo = @()

$clusviews = Get-View -ViewType ClusterComputeResource -Property Name,ConfigurationEx,Host
foreach ($cluster in $clusviews) {
	if($cluster.ConfigurationEx.VsanConfigInfo.Enabled) {
		$vmhosts = $cluster.Host
        foreach ($vmhost in $vmhosts | Sort-Object -Property Name) {
			$vmhostView = Get-View $vmhost -Property Name,ConfigManager.VsanSystem,ConfigManager.VsanInternalSystem	
			$vsanSys = Get-View -Id $vmhostView.ConfigManager.VsanSystem
			$vsanIntSys = Get-View -Id $vmhostView.ConfigManager.VsanInternalSystem
		
			$vsanProps = @("owner","uuid","isSsd","capacity","capacityUsed","capacityReserved")
			$results = $vsanIntSys.QueryPhysicalVsanDisks($vsanProps)
			$vsanStatus = $vsanSys.QueryHostStatus()
				
			$json = $results | ConvertFrom-Json
			foreach ($line in $json | Get-Member) {
				# ensure owner is owned by ESXi host
				if($vsanStatus.NodeUuid -eq $json.$($line.Name).owner) {
					if($json.$($line.Name).isSsd) {
						$totalSsdCapacity += $json.$($line.Name).capacity
						$totalSsdCapacityUsed += $json.$($line.Name).capacityUsed
						$totalSsdCapacityReserved += $json.$($line.Name).capacityReserved
					} else {
						$totalMdCapacity += $json.$($line.Name).capacity
						$totalMdCapacityUsed += $json.$($line.Name).capacityUsed
						$totalMdCapacityReserved += $json.$($line.Name).capacityReserved
					}				
				}
			}
		}
		$totalSsdCapacityReservedPercent = [int]($totalSsdCapacityReserved / $totalSsdCapacity * 100)
		$totalSsdCapacityUsedPercent = [int]($totalSsdCapacityUsed / $totalSsdCapacity * 100)
		$totalMdCapacityReservedPercent = [int]($totalMdCapacityReserved / $totalMdCapacity * 100)
		$totalMdCapacityUsedPercent = [int]($totalMdCapacityUsed / $totalMdCapacity * 100)
		
		$Details = "" |Select VSANCluster, TotalSsdCapacity, TotalSsdCapacityReserved, TotalSsdCapacityUsed,TotalSsdCapacityReservedPercent, TotalSsdCapacityUsedPercent, TotalMdCapacity, TotalMdCapacityReserved, TotalMdCapacityUsed, TotalMdCapacityReservedPercent, TotalMdCapacityUsedPercent
		$Details.VSANCluster = $cluster.Name + "`n"
		$Details.TotalSsdCapacity = [math]::round($totalSsdCapacity /1GB,2).ToString() + " GB"
		$Details.TotalSsdCapacityReserved = [math]::round($totalSsdCapacityReserved /1GB,2).ToString() + " GB"
		$Details.TotalSsdCapacityUsed = [math]::round($totalSsdCapacityUsed /1GB,2).ToString() + " GB"
		$Details.TotalSsdCapacityReservedPercent = $totalSsdCapacityReservedPercent.ToString() + "%"
		$Details.TotalSsdCapacityUsedPercent = $totalSsdCapacityUsedPercent.ToString() + "%`n"
		$Details.TotalMdCapacity = [math]::round($totalMdCapacity /1GB,2).ToString() + " GB"
		$Details.TotalMdCapacityReserved = [math]::round($totalMdCapacityReserved /1GB,2).ToString() + " GB"
		$Details.TotalMdCapacityUsed = [math]::round($totalMdCapacityUsed /1GB,2).ToString() + " GB"
		$Details.TotalMdCapacityReservedPercent = $totalMdCapacityReservedPercent.ToString() + "%"
		$Details.TotalMdCapacityUsedPercent = $totalMdCapacityUsedPercent.ToString() + "%"
		$vsanMaxConfigInfo += $Details
		
		$totalSsdCapacity = 0
		$totalSsdCapacityReserved = 0
		$totalSsdCapacityUsed = 0
		$totalSsdCapacityReservedPercent = 0
		$totalSsdCapacityUsedPercent = 0
		$totalMdCapacity = 0
		$totalMdCapacityReserved = 0
		$totalMdCapacityUsed = 0
		$totalMdCapacityReservedPercent = 0
		$totalMdCapacityUsedPercent = 0
	}
}

$vsanMaxConfigInfo

#Disconnect from vCenter
Disconnect-VIServer $vcenter -Confirm:$false
﻿Function Get-VSANHclInfo {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function demonstrates the use of vSAN Management API to retrieve
        the last time vSAN HCL was updated + HCL Health if HCL DB > 90days
    .PARAMETER Cluster
        The name of a vSAN Cluster
    .EXAMPLE
        Get-VSANHclInfo -Cluster Palo-Alto
#>
    param(
        [Parameter(Mandatory=$true)][String]$Cluster
    )

    # Scope query within vSAN/vSphere Cluster
    $clusterView = Get-Cluster -Name $Cluster -ErrorAction SilentlyContinue
    if($clusterView) {
        $clusterMoref = $clusterView.ExtensionData.MoRef
    } else {
        Write-Host -ForegroundColor Red "Unable to find vSAN Cluster $cluster ..."
        break
    }
    
    $vchs = Get-VsanView -Id VsanVcClusterHealthSystem-vsan-cluster-health-system
    $results = $vchs.VsanVcClusterGetHclInfo($clusterMoref,$null,$null,$null)
    $results | Select HclDbLastUpdate, HclDbAgeHealth
}﻿Function Get-VSANHealthChecks {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retreives all available vSAN Health Checks
    .PARAMETER Cluster
        The name of a vSAN Cluster
    .EXAMPLE
        Get-VSANHealthChecks
#>
    $vchs = Get-VSANView -Id "VsanVcClusterHealthSystem-vsan-cluster-health-system"
    $vchs.VsanQueryAllSupportedHealthChecks() | Select TestId, TestName | Sort-Object -Property TestId
}

Function Get-VSANSilentHealthChecks {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retreives the list of vSAN Health CHecks that have been silenced
    .PARAMETER Cluster
        The name of a vSAN Cluster
    .EXAMPLE
        Get-VSANSilentHealthChecks -Cluster VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$Cluster
    )
    $vchs = Get-VSANView -Id "VsanVcClusterHealthSystem-vsan-cluster-health-system"
    $cluster_view = (Get-Cluster -Name $Cluster).ExtensionData.MoRef
    $results = $vchs.VsanHealthGetVsanClusterSilentChecks($cluster_view)

    Write-Host "`nvSAN Health Checks Currently Silenced:`n"
    $results
}

Function Set-VSANSilentHealthChecks {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retreives the vSAN software version for both VC/ESXi
    .PARAMETER Cluster
        The name of a vSAN Cluster
    .PARAMETER Test
        The list of vSAN Health CHeck IDs to silence or re-enable
    .EXAMPLE
        Set-VSANSilentHealthChecks -Cluster VSAN-Cluster -Test controlleronhcl -Disable
    .EXAMPLE
        Set-VSANSilentHealthChecks -Cluster VSAN-Cluster -Test controlleronhcl,controllerfirmware -Disable
    .EXAMPLE
        Set-VSANSilentHealthChecks -Cluster VSAN-Cluster -Test controlleronhcl -Enable
    .EXAMPLE
        Set-VSANSilentHealthChecks -Cluster VSAN-Cluster -Test controlleronhcl,controllerfirmware -Enable
#>
    param(
        [Parameter(Mandatory=$true)][String]$Cluster,
        [Parameter(Mandatory=$true)][String[]]$Test,
        [Switch]$Enabled,
        [Switch]$Disabled
    )
    $vchs = Get-VSANView -Id "VsanVcClusterHealthSystem-vsan-cluster-health-system"
    $cluster_view = (Get-Cluster -Name $Cluster).ExtensionData.MoRef

    if($Enabled) {
        $vchs.VsanHealthSetVsanClusterSilentChecks($cluster_view,$null,$Test)
    } else {
        $vchs.VsanHealthSetVsanClusterSilentChecks($cluster_view,$Test,$null)
    }
}Function Get-VsanHealthSummary {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function demonstrates the use of vSAN Management API to retrieve
        the same information provided by the RVC command "vsan.health.health_summary"
    .PARAMETER Cluster
        The name of a vSAN Cluster
    .EXAMPLE
        Get-VsanHealthSummary -Cluster VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$Cluster
    )
    $vchs = Get-VSANView -Id "VsanVcClusterHealthSystem-vsan-cluster-health-system"
    $cluster_view = (Get-Cluster -Name $Cluster).ExtensionData.MoRef
    $results = $vchs.VsanQueryVcClusterHealthSummary($cluster_view,$null,$null,$true,$null,$null,'defaultView')
    $healthCheckGroups = $results.groups

    $healthCheckResults = @()
    foreach($healthCheckGroup in $healthCheckGroups) {
        switch($healthCheckGroup.GroupHealth) {
            red {$healthStatus = "error"}
            yellow {$healthStatus = "warning"}
            green {$healthStatus = "passed"}
        }
        $healtCheckGroupResult = [pscustomobject] @{
            HealthCHeck = $healthCheckGroup.GroupName
            Result = $healthStatus
        }
        $healthCheckResults+=$healtCheckGroupResult
    }
    Write-Host "`nOverall health:" $results.OverallHealth "("$results.OverallHealthDescription")"
    $healthCheckResults
}﻿Function Set-VsanLargeClusterAdvancedSetting {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function updates the ESXi Advanced Settings for enabling large vSAN Clusters
        for ESXi hosts running 5.5, 6.0 & 6.5 
    .PARAMETER ClusterName
        Name of the vSAN Cluster to update ESXi Advanced Settings for large vSAN Clusters
    .EXAMPLE
        Set-VsanLargeClusterAdvancedSetting -ClusterName VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$ClusterName
    )

    $cluster = Get-Cluster -Name $ClusterName -ErrorAction SilentlyContinue
    if($cluster -eq $null) {
        Write-Host -ForegroundColor Red "Error: Unable to find vSAN Cluster $ClusterName ..."
        break 
    }

    foreach ($vmhost in ($cluster | Get-VMHost)) {
        Write-Host "Updating Host:" $vmhost.name "..."
        # vSAN 6.x+ https://kb.vmware.com/kb/2110081
        if($vmhost.Version -eq "6.5.0") {
            Get-AdvancedSetting -Entity $vmhost -Name "VSAN.goto11" | Set-AdvancedSetting -Value 1 -Confirm:$false
            Get-AdvancedSetting -Entity $vmhost -Name "Net.TcpipHeapMax" | Set-AdvancedSetting -Value 1024 -Confirm:$false
        # vSAN 6.x+ https://kb.vmware.com/kb/2110081
        } elseif($vmhost.Version -eq "6.0.0") {
            Get-AdvancedSetting -Entity $vmhost -Name "VSAN.goto11" | Set-AdvancedSetting -Value 1 -Confirm:$false
            Get-AdvancedSetting -Entity $vmhost -Name "Net.TcpipHeapMax" | Set-AdvancedSetting -Value 1024 -Confirm:$false
            Get-AdvancedSetting -Entity $vmhost -Name "CMMDS.clientLimit" | Set-AdvancedSetting -Value 65 -Confirm:$false
        # vSAN 5.5 https://kb.vmware.com/kb/2073930
        } elseif($vmhost.Version -eq "5.5.0") {
            Get-AdvancedSetting -Entity $vmhost -Name "CMMDS.goto11" | Set-AdvancedSetting -Value 1 -Confirm:$false
        } else {
            Write-Host "$vmhost.Version is not a supported version for this script"
        }
    }
}

Function Get-VsanLargeClusterAdvancedSetting {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retrieves the ESXi Advanced Settings for enabling large vSAN Clusters
        for ESXi hosts running 5.5, 6.0 & 6.5 
    .PARAMETER ClusterName
        Name of the vSAN Cluster to update ESXi Advanced Settings for large vSAN Clusters
    .EXAMPLE
        Get-VsanLargeClusterAdvancedSetting -ClusterName VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$ClusterName
    )

    $cluster = Get-Cluster -Name $ClusterName -ErrorAction SilentlyContinue
    if($cluster -eq $null) {
        Write-Host -ForegroundColor Red "Error: Unable to find vSAN Cluster $ClusterName ..."
        break 
    }

    foreach ($vmhost in ($cluster | Get-VMHost)) {
        Write-Host "Host:" $vmhost.name "..."
        if($vmhost.Version -eq "6.5.0") {
            Get-AdvancedSetting -Entity $vmhost -Name "VSAN.goto11"
            Get-AdvancedSetting -Entity $vmhost -Name "Net.TcpipHeapMax"
        } elseif($vmhost.Version -eq "6.0.0") {
            Get-AdvancedSetting -Entity $vmhost -Name "VSAN.goto11"
            Get-AdvancedSetting -Entity $vmhost -Name "Net.TcpipHeapMax"
            Get-AdvancedSetting -Entity $vmhost -Name "CMMDS.clientLimit"
        } elseif($vmhost.Version -eq "5.5.0") {
            Get-AdvancedSetting -Entity $vmhost -Name "CMMDS.goto11"
        } else {
            Write-Host "$vmhost.Version is not a supported version for this script"
        }
    }
}﻿Function Get-VsanObjectDistribution {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function provides an overview of the distribution of vSAN Objects across
        a given vSAN Cluster
    .PARAMETER Cluster
        The name of a vSAN Cluster
    .PARAMETER ShowvSANID
        Outputs the vSAN UUID of the SSD Device in Diskgroup
    .PARAMETER ShowDiskID
        Outputs the Disk Canoical ID of the SSD Device in Diskgroup
    .EXAMPLE
        Get-VsanObjectDistribution -ClusterName "VSAN-Cluster-6.5"
    .EXAMPLE
        Get-VsanObjectDistribution -ClusterName "VSAN-Cluster-6.5" -ShowDiskID $true
#>
    param(
        [Parameter(Mandatory=$true)][String]$ClusterName,
        [Parameter(Mandatory=$false)][Boolean]$ShowvSANID,
        [Parameter(Mandatory=$false)][Boolean]$ShowDiskID
    )

    Function Get-VSANDiskMapping {
        param(
            [Parameter(Mandatory=$true)]$vmhost
        )
        $vsanSystem = Get-View ($vmhost.ExtensionData.ConfigManager.VsanSystem)
        $vsanDiskMappings = $vsanSystem.config.storageInfo.diskMapping

        $diskGroupCount = 1
        $diskGroupObjectCount = 0
        $diskGroupObjectSize = 0
        $diskGroupMappings = @{}
        foreach ($disk in $vsanDiskMappings) {
            $hdds = $disk.nonSsd
            foreach ($hdd in $hdds) {
                $diskHDD = $hdd.VsanDiskInfo.VsanUuid
                if($diskInfo[$diskHDD]) {
                    $diskGroupObjectCount += $diskInfo[$diskHDD].totalComponents
                    $diskGroupObjectSize += $diskInfo[$diskHDD].used
                    $global:clusterTotalObjects += $diskInfo[$diskHDD].totalComponents
                    $global:clusterTotalObjectSize += $diskInfo[$diskHDD].used
                }
            }
            $diskGroupObj = [pscustomobject] @{
                numObjects = $diskGroupObjectCount;
                used = $diskGroupObjectSize;
                vsanID = $disk.Ssd.VsanDiskInfo.VsanUuid;
                diskID = $disk.Ssd.canonicalName;
            }
            $diskGroupMappings.add($diskGroupCount,$diskGroupObj)

            $diskGroupObjectCount = 0
            $diskGroupObjectSize = 0
            $diskGroupCount+=1
        }
        $global:clusterResults.add($vmhost.name,$diskGroupMappings)
    }

    Function BuildDiskInfo {
        $randomVmhost = Get-Cluster -Name $ClusterName | Get-VMHost | Select -First 1
        $vsanIntSys = Get-View ($randomVmhost.ExtensionData.ConfigManager.VsanInternalSystem)
        $results = $vsanIntSys.QueryPhysicalVsanDisks($null)
        $json = $results | ConvertFrom-Json


        foreach ($line in $json | Get-Member -MemberType NoteProperty) {
            $tmpObj = [pscustomobject] @{
                totalComponents = $json.$($line.Name).numTotalComponents
                dataComponents = $json.$($line.Name).numDataComponents
                witnessComponents = ($json.$($line.Name).numTotalComponents - $json.$($line.Name).numDataComponents)
                capacity = $json.$($line.Name).capacity
                used = $json.$($line.Name).physCapacityUsed
            }
            $diskInfo.Add($json.$($line.Name).uuid,$tmpObj)
        }
    }

    $cluster = Get-Cluster -Name $ClusterName -ErrorAction SilentlyContinue
    if($cluster -eq $null) {
        Write-Host -ForegroundColor Red "Error: Unable to find vSAN Cluster $ClusterName ..."
        break
    }

    $global:clusterResults = @{}
    $global:clusterTotalObjects =  0
    $global:clusterTotalObjectSize = 0
    $diskInfo = @{}
    BuildDiskInfo

    foreach ($vmhost in $cluster | Get-VMHost) {
        Get-VSANDiskMapping -vmhost $vmhost
    }

    Write-Host "`nTotal vSAN Components: $global:clusterTotalObjects"
    $size = [math]::Round(($global:clusterTotalObjectSize / 1GB),2)
    Write-Host "Total vSAN Components Size: $size GB"

    foreach ($vmhost in $global:clusterResults.keys | Sort-Object) {
        Write-Host "`n"$vmhost
        foreach ($diskgroup in $global:clusterResults[$vmhost].keys | Sort-Object) {
            if($ShowvSANID) {
                $diskID = $clusterResults[$vmhost][$diskgroup].vsanID
            } else {
                $diskID = $clusterResults[$vmhost][$diskgroup].diskID
            }

            Write-Host "`tDiskgroup $diskgroup (SSD: $diskID)"

            $numbOfObjects = $clusterResults[$vmhost][$diskgroup].numObjects
            $objPercentage = [math]::Round(($numbOfObjects / $global:clusterTotalObjects) * 100,2)
            Write-host "`t`tComponents: $numbOfObjects ($objPercentage%)"

            $objectsUsed = $clusterResults[$vmhost][$diskgroup].used
            $objectsUsedRounded = [math]::Round(($clusterResults[$vmhost][$diskgroup].used / 1GB),2)
            $usedPertcentage =[math]::Round(($objectsUsed / $global:clusterTotalObjectSize) * 100,2)
            Write-host "`t`tSize: $objectsUsedRounded GB ($usedPertcentage%)"
        }
    }
}﻿Function Get-VSANPerformanceEntityType {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retreives all available vSAN Performance Metric Entity Types
    .EXAMPLE
        Get-VSANPerformanceEntityType
#>
    $vpm = Get-VSANView -Id "VsanPerformanceManager-vsan-performance-manager"
    $entityTypes = $vpm.VsanPerfGetSupportedEntityTypes()

    foreach ($entityType in $entityTypes | Sort-Object -Property Name) {
        $entityType.Name
    }
}

Function Get-VSANPerformanceEntityMetric {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retreives all vSAN Performance Metrics for a given Entity Type
    .PARAMETER EntityType
        The name of the vSAN Performance Entity Type you wish to retrieve metrics on
    .EXAMPLE
        Get-VSANPerformanceEntityMetric -EntityType "cache-disk"
#>
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("cache-disk","capacity-disk","cluster-domclient","cluster-domcompmgr",
        "disk-group","host-domclient","host-domcompmgr","virtual-disk","virtual-machine",
        "vsan-host-net","vsan-iscsi-host","vsan-iscsi-lun","vsan-iscsi-target","vsan-pnic-net",
        "vsan-vnic-net","vscsi"
        )]
        [String]$EntityType
    )

    $vpm = Get-VSANView -Id "VsanPerformanceManager-vsan-performance-manager"
    $entityTypes = $vpm.VsanPerfGetSupportedEntityTypes()

    $results = @()
    foreach ($et in $entityTypes) {
        if($et.Name -eq $EntityType) {
            $graphs = $et.Graphs
            foreach ($graph in $graphs) {
                foreach ($metric in $graph.Metrics) {
                    $metricObj = [pscustomobject] @{
                        MetricID = $metric.Label
                        Description = $metric.Description
                    }
                    $results+=$metricObj
                }
            }
        }
    }
    $results | Sort-Object -Property MetricID
}

Function Get-VSANPerformanceStat {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retreives a particlular vSAN Performance Metric 
        from a vSAN Cluster using vSAN Management APIs 
    .PARAMETER Cluster
        The name of the vSAN Cluster
    .PARAMETER StartTime
        The start time to scope the query (format: "04/23/2017 4:00")
    .PARAMETER EndTime
        The end time to scope the query (format: "04/23/2017 4:10")
    .PARAMETER EntityId
        The vSAN Management API Entity Reference. Please refer to vSAN Mgmt API docs
    .PARAMETER Metric
        The vSAN performance metric name for the given entity
    .EXAMPLE
        Get-VSANPerformanceStat -Cluster VSAN-Cluster -StartTime "04/23/2017 4:00" -EndTime "04/23/2017 4:05" -EntityId "disk-group:5239bee8-9297-c091-df17-241a4c197f8d" -Metric iopsSched
#>
    param(
        [Parameter(Mandatory=$true)][String]$Cluster,
        [Parameter(Mandatory=$true)][String]$StartTime,
        [Parameter(Mandatory=$true)][String]$EndTime,
        [Parameter(Mandatory=$true)][String]$EntityId,
        [Parameter(Mandatory=$true)][String]$Metric
    )
    function Convert-StringToDateTime {
        # Borrowed from https://blogs.technet.microsoft.com/heyscriptingguy/2014/12/19/powertip-convert-string-into-datetime-object/#comment-209544
        param
        (
        [Parameter(Mandatory = $true)]
        [String] $DateTimeStr
        )
        $DateFormatParts = (Get-Culture).DateTimeFormat.ShortDatePattern -split ‘/|-|\.’

        $Month_Index = ($DateFormatParts | Select-String -Pattern ‘M’).LineNumber – 1
        $Day_Index = ($DateFormatParts | Select-String -Pattern ‘d’).LineNumber – 1
        $Year_Index = ($DateFormatParts | Select-String -Pattern ‘y’).LineNumber – 1

        $DateTimeParts = $DateTimeStr -split ‘/|-|\.| ‘
        $DateTimeParts_LastIndex = $DateTimeParts.Count – 1

        $DateTime = [DateTime] $($DateTimeParts[$Month_Index] + ‘/’ + $DateTimeParts[$Day_Index] + ‘/’ + $DateTimeParts[$Year_Index] + ‘ ‘ + $DateTimeParts[3..$DateTimeParts_LastIndex] -join ‘ ‘)

        return $DateTime
    }

    $cluster_view = (Get-Cluster -Name $cluster).ExtensionData.MoRef

    $vpm = Get-VSANView -Id "VsanPerformanceManager-vsan-performance-manager"

    $start = Convert-StringToDateTime $StartTime
    $end = Convert-StringToDateTime $EndTime

    $spec = New-Object VMware.Vsan.Views.VsanPerfQuerySpec
    $spec.EntityRefId = $EntityId
    $spec.Labels = @($Metric)
    $spec.StartTime = $start
    $spec.EndTime = $end
    $vpm.VsanPerfQueryPerf(@($spec),$cluster_view)
}﻿Function Get-VSANSmartsData {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retreives SMART drive data using new vSAN
        Management 6.6 API. This can also be used outside of vSAN
        to query existing SSD devices not being used for vSAN.
    .PARAMETER Cluster
        The name of a vSAN Cluster
    .EXAMPLE
        Get-VSANSmartsData -Cluster VSAN-Cluster
#>
   param(
        [Parameter(Mandatory=$false)][String]$Cluster
    )

    if($global:DefaultVIServer.ExtensionData.Content.About.ApiType -eq "VirtualCenter") {
        if(!$cluster) {
            Write-Host "Cluster property is required when connecting to vCenter Server"
            break
        }

        $vchs = Get-VSANView -Id "VsanVcClusterHealthSystem-vsan-cluster-health-system"
        $cluster_view = (Get-Cluster -Name $Cluster).ExtensionData.MoRef
        $result = $vchs.VsanQueryVcClusterSmartStatsSummary($cluster_view)
    } else {
        $vhs = Get-VSANView -Id "HostVsanHealthSystem-ha-vsan-health-system"
        $result = $vhs.VsanHostQuerySmartStats($null,$true)
    }

    $vmhost = $result.Hostname
    $smartsData = $result.SmartStats

    Write-Host "`nESXi Host: $vmhost`n"
    foreach ($data in $smartsData) {
        if($data.stats) {
            $stats = $data.stats
            Write-Host $data.disk

            $smartsResults = @()
            foreach ($stat in $stats) {
                $statResult = [pscustomobject] @{
                    Parameter = $stat.Parameter;
                    Value =$stat.Value;
                    Threshold = $stat.Threshold;
                    Worst = $stat.Worst
                }
                $smartsResults+=$statResult
            }
            $smartsResults | Format-Table
        }
    }
}﻿Function Get-VSANVMToUUID {
    <#
        .NOTES
        ===========================================================================
         Created by:    William Lam
         Organization:  VMware
         Blog:          www.virtuallyghetto.com
         Twitter:       @lamw
            ===========================================================================
        .DESCRIPTION
            This function demonstrates the use of the vSAN Management API to retrieve
            the vSAN UUID given a VM Name
        .PARAMETER Cluster
            The name of a vSAN Cluster
        .EXAMPLE
            Get-VSANVMToUUID -Cluster VSAN-Cluster -VMName Embedded-vCenter-Server-Appliance
    #>
        param(
            [Parameter(Mandatory=$true)][String]$Cluster,
            [Parameter(Mandatory=$true)][String]$VMName
        )

        $clusterView = Get-Cluster -Name $Cluster -ErrorAction SilentlyContinue
        if($clusterView) {
            $clusterMoRef = $clusterView.ExtensionData.MoRef
            $vmMoRef = "VirtualMachine-" + (Get-VM -Name $VMName).ExtensionData.MoRef.Value
        } else {
            Write-Host -ForegroundColor Red "Unable to find vSAN Cluster $cluster ..."
            break
        }

        $vsanClusterObjectSys = Get-VsanView -Id VsanObjectSystem-vsan-cluster-object-system
        $results = $vsanClusterObjectSys.VsanQueryObjectIdentities($clusterMoRef,$null,$null,$false,$true,$false)

        $vmObjectInfo = @()
        foreach ($result in $results.Identities) {
            if($result.Vm -eq $vmMoRef) {
                $tmp = [pscustomobject] @{
                    Type=$result.type;
                    UUID=$result.uuid;
                    File=$result.description
                }
                $vmObjectInfo+=$tmp
            }
        }
        $vmObjectInfo | Sort-Object -Property Type,File
    }

Function Get-VSANUUIDToVM {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function demonstrates the use of the vSAN Management API to retrieve
        the VM Name/Object given vSAN UUID
    .PARAMETER VSANObjectID
        List of vSAN Object UUIDs
    .PARAMETER Cluster
        The name of a vSAN Cluster
    .EXAMPLE
        Get-VSANUUIDToVM -VSANObjectID @("6a887f59-6448-08f2-155d-b8aeed7c9e96") -Cluster VSAN-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$Cluster,
        [Parameter(Mandatory=$true)][String[]]$VSANObjectID
    )

    $clusterView = Get-Cluster -Name $Cluster -ErrorAction SilentlyContinue
    if($clusterView) {
        $vmhost = ($clusterView | Get-VMHost) | select -First 1
        $vsanIntSys = Get-View $vmhost.ExtensionData.configManager.vsanInternalSystem
    } else {
        Write-Host -ForegroundColor Red "Unable to find vSAN Cluster $cluster ..."
        break
    }

    $results = @()
    $jsonResult = ($vsanIntSys.GetVsanObjExtAttrs($VSANObjectID)) | ConvertFrom-JSON
    foreach ($object in $jsonResult | Get-Member) {
        # crappy way to iterate through keys ...
        if($($object.Name) -ne "Equals" -and $($object.Name) -ne "GetHashCode" -and $($object.Name) -ne "GetType" -and $($object.Name) -ne "ToString") {
            $objectID = $object.name
            $jsonResult.$($objectID)
        }
    }
}﻿Function Get-VSANVersion {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function retreives the vSAN software version for both VC/ESXi
    .PARAMETER Cluster
        The name of a vSAN Cluster
    .EXAMPLE
        Get-VSANVersion -Cluster VSAN-Cluster
#>
   param(
        [Parameter(Mandatory=$true)][String]$Cluster
    )
    $vchs = Get-VSANView -Id "VsanVcClusterHealthSystem-vsan-cluster-health-system"
    $cluster_view = (Get-Cluster -Name $Cluster).ExtensionData.MoRef
    $results = $vchs.VsanVcClusterQueryVerifyHealthSystemVersions($cluster_view)

    Write-Host "`nVC Version:"$results.VcVersion
    $results.HostResults | Select Hostname, Version
}
﻿Function Get-VSANVMDetailedUsage {
<#
    .NOTES
    ===========================================================================
    Created by:    William Lam
    Organization:  VMware
    Blog:          www.virtuallyghetto.com
    Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function demonstrates the use of vSAN Management API to retrieve
        detailed usage for all or specific VMs running on VSAN
    .PARAMETER Cluster
        The name of a vSAN Cluster
    .PARAMETER VM
        The name of a VM to query specifically
    .EXAMPLE
        Get-VSANVMDetailedUsage -Cluster "VSAN-Cluster"
    .EXAMPLE
        Get-VSANVMDetailedUsage -Cluster "VSAN-Cluster" -VM "Ubuntu-SourceVM"
#>
    param(
        [Parameter(Mandatory=$true)][String]$Cluster,
        [Parameter(Mandatory=$false)][String]$VM
    )

    $ESXiHostUsername = ""
    $ESXiHostPassword = ""

    if($ESXiHostUsername -eq "" -or $ESXiHostPassword -eq "") {
        Write-Host -ForegroundColor Red "You did not configure the ESXi host credentials, please update `$ESXiHostUsername & `$ESXiHostPassword variables and try again"
        return
    }

    # Scope query within vSAN/vSphere Cluster
    $clusterView = Get-View -ViewType ClusterComputeResource -Property Name,Host -Filter @{"name"=$Cluster}
    if(!$clusterView) {
        Write-Host -ForegroundColor Red "Unable to find vSAN Cluster $cluster ..."
        break
    }

    # Retrieve list of ESXi hosts from cluster
    # which we will need to directly connect to use call VsanQueryObjectIdentities()
    $vmhosts = $clusterView.host

    $results = @()
    foreach ($vmhost in $vmhosts) {
        $vmhostView = Get-View $vmhost -Property Name
        $esxiConnection = Connect-VIServer -Server $vmhostView.name -User $ESXiHostUsername -Password $ESXiHostPassword

        $vos = Get-VSANView -Id "VsanObjectSystem-vsan-object-system" -Server $esxiConnection
        $identities = $vos.VsanQueryObjectIdentities($null,$null,$null,$false,$true,$true)

        $json = $identities.RawData|ConvertFrom-Json
        $jsonResults = $json.identities.vmIdentities

        foreach ($vmInstance in $jsonResults) {
            $identities = $vmInstance.objIdentities
            foreach ($identity in $identities | Sort-Object -Property "type") {
                # Retrieve the VM Name
                if($identity.type -eq "namespace") {
                    $vsanIntSys = Get-View (Get-VMHost -Server $esxiConnection).ExtensionData.ConfigManager.vsanInternalSystem
                    $attributes = ($vsanIntSys.GetVsanObjExtAttrs($identity.uuid)) | ConvertFrom-JSON

                    foreach ($attribute in $attributes | Get-Member) {
                        # crappy way to iterate through keys ...
                        if($($attribute.Name) -ne "Equals" -and $($attribute.Name) -ne "GetHashCode" -and $($attribute.Name) -ne "GetType" -and $($attribute.Name) -ne "ToString") {
                            $objectID = $attribute.name
                            $vmName = $attributes.$($objectID).'User friendly name'
                        }
                    }
                }

                # Convert B to GB
                $physicalUsedGB = [math]::round($identity.physicalUsedB/1GB, 2)
                $reservedCapacityGB = [math]::round($identity.reservedCapacityB/1GB, 2)

                # Build our custom object to store only the data we care about
                $tmp = [pscustomobject] @{
                    VM = $vmName
                    File = $identity.description;
                    Type = $identity.type;
                    physicalUsedGB = $physicalUsedGB;
                    reservedCapacityGB = $reservedCapacityGB;
                }

                # Filter out a specific VM if provided
                if($VM) {
                    if($vmName -eq $VM) {
                        $results += $tmp
                    }
                } else {
                    $results += $tmp
                }
            }
        }
        Disconnect-VIServer -Server $esxiConnection -Confirm:$false
    }
    $results | Format-Table
}﻿Function Get-VSANVMThickSwap {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
        ===========================================================================
    .DESCRIPTION
        This function demonstrates the use of vSAN Management API to retrieve
        all VMs that have "thick" provisioned VM Swap
    .PARAMETER Cluster
        The name of a vSAN Cluster
    .EXAMPLE
        Get-VSANVMThickSwap -Cluster VCSA-Cluster
#>
    param(
        [Parameter(Mandatory=$true)][String]$Cluster
    )
    
    # Scope query within vSAN/vSphere Cluster 
    $clusterView = Get-Cluster -Name $Cluster -ErrorAction SilentlyContinue
    if($clusterView) {
        $clusterMoref = $clusterView.ExtensionData.MoRef
    } else {
        Write-Host -ForegroundColor Red "Unable to find vSAN Cluster $cluster ..."
        break
    }

    # Retrieve random ESXi host within vSAN/vSphere Cluster to access vSAN Internal System Object
    $randomVMhost = $clusterView | Get-VMHost | Get-Random
    $vsanIntSys = Get-View ($randomVMhost.ExtensionData.ConfigManager.VsanInternalSystem)
   
    # Create mapping of VMs within vSAN/vSphere Cluster to their associated MoRef
    $vmMoRefIdMapping = @{}
    $vms = Get-Cluster -Name $Cluster | Get-VM
    foreach ($vm in $vms) {
        $vmMoRefIdMapping[$vm.ExtensionData.MoRef] = $vm.name
    }
    
    # Retrieve all vSAN vmswap objects
    $vos = Get-VSANView -Id "VsanObjectSystem-vsan-cluster-object-system" 
    $results = $vos.VsanQueryObjectIdentities($clusterMoref,$null,'vmswap',$false,$true,$false)

    # Process results and look for vmswaps that are "thick" and return array of VM names
    $vmsWithThickSwap = @()
    foreach ($result in $results.Identities) {
        $vsanuuid = $result.uuid
        $vmMoref = $result.vm
        $vmName = $vmMoRefIdMapping[$vmMoref]
        $json = $vsanIntSys.GetVsanObjExtAttrs(@($vsanuuid)) | ConvertFrom-Json
        foreach ($line in $json | Get-Member) {
            $allocationType = $json.$($line.Name).'Allocation type'
            if($allocationType -eq "Zeroed thick") {
                $vmsWithThickSwap +=$vmName
            }
        }
    }
    $vmsWithThickSwap
}﻿Function Get-vSphereAPIUsage {
<#
    .NOTES
    ===========================================================================
     Created by:    William Lam
     Organization:  VMware
     Blog:          www.virtuallyghetto.com
     Twitter:       @lamw
    ===========================================================================
    .DESCRIPTION
        This function returns the list of vSphere APIs used by specific vCenter Server
        session id and path to vpxd.log file
    .PARAMETER VpxdLogFile
        Full path to a vpxd.log file which has been downloaded remotely from a vCenter Server
    .PARAMETER SessionId
        The vCenter Server Session Id you wish to query
    .EXAMPLE
        Get-vSphereAPIUsage -VpxdLogFile "C:\Users\lamw\Dropbox\vpxd.log" -SessionId "52bb9a98-598d-26e9-46d0-ee85d3912646"
#>
    param(
        [Parameter(Mandatory=$true)]$VpxdLogFile,
        [Parameter(Mandatory=$true)]$SessionId
    )
    $vpxdLog = Get-Content -Path $VpxdLogFile

    $apiTally = @{}
    foreach ($line in $vpxdLog) {
        if($line -match $SessionId -and $line -match "[VpxLRO]" -and $line -match "BEGIN") {
            $field = $line -split " "
            if($field[13] -match "vim" -or $field[13] -match "vmodl") {
                $apiTally[$field[13]] += 1
            }
        }
    }
    $commandDuration = Measure-Command {
        $results = $apiTally.GetEnumerator() | Sort-Object -Property Value | FT -AutoSize @{Name=”vSphereAPI”;e={$_.Name}}, @{Name=”Frequency”; e={$_.Value}}
    }

    $duration = $commandDuration.TotalMinutes
    $fileSize = [math]::Round((Get-Item -Path $vpxdLogFile).Length / 1MB,2)
    Write-host "`nFileName: $vpxdLogFile"
    Write-host "FileSize: $fileSize MB"
    Write-Host "Duration: $duration minutes"
    $results
}
Function Get-VSphereCertificateDetails {
<#
    .NOTES
    ===========================================================================
    Created by:    William Lam
    Organization:  VMware
    Blog:          www.virtuallyghetto.com
    Twitter:       @lamw
    ===========================================================================
    .DESCRIPTION
        This function returns the certificate mode of vCenter Server along with
        the certificate details of each ESXi hosts being managed by vCenter Server
    .EXAMPLE
        Get-VSphereCertificateDetails
#>
    if($global:DefaultVIServer.ProductLine -eq "vpx") {
        $vCenterCertMode = (Get-AdvancedSetting -Entity $global:DefaultVIServer -Name vpxd.certmgmt.mode).Value
        Write-Host -ForegroundColor Cyan "`nvCenter $(${global:DefaultVIServer}.Name) Certificate Mode: $vCenterCertMode"
    }

    $results = @()
    $vmhosts = Get-View -ViewType HostSystem -Property Name,ConfigManager.CertificateManager
    foreach ($vmhost in $vmhosts) {
        $certConfig = (Get-View $vmhost.ConfigManager.CertificateManager).CertificateInfo
        if($certConfig.Subject -match "vmca@vmware.com") {
            $certType = "VMCA"
        } else {
            $certType = "Custom"
        }
        $tmp = [PSCustomObject] @{
            VMHost = $vmhost.Name;
            CertType = $certType;
            Status = $certConfig.Status;
            Expiry = $certConfig.NotAfter;
        }
        $results+=$tmp
    }
    $results
}﻿Function Get-vSphereLogins {
    <#
    .SYNOPSIS Retrieve information for all currently logged in vSphere Sessions (excluding current session)
    .NOTES  Author:  William Lam
    .NOTES  Site:    www.virtuallyghetto.com
    .REFERENCE Blog: http://www.virtuallyghetto.com/2016/11/an-update-on-how-to-retrieve-useful-information-from-a-vsphere-login.html 
    .EXAMPLE
      Get-vSphereLogins
    #>
    if($DefaultVIServers -eq $null) {
        Write-Host "Error: Please connect to your vSphere environment"
        exit
    }

    # Using the first connection
    $VCConnection = $DefaultVIServers[0]

    $sessionManager = Get-View ($VCConnection.ExtensionData.Content.SessionManager)

    # Store current session key
    $currentSessionKey = $sessionManager.CurrentSession.Key

    foreach ($session in $sessionManager.SessionList) {
        # Ignore vpxd-extension logins as well as the current session
        if($session.UserName -notmatch "vpxd-extension" -and $session.key -ne $currentSessionKey) {
            $session | Select Username, IpAddress, UserAgent, @{"Name"="APICount";Expression={$Session.CallCount}}, LoginTime
        }
    }
}
﻿<#
.SYNOPSIS
   This script demonstrates an xVC-vMotion where a running Virtual Machine
   is live migrated between two vCenter Servers which are NOT part of the
   same SSO Domain which is only available using the vSphere 6.0 API.

   This script also supports live migrating a running Virtual Machine between
   two vCenter Servers that ARE part of the same SSO Domain (aka Enhanced Linked Mode)

   This script also supports migrating VMs connected to both a VSS/VDS as well as having multiple vNICs

   This script also supports migrating to/from VMware Cloud on AWS (VMC)
.NOTES
   File Name  : xMove-VM.ps1
   Author     : William Lam - @lamw
   Version    : 1.0

   Updated by  : Askar Kopbayev - @akopbayev
   Version     : 1.1
   Description : The script allows to run compute-only xVC-vMotion when the source VM has multiple disks on differnet datastores.

   Updated by  : William Lam - @lamw
   Version     : 1.2
   Description : Added additional parameters to be able to perform cold migration to/from VMware Cloud on AWS (VMC)
                 -ResourcePool
                 -uppercaseuuid

.LINK
    http://www.virtuallyghetto.com/2016/05/automating-cross-vcenter-vmotion-xvc-vmotion-between-the-same-different-sso-domain.html
.LINK
   https://github.com/lamw

.INPUTS
   sourceVCConnection, destVCConnection, vm, switchtype, switch,
   cluster, resourcepool, datastore, vmhost, vmnetworks, $xvctype, $uppercaseuuid
.OUTPUTS
   Console output
#>

Function xMove-VM {
    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [VMware.VimAutomation.ViCore.Util10.VersionedObjectImpl]$sourcevc,
    [VMware.VimAutomation.ViCore.Util10.VersionedObjectImpl]$destvc,
    [String]$vm,
    [String]$switchtype,
    [String]$switch,
    [String]$cluster,
    [String]$resourcepool,
    [String]$datastore,
    [String]$vmhost,
    [String]$vmnetworks,
    [Int]$xvctype,
    [Boolean]$uppercaseuuid
    )

    # Retrieve Source VC SSL Thumbprint
    $vcurl = "https://" + $destVC
add-type @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;

            public class IDontCarePolicy : ICertificatePolicy {
            public IDontCarePolicy() {}
            public bool CheckValidationResult(
                ServicePoint sPoint, X509Certificate cert,
                WebRequest wRequest, int certProb) {
                return true;
            }
        }
"@
    [System.Net.ServicePointManager]::CertificatePolicy = new-object IDontCarePolicy
    # Need to do simple GET connection for this method to work
    Invoke-RestMethod -Uri $VCURL -Method Get | Out-Null

    $endpoint_request = [System.Net.Webrequest]::Create("$vcurl")
    # Get Thumbprint + add colons for a valid Thumbprint
    $destVCThumbprint = ($endpoint_request.ServicePoint.Certificate.GetCertHashString()) -replace '(..(?!$))','$1:'

    # Source VM to migrate
    $vm_view = Get-View (Get-VM -Server $sourcevc -Name $vm) -Property Config.Hardware.Device

    # Dest Datastore to migrate VM to
    $datastore_view = (Get-Datastore -Server $destVCConn -Name $datastore)

    # Dest Cluster/ResourcePool to migrate VM to
    if($cluster) {
        $cluster_view = (Get-Cluster -Server $destVCConn -Name $cluster)
        $resource = $cluster_view.ExtensionData.resourcePool
    } else {
        $rp_view = (Get-ResourcePool -Server $destVCConn -Name $resourcepool)
        $resource = $rp_view.ExtensionData.MoRef
    }

    # Dest ESXi host to migrate VM to
    $vmhost_view = (Get-VMHost -Server $destVCConn -Name $vmhost)

    # Find all Etherenet Devices for given VM which
    # we will need to change its network at the destination
    $vmNetworkAdapters = @()
    $devices = $vm_view.Config.Hardware.Device
    foreach ($device in $devices) {
        if($device -is [VMware.Vim.VirtualEthernetCard]) {
            $vmNetworkAdapters += $device
        }
    }

    # Relocate Spec for Migration
    $spec = New-Object VMware.Vim.VirtualMachineRelocateSpec
    $spec.datastore = $datastore_view.Id
    $spec.host = $vmhost_view.Id
    $spec.pool = $resource

    # Relocate Spec Disk Locator
    if($xvctype -eq 1){
        $HDs = Get-VM -Server $sourcevc -Name $vm | Get-HardDisk
        $HDs | %{
            $disk = New-Object VMware.Vim.VirtualMachineRelocateSpecDiskLocator
            $disk.diskId = $_.Extensiondata.Key
            $SourceDS = $_.FileName.Split("]")[0].TrimStart("[")
            $DestDS = Get-Datastore -Server $destvc -name $sourceDS
            $disk.Datastore = $DestDS.ID
            $spec.disk += $disk
        }
    }

    # Service Locator for the destination vCenter Server
    # regardless if its within same SSO Domain or not
    $service = New-Object VMware.Vim.ServiceLocator
    $credential = New-Object VMware.Vim.ServiceLocatorNamePassword
    $credential.username = $destVCusername
    $credential.password = $destVCpassword
    $service.credential = $credential
    # For some xVC-vMotion, VC's InstanceUUID must be in all caps
    # Haven't figured out why, but this flag would allow user to toggle (default=false)
    if($uppercaseuuid) {
        $service.instanceUuid = $destVCConn.InstanceUuid
    } else {
        $service.instanceUuid = ($destVCConn.InstanceUuid).ToUpper()
    }
    $service.sslThumbprint = $destVCThumbprint
    $service.url = "https://$destVC"
    $spec.service = $service

    # Create VM spec depending if destination networking
    # is using Distributed Virtual Switch (VDS) or
    # is using Virtual Standard Switch (VSS)
    $count = 0
    if($switchtype -eq "vds") {
        foreach ($vmNetworkAdapter in $vmNetworkAdapters) {
            # New VM Network to assign vNIC
            $vmnetworkname = ($vmnetworks -split ",")[$count]

            # Extract Distributed Portgroup required info
            $dvpg = Get-VDPortgroup -Server $destvc -Name $vmnetworkname
            $vds_uuid = (Get-View $dvpg.ExtensionData.Config.DistributedVirtualSwitch).Uuid
            $dvpg_key = $dvpg.ExtensionData.Config.key

            # Device Change spec for VSS portgroup
            $dev = New-Object VMware.Vim.VirtualDeviceConfigSpec
            $dev.Operation = "edit"
            $dev.Device = $vmNetworkAdapter
            $dev.device.Backing = New-Object VMware.Vim.VirtualEthernetCardDistributedVirtualPortBackingInfo
            $dev.device.backing.port = New-Object VMware.Vim.DistributedVirtualSwitchPortConnection
            $dev.device.backing.port.switchUuid = $vds_uuid
            $dev.device.backing.port.portgroupKey = $dvpg_key
            $spec.DeviceChange += $dev
            $count++
        }
    } else {
        foreach ($vmNetworkAdapter in $vmNetworkAdapters) {
            # New VM Network to assign vNIC
            $vmnetworkname = ($vmnetworks -split ",")[$count]

            # Device Change spec for VSS portgroup
            $dev = New-Object VMware.Vim.VirtualDeviceConfigSpec
            $dev.Operation = "edit"
            $dev.Device = $vmNetworkAdapter
            $dev.device.backing = New-Object VMware.Vim.VirtualEthernetCardNetworkBackingInfo
            $dev.device.backing.deviceName = $vmnetworkname
            $spec.DeviceChange += $dev
            $count++
        }
    }

    Write-Host "`nMigrating $vmname from $sourceVC to $destVC ...`n"

    # Issue Cross VC-vMotion
    $task = $vm_view.RelocateVM_Task($spec,"defaultPriority")
    $task1 = Get-Task -Id ("Task-$($task.value)")
    $task1 | Wait-Task
}

# Variables that must be defined

$vmname = "TinyVM-2"
$sourceVC = "vcenter60-1.primp-industries.com"
$sourceVCUsername = "administrator@vghetto.local"
$sourceVCPassword = "VMware1!"
$destVC = "vcenter60-3.primp-industries.com"
$destVCUsername = "administrator@vghetto.local"
$destVCpassword = "VMware1!"
$datastorename = "la-datastore1"
$resourcepool = "WorkloadRP"
$vmhostname = "vesxi60-5.primp-industries.com"
$vmnetworkname = "LA-VM-Network1,LA-VM-Network2"
$switchname = "LA-VDS"
$switchtype = "vds"
$ComputeXVC = 1
$UppercaseUUID = $false

# Connect to Source/Destination vCenter Server
$sourceVCConn = Connect-VIServer -Server $sourceVC -user $sourceVCUsername -password $sourceVCPassword
$destVCConn = Connect-VIServer -Server $destVC -user $destVCUsername -password $destVCpassword

xMove-VM -sourcevc $sourceVCConn -destvc $destVCConn -VM $vmname -switchtype $switchtype -switch $switchname -resourcepool $resourcepool -vmhost $vmhostname -datastore $datastorename -vmnetwork  $vmnetworkname -xvcType $computeXVC -uppercaseuuid $UppercaseUUID

# Disconnect from Source/Destination VC
Disconnect-VIServer -Server $sourceVCConn -Confirm:$false
Disconnect-VIServer -Server $destVCConn -Confirm:$false
﻿Function xNew-VM {
    <#
        .SYNOPSIS This script demonstrates a Cross vCenter Clone Operation across two different vCenter Servers which can either be part of the same or different SSO Domain
        .NOTES  Author:  William Lam
        .NOTES  Site:    www.virtuallyghetto.com
        .REFERENCE Blog: http://www.virtuallyghetto.com/2018/01/cross-vcenter-clone-with-vsphere-6-0.html
    #>

    param(
    [Parameter(
        Position=0,
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)
    ]
    [VMware.VimAutomation.ViCore.Util10.VersionedObjectImpl]$sourcevc,
    [VMware.VimAutomation.ViCore.Util10.VersionedObjectImpl]$destvc,
    [String]$sourcevmname,
    [String]$destvmname,
    [String]$switchtype,
    [String]$datacenter,
    [String]$cluster,
    [String]$resourcepool,
    [String]$datastore,
    [String]$vmhost,
    [String]$vmnetworks,
    [String]$foldername,
    [String]$snapshotname,
    [Boolean]$poweron,
    [Boolean]$uppercaseuuid
    )

    # Retrieve Source VC SSL Thumbprint
    $vcurl = "https://" + $destVC
add-type @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;

            public class IDontCarePolicy : ICertificatePolicy {
            public IDontCarePolicy() {}
            public bool CheckValidationResult(
                ServicePoint sPoint, X509Certificate cert,
                WebRequest wRequest, int certProb) {
                return true;
            }
        }
"@
    [System.Net.ServicePointManager]::CertificatePolicy = new-object IDontCarePolicy
    # Need to do simple GET connection for this method to work
    Invoke-RestMethod -Uri $VCURL -Method Get | Out-Null

    $endpoint_request = [System.Net.Webrequest]::Create("$vcurl")
    # Get Thumbprint + add colons for a valid Thumbprint
    $destVCThumbprint = ($endpoint_request.ServicePoint.Certificate.GetCertHashString()) -replace '(..(?!$))','$1:'

    # Source VM to clone from
    $vm_view = Get-View (Get-VM -Server $sourcevc -Name $sourcevmname) -Property Config.Hardware.Device

    # Dest Datastore to clone VM to
    $datastore_view = (Get-Datacenter -Server $destVCConn -Name $datacenter | Get-Datastore -Server $destVCConn -Name $datastore)

    # Dest VM Folder to clone VM to
    $folder_view = (Get-Datacenter -Server $destVCConn -Name $datacenter | Get-Folder -Server $destVCConn -Name $foldername)

    # Dest Cluster/ResourcePool to clone VM to
    if($cluster) {
        $cluster_view = (Get-Datacenter -Server $destVCConn -Name $datacenter | Get-Cluster -Server $destVCConn -Name $cluster)
        $resource = $cluster_view.ExtensionData.resourcePool
    } else {
        $rp_view = (Get-Datacenter -Server $destVCConn -Name $datacenter | Get-ResourcePool -Server $destVCConn -Name $resourcepool)
        $resource = $rp_view.ExtensionData.MoRef
    }

    # Dest ESXi host to clone VM to
    $vmhost_view = (Get-VMHost -Server $destVCConn -Name $vmhost)

    # Find all Etherenet Devices for given VM which
    # we will need to change its network at the destination
    $vmNetworkAdapters = @()
    $devices = $vm_view.Config.Hardware.Device
    foreach ($device in $devices) {
        if($device -is [VMware.Vim.VirtualEthernetCard]) {
            $vmNetworkAdapters += $device
        }
    }

    # Snapshot to clone from
    if($snapshotname) {
        $snapshot = Get-Snapshot -Server $sourcevc -VM $sourcevmname -Name $snapshotname
    }

    # Clone Spec
    $spec = New-Object VMware.Vim.VirtualMachineCloneSpec
    $spec.PowerOn = $poweron
    $spec.Template = $false
    $locationSpec = New-Object VMware.Vim.VirtualMachineRelocateSpec

    $locationSpec.datastore = $datastore_view.Id
    $locationSpec.host = $vmhost_view.Id
    $locationSpec.pool = $resource
    $locationSpec.Folder = $folder_view.Id

    # Service Locator for the destination vCenter Server
    # regardless if its within same SSO Domain or not
    $service = New-Object VMware.Vim.ServiceLocator
    $credential = New-Object VMware.Vim.ServiceLocatorNamePassword
    $credential.username = $destVCusername
    $credential.password = $destVCpassword
    $service.credential = $credential
    # For some xVC-vMotion, VC's InstanceUUID must be in all caps
    # Haven't figured out why, but this flag would allow user to toggle (default=false)
    if($uppercaseuuid) {
        $service.instanceUuid = $destVCConn.InstanceUuid
    } else {
        $service.instanceUuid = ($destVCConn.InstanceUuid).ToUpper()
    }
    $service.sslThumbprint = $destVCThumbprint
    $service.url = "https://$destVC"
    $locationSpec.service = $service

    # Create VM spec depending if destination networking
    # is using Distributed Virtual Switch (VDS) or
    # is using Virtual Standard Switch (VSS)
    $count = 0
    if($switchtype -eq "vds") {
        foreach ($vmNetworkAdapter in $vmNetworkAdapters) {
            # New VM Network to assign vNIC
            $vmnetworkname = ($vmnetworks -split ",")[$count]

            # Extract Distributed Portgroup required info
            $dvpg = Get-VDPortgroup -Server $destvc -Name $vmnetworkname
            $vds_uuid = (Get-View $dvpg.ExtensionData.Config.DistributedVirtualSwitch).Uuid
            $dvpg_key = $dvpg.ExtensionData.Config.key

            # Device Change spec for VSS portgroup
            $dev = New-Object VMware.Vim.VirtualDeviceConfigSpec
            $dev.Operation = "edit"
            $dev.Device = $vmNetworkAdapter
            $dev.device.Backing = New-Object VMware.Vim.VirtualEthernetCardDistributedVirtualPortBackingInfo
            $dev.device.backing.port = New-Object VMware.Vim.DistributedVirtualSwitchPortConnection
            $dev.device.backing.port.switchUuid = $vds_uuid
            $dev.device.backing.port.portgroupKey = $dvpg_key
            $locationSpec.DeviceChange += $dev
            $count++
        }
    } else {
        foreach ($vmNetworkAdapter in $vmNetworkAdapters) {
            # New VM Network to assign vNIC
            $vmnetworkname = ($vmnetworks -split ",")[$count]

            # Device Change spec for VSS portgroup
            $dev = New-Object VMware.Vim.VirtualDeviceConfigSpec
            $dev.Operation = "edit"
            $dev.Device = $vmNetworkAdapter
            $dev.device.backing = New-Object VMware.Vim.VirtualEthernetCardNetworkBackingInfo
            $dev.device.backing.deviceName = $vmnetworkname
            $locationSpec.DeviceChange += $dev
            $count++
        }
    }

    $spec.Location = $locationSpec

    if($snapshot) {
        $spec.Snapshot = $snapshot.Id
    }

    Write-Host "`nCloning $sourcevmname from $sourceVC to $destVC ...`n"

    # Issue Cross VC-vMotion
    $task = $vm_view.CloneVM_Task($folder_view.Id,$destvmname,$spec)
    $task1 = Get-Task -Server $sourceVCConn -Id ("Task-$($task.value)")
}

# Variables that must be defined

$sourcevmname = "PhotonOS-02"
$destvmname= "PhotonOS-02-Clone"
$sourceVC = "vcenter65-1.primp-industries.com"
$sourceVCUsername = "administrator@vsphere.local"
$sourceVCPassword = "VMware1!"
$destVC = "vcenter65-3.primp-industries.com"
$destVCUsername = "administrator@vsphere.local"
$destVCpassword = "VMware1!"
$datastorename = "vsanDatastore"
$datacenter = "Datacenter-SiteB"
$cluster = "Santa-Barbara"
$resourcepool = "MyRP" # cluster property not needed if rp is used
$vmhostname = "vesxi65-4.primp-industries.com"
$vmnetworkname = "VM Network"
$foldername = "Discovered virtual machine"
$switchtype = "vss"
$poweron = $false
$snapshotname = "pristine"
$UppercaseUUID = $true

# Connect to Source/Destination vCenter Server
$sourceVCConn = Connect-VIServer -Server $sourceVC -user $sourceVCUsername -password $sourceVCPassword
$destVCConn = Connect-VIServer -Server $destVC -user $destVCUsername -password $destVCpassword

xNew-VM -sourcevc $sourceVCConn -destvc $destVCConn -sourcevmname $sourcevmname -destvmname `
    $destvmname -switchtype $switchtype -datacenter $datacenter -cluster $cluster -vmhost `
    $vmhostname -datastore $datastorename -vmnetwork  $vmnetworkname -foldername `
    $foldername -poweron $poweron -uppercaseuuid $UppercaseUUID -$snapshotname $snapshotname

# Disconnect from Source/Destination VC
Disconnect-VIServer -Server $sourceVCConn -Confirm:$false
Disconnect-VIServer -Server $destVCConn -Confirm:$false
