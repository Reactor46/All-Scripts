# Import Modules
if ((Get-Module |where {$_.Name -ilike "CiscoUcsPS"}).Name -ine "CiscoUcsPS")
	{
	Write-Host "Loading Module: Cisco UCS PowerTool Module"
	Import-Module CiscoUcsPs
	}
if ((Get-PSSnapin | where {$_.Name -ilike "Vmware*Core"}).Name -ine "VMware.VimAutomation.Core")
	{
	Write-Host "Loading PS Snap-in: VMware VimAutomation Core"
	Add-PSSnapin VMware.VimAutomation.Core -ErrorAction SilentlyContinue
	}
if ((Get-PSSnapin | where {$_.Name -ilike "Vmware*Core"}).Name -ine "VMware.VimAutomation.Core")
	{
	Write-Host "Loading PS Snap-in: VMware VimAutomation Core"
	Add-PSSnapin VMware.VimAutomation.Core -ErrorAction SilentlyContinue
	}
if ((Get-PSSnapin | where {$_.Name -ilike " VMware.DeployAutomation"}).Name -ine "VMware.DeployAutomation")
	{
	Write-Host "Loading PS Snap-in: VMware VMware.DeployAutomation"
	Add-PSSnapin VMware.DeployAutomation -ErrorAction SilentlyContinue
	}
if ((Get-PSSnapin | where {$_.Name -ilike "VMware.ImageBuilder"}).Name -ine "VMware.ImageBuilder")
	{
	Write-Host "Loading PS Snap-in: VMware VMware.ImageBuilder"
	Add-PSSnapin VMware.ImageBuilder -ErrorAction SilentlyContinue
	}
	
set-ucspowertoolconfiguration -supportmultipledefaultucs 0

# Global Variables
$ucs = "172.25.206.5"
$ucsuser = "ucs-ericwill\admin"
$ucspass = "Nbv12345!"
$ucsorg = "org-root"
$tenantname = "CL2012"
$vCenter = "172.25.206.186"
$vcuser = "Administrator"
$vcpass = "Nbv12345"
$NewImageProfile = "C:\VMware\ESX5.NextImageProfile.zip"
$ucshfpname = "2.1-0.323"
$WarningPreference = "SilentlyContinue"

try {
	# Login to UCS
	Write-Host "UCS: Logging into UCS Domain: $ucs"
	$ucspasswd = ConvertTo-SecureString $ucspass -AsPlainText -Force
	$ucscreds = New-Object System.Management.Automation.PSCredential ($ucsuser, $ucspasswd)
	$ucslogin = Connect-Ucs -Credential $ucscreds $ucs
	
	Write-Host "vC: Logging into vCenter: $vCenter"
	$vcenterlogin = Connect-VIServer $vCenter -User $vcuser -Password $vcpass | Out-Null

	Write-Host "vC: Disable the current ESXi Image DeployRule" 
	$RemDepRule = Get-DeployRule -Name "DeployESXiImage" | Remove-DeployRule
	$ESXDepot = Add-EsxSoftwareDepot "https://hostupdate.vmware.com/software/VUM/PRODUCTION/main/vmw-depot-index.xml"
	#$ESXDeppt = Add-EsxSoftwareDepot $NewImageProfile
	$LatestImageProfile = Get-EsxImageProfile | sort ModifiedTime -Descending | Select -First 1
	$pattern = "oemstring=`$SPT:$($tenantname)"
	Write-Host "vC: Creating ESXi deploy rule for '$($pattern)'"
	$NewRule = New-DeployRule -Name "AddHostsTo$($tenantname)ClusterUpdatedImage" -Item $LatestImageProfile -Pattern $pattern
	$SetActive = $NewRule | Add-DeployRule

	Write-Host "vC: Repairing active ruleset"
	$RepairRules = Get-VMHost | Test-DeployRuleSetCompliance | Repair-DeployRuleSetCompliance -ErrorAction SilentlyContinue

	Foreach ($VMHost in (Get-Cluster $tenantname | Get-VMHost )) {
		Write-Host "vC: Adding VM Hypervisor Host: $($VMHost.Name) into maintenance mode"
		$Maint = $VMHost | Set-VMHost -State Maintenance
		
		Write-Host "vC: Waiting for VM Hypervisor Host: $($VMHost.Name) to enter Maintenance Mode"
		do {
			Sleep 10
		} until ((Get-VMHost $VMHost).State -eq "Maintenance")
		
		Write-Host "vC: VM Hypervisor Host: $($VMHost.Name) now in Maintenance Mode, shutting down Host"
		$Shutdown = $VMHost.ExtensionData.ShutdownHost($true)
		
		Write-Host "UCS: Correlating VM Hypervisor Host: $($VMHost.Name) to running UCS Service Profile (SP)"

		$vmMacAddr = $vmhost.NetworkInfo.PhysicalNic | where { $_.name -ieq "vmnic0" }
		
		$sp2upgrade =  Get-UcsServiceProfile | Get-UcsVnic -Name eth0 |  where { $_.addr -ieq  $vmMacAddr.Mac } | Get-UcsParent 
		
		Write-Host "UCS: VM Hypervisor Host: $($VMhost.Name) is running on UCS SP: $($sp2upgrade.name)"
		Write-Host "UCS: Waiting to for UCS SP: $($sp2upgrade.name) to gracefully power down"
	 	do {
			if ( (get-ucsmanagedobject -dn $sp2upgrade.PnDn).OperPower -eq "off")
			{
				break
			}
			Sleep 40
		} until ((get-ucsmanagedobject -dn $sp2upgrade.PnDn).OperPower -eq "off" )
		Write-Host "UCS: UCS SP: $($sp2upgrade.name) powered down"
		
		Write-Host "UCS: Setting desired power state for UCS SP: $($sp2upgrade.name) to down"
		$poweron = $sp2upgrade | Set-UcsServerPower -State "down" -Force

		Write-Host "UCS: Changing Host Firmware pack policy for UCS SP: $($sp2upgrade.name) to '$($ucshfpname)'"
		$updatehfp = $sp2upgrade | Set-UcsServiceProfile -HostFwPolicyName (Get-UcsFirmwareComputeHostPack -Name $ucshfpname).Name -Force
		
		Write-Host "UCS: Acknowlodging any User Maintenance Actions for UCS SP: $($sp2upgrade.name)"
		if (($sp2upgrade | Get-UcsLsmaintAck| measure).Count -ge 1)
			{
				$ackuserack = $sp2upgrade | get-ucslsmaintack | Set-UcsLsmaintAck -AdminState "trigger-immediate" -Force
			}

		Write-Host "UCS: Waiting for UCS SP: $($sp2upgrade.name) to complete firmware update process for Host Firmware pack '$($ucshfpname)'"
		do {
			Sleep 40
		} until ((Get-UcsManagedObject -Dn $sp2upgrade.Dn).AssocState -ieq "associated")
		
		Write-Host "UCS: Host Firmware Pack update process comlete.  Setting desired power state for UCS SP: $($sp2upgrade.name) to 'up'"
		$poweron = $sp2upgrade | Set-UcsServerPower -State "up" -Force
		
		Write "vC: Waiting for VM Hypervisor Host: $($VMHost.Name) to connect to vCenter"
		do {
			Sleep 40
		} until (($VMHost = Get-VMHost $VMHost).ConnectionState -eq "Connected" )
	}

	# Logout of UCS
	Write-Host "UCS: Logging out of UCS Domain: $ucs"
	$ucslogout = Disconnect-Ucs 

	# Logout of vCenter
	Write-Host "vC: Logging out of vCenter: $vCenter"
	$vcenterlogout = Disconnect-VIServer $vCenter -Confirm:$false
}
Catch 
{
	 Write-Host "Error occurred in script:"
	 Write-Host ${Error}
     exit
}