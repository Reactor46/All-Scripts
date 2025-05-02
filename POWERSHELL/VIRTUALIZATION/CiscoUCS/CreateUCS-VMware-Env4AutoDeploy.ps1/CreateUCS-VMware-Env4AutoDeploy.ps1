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
	
set-ucspowertoolconfiguration -supportmultipledefaultucs $false

# Global Variables
$ucs = "172.25.206.5"
$ucsuser = "ucs-ericwill\admin"
$ucspass = "Nbv12345!"
$ucsorg = "org-root"
$tenantname = "CL2012"
$macpoolblockfrom = "00:25:B5:ED:01:01"
$macpoolblockto = "00:25:B5:ED:01:09"
$wwpnpoolblockfrom = "20:00:00:25:B5:ED:02:01"
$wwpnpoolblockto = "20:00:00:25:B5:ED:02:12"
$private = "10"
$public = "206"
$vCenter = "172.25.206.186"
$vcuser = "Administrator"
$vcpass = "Nbv12345"
$WarningPreference = "SilentlyContinue"

Try {

	# Login to UCS
	Write-Host "UCS: Logging into UCS Domain: $ucs"
	$ucspasswd = ConvertTo-SecureString $ucspass -AsPlainText -Force
	$ucscreds = New-Object System.Management.Automation.PSCredential ($ucsuser, $ucspasswd)
	$ucslogin = Connect-Ucs -Credential $ucscreds $ucs

	# Create Deploy Rule for Hypervisor to be deploy based on SP name and another rule for cluster name
	Write-Host "vC: Logging into vCenter: $vCenter"
	$vcenterlogin = Connect-VIServer $vCenter -User $vcuser -Password $vcpass | Out-Null

	# Create Mac Pool
	Write-Host "UCS: Creating MAC Pool: $tenantname"
	$startucstransaction = Start-UcsTransaction
	$macpool =  Get-UcsManagedObject -Dn $ucsorg | Add-UcsMacPool -name $tenantname -ModifyPresent
	$macpoolblock = $macpool | Add-UcsMacMemberBlock -From $macpoolblockfrom -To $macpoolblockto -ModifyPresent
	$completeucstransaction = Complete-UcsTransaction

	# Create WWPN Pool
	Write-Host "UCS: Creating WWPN Pool: $tenantname"
	$startucstransaction = Start-UcsTransaction
	$wwpnpool = Get-UcsManagedObject -Dn $ucsorg  | Add-UcsWwnPool -Name $tenantname -Purpose "port-wwn-assignment" -ModifyPresent
	$wwpnpoolblock = $wwpnpool | Add-UcsWwnMemberBlock -From $wwpnpoolblockfrom -To $wwpnpoolblockto -ModifyPresent
	$completeucstransaction = Complete-UcsTransaction

	# Create Server Pool
	Write-Host "UCS: Creating Server Pool: $tenantname"
	$serverpool =  Get-UcsManagedObject -Dn $ucsorg | Add-UcsServerPool -Name $tenantname -ModifyPresent

	# Create Server Qualification Policy
	Write-Host "UCS: Creating Server Qualification Policy: $tenantname"
	$startucstransaction = Start-UcsTransaction
	$serverqualpol = Get-UcsManagedObject -Dn $ucsorg | Add-UcsServerPoolQualification -Name $tenantname -ModifyPresent
	$serveradaptorqual = $serverqualpol | Add-UcsAdaptorQualification -ModifyPresent
	$serveradaptorcapqual = $serveradaptorqual | Add-UcsAdaptorCapQualification -Maximum "unspecified" -Model "N20-AC0002" -Type "virtualized-eth-if" -ModifyPresent
	$servercpuqual = $serverqualpol | Add-UcsCpuQualification -Model "N20-X00002" 
	$completeucstransaction = Complete-UcsTransaction

	# Create Server Pool Policy (for dynamic server pools based on qualification policy)
	Write-Host "UCS: Creating Server Pool Policy: $tenantname"
	$startucstransaction = Start-UcsTransaction
	$serverpoolpol = Get-UcsManagedObject -Dn $ucsorg | Add-UcsServerPoolPolicy -Name $tenantname -PoolDn $serverpool.dn -Qualifier $serverqualpol.Name -ModifyPresent
	$completeucstransaction = Complete-UcsTransaction

	# Create Boot Policy
	Write-Host "UCS: Creating Boot Policy: $tenantname"
	$startucstransaction = Start-UcsTransaction
	$bootpol = Get-UcsManagedObject -Dn $ucsorg | Add-UcsBootPolicy -EnforceVnicName "yes" -Name $tenantname -RebootOnUpdate "no" -ModifyPresent
	$pxe = $bootpol | Add-UcsLsbootLan -ModifyPresent -Order "1" -Prot "pxe"  
	$eth0 = $pxe | Add-UcsLsbootLanImagePath -BootIpPolicyName "" -ISCSIVnicName "" -ImgPolicyName "" -ImgSecPolicyName "" -ProvSrvPolicyName "" -Type "primary" -VnicName "eth0"  -ModifyPresent
	$completeucstransaction = Complete-UcsTransaction

	# Create any needed VLANs
	Write-Host "UCS: Creating VLAN vlan$private"
	$vlanprivate = get-ucslancloud | Add-UcsVlan -Name vlan$private -Id $private -ModifyPresent
	Write-Host "UCS: Creating VLAN vlan$public"
	$vlanpublic = get-ucslancloud | Add-UcsVlan -Name vlan$public -Id $public -ModifyPresent

	# Create any needed VSANs
	Write-Host "UCS: Creating VSAN default"
	$vsan1 = Get-UcsSanCloud | Add-UcsVsan -Name default -Id 1 -ModifyPresent

	# Create Service Profile Template (using MAC, WWPN, Server Pools, VLANs, VSans, Boot Policy, etc. previously created steps) with desired power state to down
	Write-Host "UCS: Creating SP Template: $tenantname in UCS org: $ucsorg"
	$startucstransaction = Start-UcsTransaction
	$sptemplate = Get-UcsManagedObject -Dn $ucsorg |  Add-UcsServiceProfile -BootPolicyName $tenantname -HostFwPolicyName "default" -IdentPoolName "default" -LocalDiskPolicyName "default" -MaintPolicyName "default" -Name $tenantname -PowerPolicyName "default" -StatsPolicyName "default" -Type "initial-template" 
	$sptvnic1 = $sptemplate | Add-UcsVnic -AdaptorProfileName "VMWare" -AdminVcon "any" -IdentPoolName $tenantname -Mtu 1500 -Name "eth0" -Order "1" -StatsPolicyName "default" -SwitchId "A-B"
	$vnic1private = $sptvnic1 | Add-UcsVnicInterface -DefaultNet "yes" -Name $vlanprivate.Name
	$sptvnic2 = $sptemplate | Add-UcsVnic -AdaptorProfileName "VMWare" -AdminVcon "any" -IdentPoolName $tenantname -Mtu 1500 -Name "eth1" -Order "2" -StatsPolicyName "default" -SwitchId "B-A"
	$vnic2public = $sptvnic2 | Add-UcsVnicInterface -DefaultNet "yes" -Name $vlanpublic.Name
	$sptfc0 = $sptemplate | Add-UcsVhba -AdaptorProfileName "VMWare" -AdminVcon "any" -IdentPoolName $tenantname -MaxDataFieldSize 2048 -Name "fc0" -Order "3" -PersBind "disabled" -PersBindClear "no" -StatsPolicyName "default" -SwitchId "A"
	$fc0vsan = $sptfc0 | Add-UcsVhbaInterface -Name $vsan1.Name
	$sptfc1 = $sptemplate | Add-UcsVhba -AdaptorProfileName "VMWare" -AdminVcon "any" -IdentPoolName $tenantname -MaxDataFieldSize 2048 -Name "fc1" -Order "4" -PersBind "disabled" -PersBindClear "no" -StatsPolicyName "default" -SwitchId "B"
	$fc1vsan = $sptfc1 | Add-UcsVhbaInterface -Name $vsan1.Name
	$sptwwnn = $sptemplate | Add-UcsVnicFcNode -IdentPoolName "node-default"
	$sptserverpool = $sptemplate | Add-UcsServerPoolAssignment -Name $tenantname -RestrictMigration "no"
	$sptpowerstate = $sptemplate | Set-UcsServerPower -State "down" -Force
	$completeucstransaction = Complete-UcsTransaction

	# Add ESXi Image Profile from VMware.com
	Write-Host "vC: Adding VMware ESXi Library from online depot"
	$SoftwareDepot = Add-EsxSoftwareDepot "https://hostupdate.vmware.com/software/VUM/PRODUCTION/main/vmw-depot-index.xml"
	$LatestImageProfile = Get-EsxImageProfile | Where { $_.Name -match "no-tools" } | Sort ModifiedTime -desc | select -first 1

	# Create Auto-deploy rule using the Service Profile Template OEM string from UCS
	$pattern = "oemstring=`$SPT:$($SPTemplate.name)"
	Write-Host "vC: Creating ESXi deploy rule for '$($pattern)'"
	$DeployRule = New-DeployRule -Name "DeployESXiImage" -Item $LatestImageProfile -Pattern $pattern
	$AddRule = $DeployRule | Add-DeployRule

	# Create vCenter Cluster
	Write-Host "vC: Creating vCenter Cluster: $tenantname if it doesnt already exist"
	If (-Not (Get-Cluster $tenantname -ErrorAction SilentlyContinue)) {
		$Cluster = Get-Datacenter | Select -First 1 | New-Cluster $tenantname
	} Else {
		$Cluster = Get-Cluster $tenantname
	}

	Write-Host "vC: Enabling DRS and setting mode to 'FullyAutomated' for Cluster: $($Cluster)"
	$DRSenable = $Cluster | Set-Cluster -DrsEnabled:$true -DrsMode FullyAutomated -drsautomationlevel FullyAutomated -Confirm:$false
	
	Write-Host "vC: Creating vCenter Cluster rule for '$($pattern)'"
	$DeployRule = New-DeployRule -Name "AddHostsTo$($tenantname)Cluster" -Item $Cluster -Pattern $pattern
	$AddRule = $DeployRule | Add-DeployRule

	Write-Host "vC: Repairing active ruleset"
	$RepairRules = Get-VMHost | Test-DeployRuleSetCompliance | Repair-DeployRuleSetCompliance -ErrorAction SilentlyContinue

	# Create Initial ESXi Host from SP Template in UCS
	Write-Host "UCS: Creating new Service Profile from SP Template: $($sptemplate.name)"
	$createnewsp = $sptemplate | Add-UcsServiceProfileFromTemplate -Count 1 -DestinationOrg (Get-UcsManagedObject -Dn $ucsorg) -Prefix "esxi-host"
	$spmacaddr = $createnewsp | Get-UcsVnic -Name eth0

	# Monitor UCS SP Associate for completion
	Write-Host "UCS: Waiting for UCS SP: $($createnewsp.name) to complete SP association process"
		do {
			if ( (Get-UcsManagedObject -Dn $createnewsp.Dn).AssocState -ieq "associated")
			{
				break
			}
			Sleep 40
		} until ((Get-UcsManagedObject -Dn $createnewsp.Dn).AssocState -ieq "associated")
			
	# Set SP Desired Power State to UP in newly created SP
	Write-Host "UCS: Setting Desired Power State to 'up' of Service Profile: $createnewsp"
	$powerspon = $createnewsp | Set-UcsServerPower -State "up" -Force

	# Wait for Hypervisor to boot from network w/ Auto-Deploy and connect to vCenter
	Write "vC: Waiting for Host with MAC address of $($spmacaddr.Addr) to connect to vCenter"
	do {
		Sleep 40
	} until ($VMHost = (get-vmhost -ErrorAction SilentlyContinue | foreach { $_.NetworkInfo.PhysicalNic | where { $_.Mac -ieq $spmacaddr.Addr } } | select -expandproperty vmhost ))

	If ($VMHost.State -ne "Maintenance") {
		$Maint = $VMHost | Set-VMHost -State Maintenance
	}

	# Config changes to host before host profile
	Write "vC: COnfiguring VM Hypervisor Host: $($VMhost.name) before creating a Host Profile"
	Write "vC: Removing defualt VM Network"
	$VMNetwork = Get-VirtualPortGroup -Name "VM Network" -VMHost $VMHost | Remove-VirtualPortGroup -Confirm:$false
	Write "vC: Creating VMotion Network"
	$Vmotion = New-VMHostNetworkAdapter -VMHost $vmhost -PortGroup "VMotion" -VirtualSwitch (Get-VirtualSwitch -VMHost $vmhost -Name vSwitch0) -VMotionEnabled:$true
	Write "vC: Creating New Virtual Switch"
	$vSwitch1 = New-VirtualSwitch -VMHost $vmhost -Name "vSwitch1" -Nic (Get-VMHostNetworkAdapter -VMHost $vmhost -Name "vmnic1")
	Write "vC: Creating New VM Network"
	$VMNetwork = $vSwitch1 | New-VirtualPortGroup "VM Network"
	Write "vC: Setting Syslog Server"
	$SyslogServer = Set-VMHostSysLogServer -SysLogServer $vCenter -SysLogServerPort 514 -VMHost $VMhost

	Write-Host "vC: Creating Host Profile: $tenantname from VM Hypervisor Host: $($VMhost.name)"
	If (-Not (Get-VMHostProfile $tenantname -ErrorAction SilentlyContinue) ) {
		$HostProfile = New-VMHostProfile -Name $tenantname -Description "Automatically generated host profile for $tenantname" -ReferenceHost $VMHost
		# Edit the host profile to add the static password entry

		function Copy-Property ($From, $To, $PropertyName ="*")
		{
		   foreach ($p in Get-Member -In $From -MemberType Property -Name $propertyName)
		   {  trap {
		         Add-Member -In $To -MemberType NoteProperty -Name $p.Name -Value $From.$($p.Name) -Force
		         continue
		      }
		      $To.$($P.Name) = $From.$($P.Name)
		   }
		}

		$hostProfileName = $tenantname
		$newAdminPswd = "VMw@re123"

		$spec = New-Object VMware.Vim.HostProfileCompleteConfigSpec
		Copy-Property -From $HostProfile.ExtensionData.Config -To $spec

		$secpol = New-Object VMware.Vim.ProfilePolicy
		$secpol.Id = "AdminPasswordPolicy"
		$secpol.PolicyOption = New-Object VMware.Vim.PolicyOption
		$secpol.PolicyOption.Id = "FixedAdminPasswordOption"
		$secpol.PolicyOption.Parameter += New-Object VMware.Vim.KeyAnyValue
		$secpol.PolicyOption.Parameter[0].Key = "password"
		$secpol.PolicyOption.Parameter[0].Value = New-Object VMware.Vim.PasswordField
		$secpol.PolicyOption.Parameter[0].Value.Value = $newAdminPswd
		$spec.ApplyProfile.Security.Policy = @($secpol)

		$ChangeHostProfile = $HostProfile.ExtensionData.UpdateHostProfile($spec)
	} else {
		$Hostprofile = Get-VMHostProfile $tenantname
	}


	# Add a new Deployrule to associate the host profile to the Hosts
	Write-Host "vC: Adding rule to use host profile"
	$DeployRule = New-DeployRule -Name "$($tenantname)HostProfile" -Item $HostProfile -Pattern $pattern
	$AddRule = $DeployRule | Add-DeployRule

	Write-Host vC: "Repairing active ruleset"
	$RepairRules = $VMHost | Test-DeployRuleSetCompliance | Repair-DeployRuleSetCompliance -ErrorAction SilentlyContinue

	Write-Host "vC: Assigning HostProfile to new host"
	$Assign = $VMHost | Apply-VMHostProfile -profile $HostProfile -Confirm:$false
	$Test = $VMHost | Test-VMHostProfileCompliance

	Write-Host "vC: VM Hypervisor Host: $($VMHost.Name) Ready to use, removing Maintenance mode"
	$Maint = $VMHost | Set-VMHost -State Connected

	# Logout of UCS
	Write-Host "UCS: Logging out of UCS: $ucs"
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