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
	
	
set-ucspowertoolconfiguration -supportmultipledefaultucs 0

# Global Variables
$ucs = "172.25.206.5"
$ucsuser = "ucs-ericwill\admin"
$ucspass = "Nbv12345!"
$vCenter = "172.25.206.186"
$vcuser = "Administrator"
$vcpass = "Nbv12345"
$tenantname = "CL2012"
$ucsorg = "org-root"
$WarningPreference = "SilentlyContinue"

Try
{
	# Login to UCS
	Write-Host "UCS: Logging into UCS Domain: $ucs"
	$ucspasswd = ConvertTo-SecureString $ucspass -AsPlainText -Force
	$ucscreds = New-Object System.Management.Automation.PSCredential ($ucsuser, $ucspasswd)
	$ucslogin = Connect-Ucs -Credential $ucscreds $ucs

	# Login to vCenter
	Write-Host "vC: Logging into vCenter: $vCenter"
	$vcenterlogin = Connect-VIServer $vCenter -User $vcuser -Password $vcpass | Out-Null

	# Create Initial ESXi Host
	Write-Host "UCS: Creating new Service Profile from SP Template: ${tenantname}"
	$sptemplate = get-ucsmanagedobject -dn $ucsorg | Get-UcsServiceProfile -Type "initial-template" -Name $tenantname
	$createnewsp = $sptemplate | Add-UcsServiceProfileFromTemplate -Count 1 -DestinationOrg (Get-UcsManagedObject -Dn $ucsorg) -Prefix "esxi-host"
	$spmacaddr = $createnewsp | Get-UcsVnic -Name eth0

	# Monitor UCS SP Associate for completion
	Write-Host "UCS: Waiting for UCS SP: $($createnewsp.name) to complete SP association process"
		do {
			if ( (Get-UcsManagedObject -Dn $createnewsp.Dn).AssocState -ieq "associated" )
			{
				break
			}
			Sleep 40
		} until ( (Get-UcsManagedObject -Dn $createnewsp.Dn).AssocState -ieq "associated" )
			
	Write-Host "UCS: Setting Desired Power State to 'up' of Service Profile: $($createnewsp.name)"
	$powerspon = $createnewsp | Set-UcsServerPower -State "up" -Force

	Write "vC: Waiting for VM Hypervisor Host with MAC address of $($spmacaddr.Addr) to connect to vCenter"
	do {
		Sleep 40
	} until ($VMHost = (get-vmhost -ErrorAction SilentlyContinue | foreach { $_.NetworkInfo.PhysicalNic | where { $_.Mac -ieq $spmacaddr.Addr } } | select -expandproperty vmhost ))

	Write "vC: Putting VM Hypervisor Host: $($VMhost.Name) into Maintenance mode"
	If ($VMHost.State -ne "Maintenance") {
		$Maint = $VMHost | Set-VMHost -State Maintenance
	}
	Write-Host "vC: Checking HostProfile Compliance against new VM Hypervisor Host: $($VMhost.Name) "
	$Test = $VMHost | Test-VMHostProfileCompliance

	Write-Host "vC: VM Hypervisor Host: $($VMhost.Name) Ready to use, removing Maintenance mode"
	$Maint = $VMHost | Set-VMHost -State Connected

	#Disconnect from UCS
	Write-Host "UCS: Logging out of UCS Domain: $ucs"
	$ucslogout = Disconnect-Ucs 

	#disconnect from vCenter
	Write-Host "vC: Logging out of vCenter: $vCenter"
	$vcenterlogout = Disconnect-VIServer $vCenter -confirm:$false
}
Catch
{
	 Write-Host "Error occurred in script:"
     Write-Host ${Error}
     exit
}