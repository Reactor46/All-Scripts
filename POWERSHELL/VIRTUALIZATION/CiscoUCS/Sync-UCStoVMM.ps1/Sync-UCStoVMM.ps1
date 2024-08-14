<#

.SYNOPSIS 
UCS to Microsoft Virtual Machine Manger 2012 Sync

.DESCRIPTION
This commandlet calls up your configuration within UCSM and configures Microsoft Virtual Machine Manager 2012.

.PARAMETER UCSIP
IP Address or comma seperated set if IP Addresses of the UCSM instance/s.  If an IP is entered here the UCSM .csv list will be ignored.

.PARAMETER UCSMCSV
Full path to a .csv formatted list of UCSM IP's to interact with.  Must be 2 rows with a "UCSM" header in row 1, and a "IP" header in row 2.

.PARAMETER ORGTARGET
Target Orginization you want to sync with.  If left blank will enumerate all service profiles in all orgs in the target UCSM Domain.

.PARAMETER VMMSERVER
NetBios, DNS, or IP address of the target Virtual Machine Manager 2012 instance.
You will need administrative priviledges on the VMM Server.

.EXAMPLE
Sync-UcsToVMM -UCSIP "10.5.11.254" -UCSIP "10.254.10.2" -ORGTARGET "Core-Testing" -vmmserver "vmm2012-sp1"

Sync UCS Organization "Core-Testing" to VMM Server "vmm2012-sp1"

.NOTES
Author: Chris Shockey, Architect, Cisco Advanced Services
Email: chris.shockey@cisco.com

Version: 0_9_2

.LINK
http://developer.cisco.com
#>
param(
      [string]$UCSIP,
      [string]$UCSMCSV,
	  [string]$VMMANAGER,
	  [string]$HOSTPROFILE,
	  [string]$ORGTARGET
)
#_______________________________________________________________________________
#__________________ GLOBALS_____________________________________________________
#_______________________________________________________________________________
# -ucsip "10.5.11.234" -vmmanager "UCS-VMM-01" -orgtarget "Cloud_Hosts"
$ReportErrorShowExceptionClass = $true
# It is not necessary to change any of these settings.
$global:rootDir = "C:\zUcsToVMMLog\"
[string]$VMMIPMIACCESSPROFILE = "VMMAccess"
[string]$VMMIPMIACCESSPROFILEUSER = "vmmIPMIuser"
[string]$password = "Cisco.123"
if (!($HOSTPROFILE))
{
	[string]$HOSTPROFILE = "W2012_CLU_HOST"
}
[string]$BMCPROTOCOL = "IPMI"
[string]$BMCDEFAULTPORT = 623
[array]$spWorking = $null

#_______________________________________________________________________________
#__________________ FUNCTIONS __________________________________________________
#_______________________________________________________________________________
Function Get-UCSMList 
{
	if (!($UCSIP) -and !($UCSMCSV))
	{
		Write-Host "Error: No UCSM instance specified or CSV of UCSM instances specified."
		Write-Host 'Use "-UCSIP YourIP,YourIP,YourIP", or "-ucsmcsv {csvPath}" to specify a UCSM instance or instances'
		Write-Host "Exiting..."
		exit 1
	}
	if ($UCSIP)
	{
		$ipsIn = $UCSIP.Split(",")
		foreach ($IP in $ipsIn)
		{
	    	[array]$mgrList += $IP
		}
	}
    if ($UCSMCSV)
    {
        if (Get-Item $UCSMCSV -ErrorAction SilentlyContinue)
        {
            $global:csvImport = Import-Csv $UCSMCSV
			$global:csvin = $true
			foreach ($ucsm in $csvImport)
			{
		    	[array]$mgrList += $ucsm.UCSIP
			}
        }
        else
        {
              Write-Host "Error: CSV File not found at $UCSMCSV. Please check the path"
              Write-Host "Exiting...."
              exit 1
        }
    }
    return [array]$mgrList
}
Function Validate-UcsFIConnect ([string]$ucsm)
{
	$ucsSession = $null
	if (!($cred))
	{
		Write-Host "Enter your UCS Credentials"
		$global:cred = Get-Credential -ErrorAction SilentlyContinue
	}
	$ucsSession = Connect-Ucs -Name $ucsm -Credential $cred -NoSsl -ErrorAction SilentlyContinue
	if (!($ucsSession))
	{
		$ucsSession = Connect-Ucs $ucsm -Credential $cred -ErrorAction SilentlyContinue
	}
	if (!($RunningCheck))
	{
		if (!($ucsSession))
		{
			Write-Host "Error Connecting to $ucsm, Most likely causes:"
			write-host "	- Bad Password"
			write-host "	- Invalid VIP to the target Fabric Interconnects"
			write-host "	- Bad Proxy, check your browser proxy settings"
		    Write-Host "Exiting..."
			Exit 1
		}
	}
	return $ucsSession
}
Function Validate-ScVMMLibrary2012
{
	if (!(Get-Module virtualmachinemanager -ErrorAction SilentlyContinue))
	{	
		Write-Host "Global: SCVMM 2012: Not Loaded, attempting to load..."
		Import-Module virtualmachinemanager -ErrorAction SilentlyContinue | Out-Null
		if (!(Get-Module virtualmachinemanager -ErrorAction SilentlyContinue))
		{
			Write-Host "Failed to load the SCVMM 2012 Library, please install the System Center Virtual Machine Manager Console on this computer."
			Write-Host "Exiting..."
			exit 1
		}
		else
		{
			Write-Host "Global: VMM 2012: Validate: Successfully imported the SCVMM 2012 Library.  Continuing..."
		}
	}
	else
	{
		Write-Host "Global: VMM 2012: Validate: 2012 VMM Library Loaded"
	}
}
Function Validate-ScVMMLibrary
{
	if (!(Get-PSSnapin -Name Microsoft.SystemCenter.VirtualMachineManager -ErrorAction SilentlyContinue)) 
	{ 
		Write-Host "Did not find VMM Snapping Loaded, Loading..."
		Add-PSSnapin -Name Microsoft.SystemCenter.VirtualMachineManager -ErrorAction SilentlyContinue
		if (!(Get-PSSnapin -Name Microsoft.SystemCenter.VirtualMachineManager -ErrorAction SilentlyContinue)) 
		{
			Write-Host "Global: VMM 2007: Snapin Failed loading.  Will try 2012 VMM Library..."
			Validate-ScVMMLibrary2012
		}
	}
}
Function Validate-VMMServerConnect ([string]$VMMANAGER)
{
	$vmmSession = $null
	Validate-ScVMMLibrary
	if (!($VMMANAGER))
	{	
		"The -VMMSERVER switch is empty.  You must enter a valid VMM server name or ip"
		Get-Help ".\Sync-UCStoVMM.ps1" -Detailed
		exit 1
	}
	if (!($vmmcred))
	{
		Write-Host "Please enter your VMM 2012 Server Admin Credentials..."
		$global:vmmcred = Get-Credential -ErrorAction SilentlyContinue
	}
	if (!($vmmSession))
	{
		$vmmSession = Get-SCVMMServer -ComputerName $VMMANAGER -Credential $vmmcred -ErrorAction SilentlyContinue
	}
	if (!($vmmSession))
	{
		Write-Host "Error Connecting to $VMMANAGER, Most likely causes:"
		write-host "	- Bad Password"
		write-host "	- Bad Proxy, check your browser proxy settings"
	    Write-Host "Exiting..."
		Exit 1
	}		
	return $vmmSession
}
Function Get-Orgs ($ucsSession)
{
	if (!($ORGTARGET))
	{
		[array]$a = Get-UcsOrg -Ucs $ucsSession | where {(Get-UcsServiceProfile -Org $_ -Type instance -LimitScope -AssignState assigned) -ne $null}
	}
	else
	{
		$a = Get-UcsOrg -Ucs $ucsSession -Name $ORGTARGET -ErrorAction SilentlyContinue
		if ($a)
		{
			return $a
		}
		else
		{
			Write-Host "$ORGTARGET : Does not exist on $($ucsSession.Ucs)"
		}
	}
	Return $a
}
Function Sync-UCSToVMM ([string]$org, [object]$vmmSession, $ucsSession)
{
	[array]$spWorking = Get-UcsServiceProfile -Org $org -Type instance -Ucs $ucsSession
	$i = 0
	foreach ($sp in $spWorking)
	{
		if ($sp)
		{
			$validate = New-UCSSpToVMM $sp $org $vmmSession $ucsSession
			$i++
		}
	}
	
	Write-Host "Global: Complete: $i Service Profiles processed under Organization: $org ."
}
Function New-UCSSpToVMM ($spIn, $org, [object]$vmmSession, $ucsSession)
{
	if (Get-SCVMHost -VMMServer $vmmSession -ComputerName $spIn.Name -ErrorAction SilentlyContinue)
	{
		Write-Host "	$($vmmSession.name): SP: Sync: $($spIn.Name) Already Exists, skipping."
		continue
	}
	else
	{
		# Get UUID via the IMPI Interface: Find-ScComputer
		# Get UUID from running host: Get-ScVmHost || $_.PhysicalComputer (Read-ScVmHost -
		##Read-SCVMHost -RefreshOutOfBandProperties
		#Find-SCComputer -
		Write-Host "	$($vmmSession.name): SP: Sync: Deploying $($spIn.Name) to $($vmmSession.Name) in $org"
		$blade = Get-UcsBlade -Dn $spIn.PnDn -Ucs $ucsSession
		$bladeMgmtController = Get-UcsMgmtController -Blade $blade -Ucs $ucsSession -Subject blade
		$mgmtIP = Get-UcsVnicIpV4PooledAddr -MgmtController $bladeMgmtController -Ucs $ucsSession
		Validate-UCSIPMIAccessProfileSet $spIn $ucsSession $vmmSession
		$bmcRunAs = Get-SCRunAsAccount -VMMServer $vmmSession -Name $VMMIPMIACCESSPROFILEUSER
		$hProfile = Get-SCVMHostProfile -VMMServer $vmmSession -Name $HOSTPROFILE
		$hGroup = Get-SCVMHostGroup -VMMServer $vmmSession -Name $org
		$u = $spIn.Uuid.ToCharArray()
		$uW = $spIn.Uuid.Split("-")
		$uuid = $u[6]+$u[7]+$u[4]+$u[5]+$u[2]+$u[3]+$u[0]+$u[1]+"-"+$u[11]+$u[12]+$u[9]+$u[10]+"-"+$u[16]+$u[17]+$u[14]+$u[15]+"-"+$uW[3]+"-"+$uW[4]
		New-SCVMHost -VMMServer $vmmSession -ComputerName $spIn.Name -BMCAddress $mgmtIP.Addr -BMCPort $BMCDEFAULTPORT -BMCProtocol $bmcProtocol `
					 -BMCRunAsAccount $bmcRunAs -VMHostProfile $hProfile -VMHostGroup $hGroup -SMBiosGuid $uuid -BypassADMachineAccountCheck -RunAsynchronously #| Out-Null
	}
}
Function Sync-UCStoVMMHostGroups ([string]$org, [object]$vmmSession)
{
	Write-Host "	$($vmmSession.name): VMM: Validate: Validating UCS Orginization $org exists in VMM."
	if (!(Get-SCVMHostGroup -Name $org))
	{
		Write-Host "	$($vmmSession.name): VMM: Configure: UCS Orginization $org does not exist, creating..." -NoNewline
		New-SCVMHostGroup -Name $org -VMMServer $vmmSession | Out-Null
		If (Get-SCVMHostGroup -Name $org)
		{	
			Write-Host ".Complete"
		}
		else
		{
			Write-Host ".Failed, exiting..."
			exit
		}
	}
}
Function Validate-UCSIPMIAccessProfileCreated ($ucsSession, $vmmSession)
{
	if ($ipmiProfileCreated -ne $true)
	{
		Write-Host "	$($vmmSession.name): IPMI: Validate: Validating IMPI profile is Created."
		if (!(Get-UCSIPmiAccessProfile -Name $VMMIPMIACCESSPROFILE -Ucs $ucsSession -ErrorAction SilentlyContinue))
		{	
			Write-Host "	$($vmmSession.name): IPMI: Configure: IMPI profile not present, creating..."
			# Add IPMI Access Profile and User to UCSM for use by VMM
			Add-UCSIPmiAccessProfile -Name $VMMIPMIACCESSPROFILE -Org "root" -Ucs $ucsSession
			if ($password)
			{
				$pass = $password
			}
			else
			{
				$pass = [Guid]::NewGuid().ToString().Replace("-","").Remove(20,12)
			}
			$ipmiPassword = convertto-securestring -string $pass -AsPlainText -Force
			Add-UcsAaaEpUser -IpmiAccessProfile $VMMIPMIACCESSPROFILE -Name $VMMIPMIACCESSPROFILEUSER -Pwd $pass -Priv admin -Ucs $ucsSession
			# Take created account and add a VMM Run As account using the information above.
			if (!(Get-SCRunAsAccount -VMMServer $vmmSession -Name $VMMIPMIACCESSPROFILEUSER -ErrorAction SilentlyContinue))
			{
				$credential = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $VMMIPMIACCESSPROFILEUSER,$ipmiPassword
				$runAsAccount = New-SCRunAsAccount -VMMServer $vmmSession -Credential $credential -Name $VMMIPMIACCESSPROFILEUSER | out-Null
			}
		}
		$global:ipmiProfileCreated = $true
	}
}
Function Validate-UCSIPMIAccessProfileSet($spIn, $ucsSession, $vmmSession)
{
	$spTemplate = $spin.SrcTemplName
	if ($spTemplate)
	{
		if ($spIn.MgmtAccessPolicyName -ne $VMMIPMIACCESSPROFILE)
		{
			$stTemplateIn = Get-UcsServiceProfile -Name $spTemplate -Ucs $ucsSession
			Set-UcsServiceProfile -ServiceProfile $stTemplateIn -MgmtAccessPolicyName $VMMIPMIACCESSPROFILE -Force -Ucs $ucsSession
		}
	}
	if ($spIn.MgmtAccessPolicyName -ne $VMMIPMIACCESSPROFILE)
	{
		Set-UcsServiceProfile -ServiceProfile $spin -MgmtAccessPolicyName $VMMIPMIACCESSPROFILE -Force -Ucs $ucsSession
	}
}
Function Sync-UCStoVMMVLans ($org, $ucsSession, $vmmSession) # INCOMPLETE/NOT-IMPLEMENTED
{
	$vlans = Get-UcsVlan -Ucs $ucsSession
#	$logicalNetwork = New-SCLogicalNetwork -Name "Blah"
#	$allHostGroups = @()
#	$allHostGroups += Get-SCVMHostGroup -ID "9a49b1d6-60d1-4abb-acb8-ccde5e657c72"
#	$allSubnetVlan = @()
#	$allSubnetVlan += New-SCSubnetVLan -Subnet "10.254.9.0/24" -VLanID 909
}
#_____________________________________________________________________________
#__________________MAIN PROGRAM ________________________________________________
#_______________________________________________________________________________
. {
	$UCSMList = Get-UCSMList
	#if (Get-Item $rootDir -ErrorAction SilentlyContinue)
	#{
	#	rd -Path $rootDir -Recurse -Force | out-Null
	#}
	#if ((!(Get-Item $rootDir -ErrorAction SilentlyContinue)))
	#{
	#	md -Path $rootDir -Force | out-Null
	#}
	foreach ($ucsm in $ucsmList)
	{
		[array]$ucsHandleList += Validate-UcsFIConnect $ucsm
	}
	foreach ($ucsSession in $ucsHandleList)
	{
		Write-Host "Global: Configuration: Beginning VMM Configuration"
		$vmmSession = Validate-VMMServerConnect $VMMANAGER
		[array]$orgs = Get-Orgs -ucsSession $ucsSession
		Validate-UCSIPMIAccessProfileCreated $ucsSession $vmmSession
		foreach ($org in $orgs)
		{
			Write-Host "	$($vmmSession.name): Org: Sync: Syncing UCS Organization $($org.Name) to VMM."
			Sync-UCStoVMMHostGroups $org.Name $vmmSession
			Write-Host "	$($vmmSession.name): Global SP: Sync: Syncing UCS Organization $($org.Name) Service Profiles to VMM."
			Sync-UCSToVMM -org $org.Name -vmmSession $vmmSession -ucsSession $ucsSession 
		}
	}
	Disconnect-Ucs -Ucs $ucsHandleList
}
