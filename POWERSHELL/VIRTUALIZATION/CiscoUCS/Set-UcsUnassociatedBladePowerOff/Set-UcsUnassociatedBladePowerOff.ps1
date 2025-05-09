﻿<#

.SYNOPSIS
	This script allows the user to power off a blade that is powered on but is not associated to a service profile

.DESCRIPTION
	This script allows the user to power off a blade that is powered on but is not associated to a service profile
	UCSM GUI or CLI does not have the capability to power off a blade that is not associated with a service profile
	This script creates a temporary service profile and with the default to power off the server
	The script then disassociates the service profile and deletes it

.EXAMPLE
	Set-UcsUnassociatedBladePowerOff.ps1
	This script can be run without any command line parameters.  User will be prompted for all parameters and options required

.EXAMPLE
	Set-UcsUnassociatedBladePowerOff.ps1 -ucs "1.2.3.4" -ucred -blade "1,7"
	-ucs -- UCS Manager IP address(s) or Host Name(s) -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-ucred -- UCS Manager Credential Switch -- Adding this switch will immediately prompt you for your UCSM username and password
	-blade -- UCS Chassis,Slot to power off. Example: 1,6
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords

.EXAMPLE
	Set-UcsUnassociatedBladePowerOff.ps1 -ucs "1.2.3.4" -saved "myucscred.csv" -blade "all" -skiperror
	-ucs -- UCS Manager IP address(s) or Host Name(s) -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-savedcred -- UCSM credentials file -- Example: -savedcred "myucscred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\myucscred.csv
		Make sure the password file is located in the same folder as the script
	-blade -- UCS Chassis,Slot to power off. Example: 1,6 or ALL to go through all blades that are powered on but are unassociated
	-skiperrors -- Tells the script to skip any prompts for errors and continues with 'y'
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords

.NOTES
	Author: Joe Martin
	Email: joemar@cisco.com
	Company: Cisco Systems, Inc.
	Version: v0.1.01
	Date: 6/8/2015
	Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
	***Provide any additional inputs here
	UCSM IP Address(s) or Hostname(s)
	UCSM Credentials Filename
	UCSM Username and Password
	UCSM Chassis and Slot

.OUTPUTS
	***None by default but your custom code may provide some.
	
.LINK
	http://communities.cisco.com/people/joemar/content

#>

#Command Line Parameters
param(
	[string]$UCSM,				# IP Address(s) or Hostname(s).  If multiple entries, separate by commas
	[switch]$UCREDENTIALS,		# UCSM Credentials (Username and Password).  Requires all domains to use the same credentials
	[string]$SAVEDCRED,			# Saved UCSM Credentials.  To create do: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} | Export-CSV -NoTypeInformation .\myucscred.csv
	[switch]$SKIPERROR,			# Skip any prompts for errors and continues with 'y'
	[string]$BLADE				# UCS Chassis and Blade: Example: 1,6 or ALL
)

#Clear the screen
clear-host

#Script kicking off
Write-Output "Script Running..."
Write-Output ""

#Tell the user what the script does
Write-Output "This script allows the user to power off a blade that is powered on but is not associated to a service profile"
Write-Output "UCSM GUI or CLI does not have the capability to power off a blade that is not associated with a service profile"
Write-Output "This script creates a temporary service profile and with the default to power off the server"
Write-Output "The script then disassociates the service profile and deletes it"
Write-Output ""

if ($UCREDENTIALS)
	{
		Write-Output "Enter UCSM Credentials"
		Write-Output ""
		$cred = Get-Credential -Message "Enter UCSM Credentials"
	}

#Change directory to the script root
cd $PSScriptRoot

#Check to see if credential files exists
if ($SAVEDCRED)
	{
		if ((Test-Path $SAVEDCRED) -eq $false)
			{
				Write-Output ""
				Write-Output "Your credentials file $SAVEDCRED does not exist in the script directory"
				Write-Output "	Exiting..."
				Disconnect-Ucs
				exit
			}
	}

#Do not show errors in script
$ErrorActionPreference = "SilentlyContinue"
#$ErrorActionPreference = "Stop"
#$ErrorActionPreference = "Continue"
#$ErrorActionPreference = "Inquire"

#Verify PowerShell Version for script support
Write-Output "Checking for proper PowerShell version"
$PSVersion = $psversiontable.psversion
$PSMinimum = $PSVersion.Major
if ($PSMinimum -ge "3")
	{
		Write-Output "	Your version of PowerShell is valid for this script."
		Write-Output "		You are running version $PSVersion"
		Write-Output ""
	}
else
	{
		Write-Output "	This script requires PowerShell version 3 or above"
		Write-Output "		You are running version $PSVersion"
		Write-Output "	Please update your system and try again."
		Write-Output "	You can download PowerShell updates here:"
		Write-Output "		http://search.microsoft.com/en-us/DownloadResults.aspx?rf=sp&q=powershell+4.0+download"
		Write-Output "	If you are running a version of Windows before 7 or Server 2008R2 you need to update to be supported"
		Write-Output "			Exiting..."
		Disconnect-Ucs
		exit
	}

#Load the UCS PowerTool
Write-Output "Checking Cisco PowerTool"
$PowerToolLoaded = $null
$Modules = Get-Module
$PowerToolLoaded = $modules.name
if ( -not ($Modules -like "ciscoUcsPs"))
	{
		Write-Output "	Loading Module: Cisco UCS PowerTool Module"
		Import-Module ciscoUcsPs
		$Modules = Get-Module
		if ( -not ($Modules -like "ciscoUcsPs"))
			{
				Write-Output ""
				Write-Output "	Cisco UCS PowerTool Module did not load.  Please correct his issue and try again"
				Write-Output "		Exiting..."
				exit
			}
		else
			{
				$PTVersion = (Get-Module CiscoUcsPs).Version
				Write-Output "		PowerTool version $PTVersion is now Loaded"
			}
	}
else
	{
		$PTVersion = (Get-Module CiscoUcsPs).Version
		Write-Output "	PowerTool version $PTVersion is already Loaded"
	}

#Select UCS Domain(s) for login
if ($UCSM -ne "")
	{
		$myucs = $UCSM
	}
else
	{
		$myucs = Read-Host "Enter UCS system IP or Hostname or a list of systems separated by commas"
	}
[array]$myucs = ($myucs.split(",")).trim()
if ($myucs.count -eq 0)
	{
		Write-Output ""
		Write-Output "You didn't enter anything"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Make sure we are disconnected from all UCS Systems
Disconnect-Ucs

#Test that UCSM(s) are IP Reachable via Ping
Write-Output ""
Write-Output "Testing PING access to UCSM"
foreach ($ucs in $myucs)
	{
		$ping = new-object system.net.networkinformation.ping
		$results = $ping.send($ucs)
		if ($results.Status -ne "Success")
			{
				Write-Output "	Can not access UCSM $ucs by Ping"
				Write-Output "		It is possible that a firewall is blocking ICMP (PING) Access.  Would you like to try to log in anyway?"
				if ($SKIPERROR)
					{
						$Try = "y"
					}
				else
					{
						$Try = Read-Host "Would you like to try to log in anyway? (Y/N)"
					}				if ($Try -ieq "y")
					{
						Write-Output "				Will try to log in anyway!"
					}
				elseif ($Try -ieq "n")
					{
						Write-Output ""
						Write-Output "You have chosen to exit"
						Write-Output "	Exiting..."
						Disconnect-Ucs
						exit
					}
				else
					{
						Write-Output ""
						Write-Output "You have provided invalid input"
						Write-Output "	Exiting..."
						Write-Output ""
						Disconnect-Ucs
						exit
					}			
			}
		else
			{
				Write-Output "	Successful access to $ucs by Ping"
			}
	}

#Log into the UCS System(s)
$multilogin = Set-UcsPowerToolConfiguration -SupportMultipleDefaultUcs $true
Write-Output ""
Write-Output "Logging into UCS"
#Verify PowerShell Version to pick prompt type
$PSVersion = $psversiontable.psversion
$PSMinimum = $PSVersion.Major
if (!$UCREDENTIALS)
	{
		if (!$SAVEDCRED)
			{
				if ($PSMinimum -ge "3")
					{
						Write-Output "	Enter your UCSM credentials"
						$cred = Get-Credential -Message "UCSM(s) Login Credentials" -UserName "admin"
					}
				else
					{
						Write-Output "	Enter your UCSM credentials"
						$cred = Get-Credential
					}
			}
		else
			{
				$CredFile = import-csv $SAVEDCRED
				$Username = $credfile.UserName
				$Password = $credfile.EncryptedPassword
				$cred = New-Object System.Management.Automation.PsCredential $Username,(ConvertTo-SecureString $Password)			
			}

	}
foreach ($myucslist in $myucs)
	{
		Write-Output "		Logging into: $myucslist"
		$myCon = $null
		$myCon = Connect-Ucs $myucslist -Credential $cred
		if (($mycon).Name -ne ($myucslist)) 
			{
				#Exit Script
				Write-Output "			Error Logging into this UCS domain"
				if ($myucs.count -le 1)
					{
						$continue = "n"
					}
				else
					{
						$continue = Read-Host "Continue without this UCS domain (Y/N)"
					}
				if ($continue -ieq "n")
					{
						Write-Output "				You have chosen to exit..."
						Write-Output ""
						Write-Output "Exiting Script..."
						Disconnect-Ucs
						exit
					}
				else
					{
						Write-Output "				Continuing..."
					}
			}
		else
			{
				Write-Output "			Login Successful"
			}
		sleep 1
	}
$myCon = (Get-UcsPSSession | measure).Count
if ($myCon -eq 0)
	{
		Write-Output ""
		Write-Output "You are not logged into any UCSM systems"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Collect SLOT and BLADE to power OFF
if (!$BLADE)
	{
		$BLADE = Read-Host "Enter CHASSIS,SLOT"
	}
if ($BLADE -ieq "all")
	{
		$Blades = Get-UcsBlade -OperState unassociated -OperPower on
	}
else
	{
		$BladeSplit = $BLADE.Split(",").trim(" ")
		$Chassis = $BladeSplit[0]
		$Slot = $BladeSplit[1]

		#Validate Blade is disassocated and powered on
		$Validate = Get-UcsBlade -Chassis $Chassis -SlotId $Slot
		Write-Output ""
		if (($Validate.Association -eq "none") -and ($Validate.OperPower -eq "on"))
			{
				Write-Output "This blade has been validated as disassociated and powered on"
				Write-Output ""
			}
		else
			{
				Write-Output "This server is not available to be modified or is already off"
				Write-Output "	Exiting..."
				Disconnect-Ucs
				exit
			}
	}
#Check to see if Service Profile already exists
$SP = Get-UcsServiceProfile | where {$_.Name -eq "turn_off_blade"}

if ($SP)
	{
		Write-Output "The Service Profile for powering off the blade already exists"
		Write-Output "	Skipping creation..."
		Write-Output ""
		if ($SP.AssocState -eq "associated")
			{
				$PnDn = $SP.PnDn
				Write-Output "The service profile is currently associated to $PnDn"
				Write-Output "	Disassociating the Service Profile from the blade"
				#Disassociate SP from Server
				$DontShow = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "turn_off_blade" -LimitScope | Get-UcsLsBinding | Remove-UcsLsBinding

				#Wait for SP Association to complete
				$Count = 0
				do
					{
						$IsDone = Get-UcsBlade -Chassis $Chassis -SlotId $Slot
						if ($IsDone.Association -eq "none")
							{
								$Associated = "n"
								Write-Output "		Disassociation Complete..."
								Write-Output ""
							}
						else
							{
								$Associated = "y"
								$Count++
								if ($Count -ge 120)
									{
										Write-Output "Service Profile failed to disassociate in the alloted time"
										Disconnect-Ucs
										exit
									}
								sleep -Seconds 10
							}
					}
				while ($Associated -eq "y")
			}
	}
else
	{
		#Create SP
		Write-Output "Creating Service Profile"
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsServiceProfile -AgentPolicyName "" -BiosProfileName "" -BootPolicyName "default" -Descr "" -DynamicConPolicyName "" -ExtIPPoolName "ext-mgmt" -ExtIPState "none" -HostFwPolicyName "" -IdentPoolName "" -LocalDiskPolicyName "default" -MaintPolicyName "" -MgmtAccessPolicyName "" -MgmtFwPolicyName "" -Name "turn_off_blade" -PolicyOwner "local" -PowerPolicyName "default" -ResolveRemote "yes" -ScrubPolicyName "" -SolPolicyName "" -SrcTemplName "" -StatsPolicyName "default" -UsrLbl "" -Uuid "FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF" -VconProfileName ""
		$mo_1 = $mo | Add-UcsVnicDefBeh -ModifyPresent -Action "none" -Descr "" -Name "" -NwTemplName "" -PolicyOwner "local" -Type "vhba"
		$mo_2 = $mo | Add-UcsVnicDefBeh -ModifyPresent -Action "none" -Descr "" -Name "" -NwTemplName "" -PolicyOwner "local" -Type "vnic"
		$mo_3 = $mo | Add-UcsVnicFcNode -ModifyPresent -Addr "pool-derived" -IdentPoolName "node-default"
		$mo_4 = $mo | Set-UcsServerPower -State "admin-down" -Force
		$mo_5 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "1" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
		$mo_6 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "2" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
		$mo_7 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "3" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
		$mo_8 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "4" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
		$DontShow = Complete-UcsTransaction
		$SP = Get-UcsServiceProfile | where {$_.Name -eq "turn_off_blade"}
		if (!$SP)
			{
				Write-Output "	Service Profile Failed to Create"
				Write-Output "		Exiting..."
				Disconnect-Ucs
				exit
			}
		else
			{
				Write-Output "	Complete"
				Write-Output ""
			}
	}

function PowerOff($Chassis, $Slot)
	{
		Write-Host "Associating Service Profile to blade ("$Chassis"/"$Slot" )"
		$PnDn = "sys/chassis-$Chassis/blade-$Slot"
		$DontShow = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "turn_off_blade" -LimitScope | Add-UcsLsBinding -ModifyPresent  -PnDn $PnDn -RestrictMigration "no"

		#Wait for SP Association to complete
		$Count = 0
		do
			{
				$IsDone = Get-UcsBlade -Chassis $Chassis -SlotId $Slot
				if ($IsDone.Association -eq "associated")
					{
						$Associated = "y"
						Write-Output "	Complete"
						Write-Output ""
					}
				else
					{
						$Associated = "n"
						$Count++
						if ($Count -ge 120)
							{
								Write-Output "	Service Profile failed to associate in the alloted time"
								Write-Output "		Deleting Service Profile..."
								$DontShow = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "turn_off_blade" -LimitScope | Remove-UcsServiceProfile
								Disconnect-Ucs
								Write-Output "			Exiting..."
								exit
							}
						sleep -Seconds 10
					}
			}
		while ($Associated -eq "n")

		#Disassociate SP from Server
		Write-Host "Disassociating Service Profile from blade ("$Chassis"/"$Slot" )"
		$DontShow = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "turn_off_blade" -LimitScope | Get-UcsLsBinding | Remove-UcsLsBinding -Force

		#Wait for SP Association to complete
		$Count = 0
		do
			{
				$IsDone = Get-UcsBlade -Chassis $Chassis -SlotId $Slot
				if ($IsDone.Association -eq "none")
					{
						$Associated = "n"
						Write-Output "	Complete"
						Write-Output ""
					}
				else
					{
						$Associated = "y"
						$Count++
						if ($Count -ge 120)
							{
								Write-Output "	Service Profile failed to associate in the alloted time"
								Write-Output "		Deleting Service Profile..."
								$DontShow = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "turn_off_blade" -LimitScope | Remove-UcsServiceProfile -Force
								Disconnect-Ucs
								Write-Output "			Exiting..."
								exit
							}
							sleep -Seconds 10
					}
			}
		while ($Associated -eq "y")
	}

#Associate SP with Server
if ($BLADE -ieq "all")
	{
		foreach ($Item in $Blades)
			{
				PowerOff $Item.ChassisID $Item.SlotId
			}
	}
else
	{
		PowerOff $Chassis $Slot
	}

#Delete Service Profile
Write-Output "Deleting Service Profile"
$DontShow = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "turn_off_blade" -LimitScope | Remove-UcsServiceProfile -Force

#Check to see if Service Profile has been removed
$SP = Get-UcsServiceProfile | where {$_.Name -eq "turn_off_blade"}

if (!$SP)
	{
		Write-Output "	Complete"
		Write-Output ""
	}
else
	{
		Write-Output "	The Service Profile could not be deleted.  Please do this manually"
		Write-Output "		FAILED"
	}

#Validate Blade is disassocated and powered on
if ($BLADE -ieq "all")
	{
		foreach ($Item in $Blades)
			{
				$Validate = Get-UcsBlade -Chassis $item.ChassisId -SlotId $Item.SlotId
				if ($Validate.OperPower -eq "off")
					{
						Write-Host "This blade ("$item.ChassisID"/"$item.SlotId") has been validated as powered off"
						Write-Output "	SUCCESS"
						Write-Output ""
					}
				else
					{
						Write-Host "This blade ("$item.ChassisID"/"$item.SlotId") is still powered on....ARGH!"
						Write-Output "	FAILED"
						Write-Output ""
					}		
			}
	}
else
	{
		$Validate = Get-UcsBlade -Chassis $Chassis -SlotId $Slot
		if ($Validate.OperPower -eq "off")
			{
				Write-Output "This blade has been validated as powered off"
				Write-Output "	SUCCESS"
			}
		else
			{
				Write-Output "This server is still powered on....ARGH!"
				Write-Output "	FAILED"
				Disconnect-Ucs
				exit
			}
	}
	
#Disconnect from UCSM(s)
Disconnect-Ucs

#Exit the Script
Write-Output ""
Write-Output "Script Complete"
exit