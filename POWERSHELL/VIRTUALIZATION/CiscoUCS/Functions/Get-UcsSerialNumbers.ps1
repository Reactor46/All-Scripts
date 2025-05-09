﻿<#

.SYNOPSIS
	Collect all serial numbers in UCS domain(s) and write them to a formatted Excel spreadsheet

.DESCRIPTION
	This script will log into a single or multiple UCS domains and collect all serial numbers and output the collected information to a formatted Excel Spreadsheet
	This script can take a LONG time to run due to the amount of data collected and the speed of writing information to Excel.  Please be patient!

.EXAMPLE
	Get-UcsSerialNumbers.ps1
	This script can be run without any command line parameters.  User will be prompted for all parameters and options required

.EXAMPLE
	Get-UcsSerialNumbers.ps1 -ucs "1.2.3.4" -ucred
	-ucs -- UCS Manager IP or Host Name -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-ucred -- UCS Manager Credential Switch -- Adding this switch will immediately prompt you for your UCSM username and password
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords
	
.EXAMPLE
	Get-UcsSerialNumbers.ps1 -ucs "1.2.3.4" -saved "myucscred.csv" -skiperrors
	-ucs -- UCS Manager IP or Host Name -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-savedcred -- UCSM credentials file -- Example: -savedcred "myucscred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\myucscred.csv
		Make sure the password file is located in the same folder as the script
	-skiperrors -- Tells the script to skip any prompts for errors and continues with 'y'
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords

.NOTES
	Author: Joe Martin
	Email: joemar@cisco.com
	Company: Cisco Systems, Inc.
	Version: v0.9.10
	Date: 8/12/2014
	Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
	UCSM IP Address(s) or Hostname(s)
	UCSM Username and Password
	UCSM Credentials Filename

.OUTPUTS
	Output will be sent to Excel.  This requires Microsoft Excel to be installed on the client workstation running this script
	File format are .XLS
	
.LINK
	http://communities.cisco.com/people/joemar/content

#>

#Command Line Parameters
param(
	[string]$UCSM,				# IP Address(s) or Hostname(s).  If multiple entries, separate by commas
	[switch]$UCREDENTIALS,		# UCSM Credentials (Username and Password)
	[string]$SAVEDCRED,			# Saved UCSM Credentials.  To create do: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} | Export-CSV -NoTypeInformation .\myucscred.csv
	[switch]$SKIPERROR			# Skip any prompts for errors and continues with 'y'
)

#Clear the screen
clear-host

#Script kicking off
Write-Output "Script Running..."

#Gather any credentials requested from command line
if ($UCREDENTIALS)
	{
		Write-Output ""
		Write-Output "Enter UCSM Credentials"
		$cred = Get-Credential -Message "Enter UCSM Credentials"
	}

#Change directory to the script root
cd $PSScriptRoot

#Tell the user what the script does
Write-Output ""
Write-Output "This script allows you to log into a single or multiple UCS Domains and it will"
Write-Output "create an Excel spreadsheet of all the devices and their serial number along"
Write-Output "with other relevant associated information."
Write-Output ""
Write-Output "This script collects a lot of data and writes it to Excel and can take a LONG time to execute"
Write-Output ""
Write-Output "Prerequisites:"
Write-Output "	PowerShell v3 enabled on your client machine"
Write-Output "	Network access to your UCSM"
Write-Output "	An account on your UCSM"
Write-Output "	Excel installed on the client machine"
Write-Output "	Cisco PowerTool for PowerShell installed in the client machine."
Write-Output "		It can be downloaded at http:\\www.cisco.com"

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
$PSVersion = $psversiontable.psversion
$PSMinimum = $PSVersion.Major
if ($PSMinimum -ge "3")
	{
	}
else
	{
		Write-Output "This script requires PowerShell version 3 or above"
		Write-Output "Please update your system and try again."
		Write-Output "You can download PowerShell updates here:"
		Write-Output "	http://search.microsoft.com/en-us/DownloadResults.aspx?rf=sp&q=powershell+4.0+download"
		Write-Output "If you are running a version of Windows before 7 or Server 2008R2 you need to"
		Write-Output "update to be supported"
		Write-Output "		Exiting..."
		Disconnect-Ucs
		exit
	}

#Load the UCS PowerTool
Write-Output ""
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
				Write-Output "	Cisco UCS PowerTool did not load.  Please correct his issue and try again"
				Write-Output "		Exiting..."
				exit
			}
		else
			{
				Write-Output "	PowerTool is Loaded"
			}
	}
else
	{
		Write-Output "	PowerTool is Loaded"
	}

#Define UCS Domain(s)
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

#Make sure we are disconnected from other UCS Systems
Disconnect-Ucs

#Test that UCSM is IP Reachable via Ping
Write-Output ""
Write-Output "Testing PING access to UCSM"
foreach ($ucs in $myucs)
	{
		$ping = new-object system.net.networkinformation.ping
		$results = $ping.send($ucs)
		if ($results.Status -ne "Success")
			{
				Write-Output "	Can not access UCSM $ucs by Ping"
				Write-Output ""
				Write-Output "It is possible that a firewall is blocking ICMP (PING) Access."
				Write-Output "	Would you like to try to log in anyway?"
				if ($SKIPERROR)
					{
						$Try = "y"
					}
				else
					{
						$Try = Read-Host "Would you like to try to log in anyway? (Y/N)"
					}
				if ($Try -ieq "y")
					{
						Write-Output ""
						Write-Output "Trying to log in anyway!"
						Write-Output ""
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
				Write-Output ""
			}
	}

#Log into the UCS System
$multilogin = Set-UcsPowerToolConfiguration -SupportMultipleDefaultUcs $true
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
		Write-Output ""
		Write-Output "logging into: $myucslist"
		$myCon = $null
		$myCon = Connect-Ucs $myucslist -Credential $cred
		if (($mycon).Name -ne ($myucslist)) 
			{
				#Exit Script
				Write-Output "     Error Logging into this UCS domain"
				if ($myucs.count -le 1)
					{
						$continue = "n"
					}
				else
					{
						if ($SKIPERROR)
							{
								$continue = "y"
							}
						else
							{
								$continue = Read-Host "Continue without this UCS domain (Y/N)"
							}
					}
				if ($continue -ieq "n")
					{
						Write-Output "Exiting Script..."
						Disconnect-Ucs
						exit
					}
				else
					{
						Write-Output "		Continuing..."
					}
			}
		else
			{
				Write-Output "     Login Successful"
			}
		sleep 1
	}
if ($myCon.count -eq 0)
	{
		Write-Output ""
		Write-Output "You are not logged into any UCSM systems"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Launch Excel application
Write-Output ""
Write-Output "Launching Excel in the background (Hidden)"
$Excel = New-Object -comobject Excel.Application

#Keeps Excel hidden in the background
$Excel.Visible = $False

#Create a new Excel workbook
Write-Output "	Creating new workbook"
$Workbook = $Excel.Workbooks.Add()

#Create a new Excel worksheet
$Worksheet += $Workbook.Worksheets.Item(1)

#Rename the Excel worksheet
Write-Output "	Creating worksheet: Serial Numbers"
$Worksheet.Name = "Serial Numbers"

#Create Excel headers
Write-Output "	Setting worksheet headers"
$Worksheet.Cells.Item(1,1) = "Category"
$Worksheet.Cells.Item(1,2) = "UCS Domain"
$Worksheet.Cells.item(1,3) = "Name"
$Worksheet.Cells.item(1,4) = "Service Profile"
$Worksheet.Cells.item(1,5) = "Model Number"
$Worksheet.Cells.item(1,6) = "Serial Number"
$Worksheet.Cells.item(1,7) = "MFG Date"

#Format Excel cell headers
Write-Output "	Formatting cells"
$Worksheet.Cells.item(1,1).font.size=12
$Worksheet.Cells.item(1,1).font.bold=$true
$Worksheet.Cells.item(1,1).font.underline=$true
$Worksheet.Cells.item(1,2).font.size=12
$Worksheet.Cells.item(1,2).font.bold=$true
$Worksheet.Cells.item(1,2).font.underline=$true
$Worksheet.Cells.item(1,3).font.size=12
$Worksheet.Cells.item(1,3).font.bold=$true
$Worksheet.Cells.item(1,3).font.underline=$true
$Worksheet.Cells.item(1,4).font.size=12
$Worksheet.Cells.item(1,4).font.bold=$true
$Worksheet.Cells.item(1,4).font.underline=$true
$Worksheet.Cells.item(1,5).font.size=12
$Worksheet.Cells.item(1,5).font.bold=$true
$Worksheet.Cells.item(1,5).font.underline=$true
$Worksheet.Cells.item(1,6).font.size=12
$Worksheet.Cells.item(1,6).font.bold=$true
$Worksheet.Cells.item(1,6).font.underline=$true
$Worksheet.Cells.item(1,7).font.size=12
$Worksheet.Cells.item(1,7).font.bold=$true
$Worksheet.Cells.item(1,7).font.underline=$true
$Worksheet.columns.item(1).columnWidth = 30
$Worksheet.columns.item(2).columnWidth = 35
$Worksheet.columns.item(3).columnWidth = 45
$Worksheet.columns.item(4).columnWidth = 35
$Worksheet.columns.item(5).columnWidth = 40
$Worksheet.columns.item(6).columnWidth = 25
$Worksheet.columns.item(7).columnWidth = 25

#Gather info from UCS
Write-Output ""
Write-Output "Gathering information from UCSM...(Please Wait)"
$FIs = Get-UcsNetworkElement | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,RN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$FIModules = Get-UcsFiModule | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$Chassis = Get-UcsChassis | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$IOM = Get-UcsIom | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$FEX = Get-UcsFex | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$Blade = Get-UcsBlade | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$Rack = Get-UcsRackUnit | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$Adapter = Get-UcsAdaptorUnit | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$AdapterExpansion = Get-UcsAdaptorUnitExtn | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$DIMMs = Get-UcsMemoryUnit | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null) -and ($_.Serial -ne "NO DIMM"))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$StorageController = Get-UcsStorageController | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$Storage = Get-UcsStorageLocalDisk | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$Fan = Get-UcsFan | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,MfgTime,Serial
Write-Host "." -NoNewline
$PSU = Get-UcsPSU | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,Serial
Write-Host "." -NoNewline
$Disk = Get-UcsStorageLocalDisk | where {(($_.Serial -ne "") -and ($_.Serial -ne "N/A") -and ($_.Serial -ne $null))} | Sort-Object -CaseSensitive -Property Ucs,DN,Model,Serial
Write-Host "." -NoNewline
$License = Get-UcsLicense | Sort-Object -CaseSensitive -Property Ucs,Scope,AbsQuant,DefQuant,UsedQuant,GracePeriodUsed,OperState,PeerStatus
Write-Host "." -NoNewline
Write-Host ""

#Put important data into tables
Write-Output ""
Write-Output "Creating tables of data...(Please Wait)"
$FITable = @{"UCS" = $FIs.Ucs ; "DN" = $FIs.RN ; "Model" = $FIs.Model ; "Serial" = $FIs.Serial ; "ServiceProfile" = "N/A" ; "MfgTime" = $FIs.MfgTime}
Write-Host "." -NoNewline
$FIModuleTable = @{"UCS" = $FIModules.Ucs ; "DN" = $FIModules.Dn ; "Model" = $FIModules.Model ; "Serial" = $FIModules.Serial ; "ServiceProfile" = "N/A" ; "MfgTime" = $FIModules.MfgTime}
Write-Host "." -NoNewline
$ChassisTable = @{"UCS" = $Chassis.Ucs ; "DN" = $Chassis.Dn ; "Model" = $Chassis.Model ; "Serial" = $Chassis.Serial ; "ServiceProfile" = "N/A" ; "MfgTime" = $Chassis.MfgTime}
Write-Host "." -NoNewline
$IOMTable = @{"UCS" = $IOM.Ucs ; "DN" = $IOM.Dn ; "Model" = $IOM.Model ; "Serial" = $IOM.Serial ; "ServiceProfile" = "N/A" ; "MfgTime" = $IOM.MfgTime}
Write-Host "." -NoNewline
$FEXTable = @{"UCS" = $FEX.Ucs ; "DN" = $FEX.Dn ; "Model" = $FEX.Model ; "Serial" = $FEX.Serial ; "ServiceProfile" = "N/A" ; "MfgTime" = $FEX.MfgTime}
Write-Host "." -NoNewline
$BladeTable = @{"UCS" = $Blade.Ucs ; "DN" = $Blade.Dn ; "Model" = $Blade.Model ; "Serial" = $Blade.Serial ; "ServiceProfile" = $Blade.AssignedToDn ; "MfgTime" = $Blade.MfgTime}
Write-Host "." -NoNewline
$RackTable = @{"UCS" = $Rack.Ucs ; "DN" = $Rack.Dn ; "Model" = $Rack.Model ; "Serial" = $Rack.Serial ; "ServiceProfile" = $Rack.AssignedToDn ; "MfgTime" = $Rack.MfgTime}
Write-Host "." -NoNewline
$LicenseTable = @{"UCS" = $License.Ucs ; "FI" = $License.Scope ; "Abs" = $License.AbsQuant ; "Def" = $License.DefQuant ; "Used" = $License.UsedQuant ; "Grace" = $License.GracePeriodUsed ; "Oper" = $License.OperState ; "Peer" = $License.PeerStatus}
Write-Host "." -NoNewline

$AdapterLoop = ($Adapter.Dn).count
$Loop = 1
$SPArray = @()
do
	{
		$BladeFull = $Adapter.Dn[$Loop - 1] -match "(?<content>.*)/adaptor-"
		$BladeReadable = $matches['content']
		$ServiceProfile = Get-UcsServiceProfile -Ucs $Adapter.Ucs[$Loop - 1] -PnDn $BladeReadable
		if ($ServiceProfile -eq $null)
			{
				$SPArray += "<<UNASSOCIATED>>"
			}
		else
			{
				$SPArray += $ServiceProfile.Name
			}
		Write-Host "." -NoNewline
		$Loop++
	}
while
	(
		$Loop -le $AdapterLoop
	)
$AdapterTable = @{"UCS" = $Adapter.Ucs ; "DN" = $Adapter.Dn ; "Model" = $Adapter.Model ; "Serial" = $Adapter.Serial ; "ServiceProfile" = $SPArray ; "MfgTime" = $Adapter.MfgTime}
Write-Host "." -NoNewline
$AdapterExpansionLoop = ($AdapterExpansion.Dn).count
$Loop = 1
$SPArray = @()
do
	{
		$AdapterExpansionFull = $AdapterExpansion.Dn[$Loop - 1] -match "(?<content>.*)/adaptor-[0-9]"
		$AdapterExpansionReadable = $matches['content']
		$ServiceProfile = Get-UcsServiceProfile -Ucs $AdapterExpansion.Ucs[$Loop - 1] -PnDn $AdapterExpansionReadable
		if ($ServiceProfile -eq $null)
			{
				$SPArray += "<<UNASSOCIATED>>"
			}
		else
			{
				$SPArray += $ServiceProfile.Name
			}
		Write-Host "." -NoNewline
		$Loop++
	}
while
	(
		$Loop -le $AdapterExpansionLoop
	)
$AdapterExpansionTable = @{"UCS" = $AdapterExpansion.Ucs ; "DN" = $AdapterExpansion.Dn ; "Model" = $AdapterExpansion.Model ; "Serial" = $AdapterExpansion.Serial ; "ServiceProfile" = $SPArray ; "MfgTime" = $AdapterExpansion.MfgTime}
Write-Host "." -NoNewline
$DIMMsLoop = ($DIMMs.Dn).count
$Loop = 1
$SPArray = @()
do
	{
		$DIMMsFull = $DIMMs.Dn[$Loop - 1] -match "(?<content>.*)/board"
		$DIMMsReadable = $matches['content']
		$ServiceProfile = Get-UcsServiceProfile -Ucs $DIMMs.Ucs[$Loop - 1] -PnDn $DIMMsReadable
		if ($ServiceProfile -eq $null)
			{
				$SPArray += "<<UNASSOCIATED>>"
			}
		else
			{
				$SPArray += $ServiceProfile.Name
			}
		Write-Host "." -NoNewline
		$Loop++
	}
while
	(
		$Loop -le $DIMMsLoop
	)
$DIMMsTable = @{"UCS" = $DIMMs.Ucs ; "DN" = $DIMMs.Dn ; "Model" = $DIMMs.Model ; "Serial" = $DIMMs.Serial ; "ServiceProfile" = $SPArray ; "MfgTime" = $DIMMs.MfgTime}
Write-Host "." -NoNewline
$StorageControllerLoop = ($StorageController.Dn).count
$Loop = 1
$SPArray = @()
do
	{
		$StorageControllerFull = $StorageController.Dn[$Loop - 1] -match "(?<content>.*)/board"
		$StorageControllerReadable = $matches['content']
		$ServiceProfile = Get-UcsServiceProfile -Ucs $StorageController.Ucs[$Loop - 1] -PnDn $StorageControllerReadable
		if ($ServiceProfile -eq $null)
			{
				$SPArray += "<<UNASSOCIATED>>"
			}
		else
			{
				$SPArray += $ServiceProfile.Name
			}
		Write-Host "." -NoNewline
		$Loop++
	}
while
	(
		$Loop -le $StorageControllerLoop
	)
$StorageControllerTable = @{"UCS" = $StorageController.Ucs ; "DN" = $StorageController.Dn ; "Model" = $StorageController.Model ; "Serial" = $StorageController.Serial ; "ServiceProfile" = $SPArray ; "MfgTime" = $StorageController.MfgTime}
Write-Host "." -NoNewline
$StorageLoop = ($Storage.Dn).count
$Loop = 1
$SPArray = @()
do
	{
		$StorageFull = $Storage.Dn[$Loop - 1] -match "(?<content>.*)/board"
		$StorageReadable = $matches['content']
		$ServiceProfile = Get-UcsServiceProfile -Ucs $Storage.Ucs[$Loop - 1] -PnDn $StorageReadable
		if ($ServiceProfile -eq $null)
			{
				$SPArray += "<<UNASSOCIATED>>"
			}
		else
			{
				$SPArray += $ServiceProfile.Name
			}
		Write-Host "." -NoNewline
		$Loop++
	}
while
	(
		$Loop -le $StorageLoop
	)
$StorageTable = @{"UCS" = $Storage.Ucs ; "DN" = $Storage.Dn ; "Model" = $Storage.Model ; "Serial" = $Storage.Serial ; "ServiceProfile" = $SPArray ; "MfgTime" = $Storage.MfgTime}
Write-Host "." -NoNewline
$FanTable = @{"UCS" = $Fan.Ucs ; "DN" = $Fan.Dn ; "Model" = $Fan.Model ; "Serial" = $Fan.Serial ; "MfgTime" = $Fan.MfgTime}
Write-Host "." -NoNewline
$PSULoop = ($PSU.Dn).count
$Loop = 1
$SPArray = @()
do
	{
		$PSUFull = $PSU.Dn[$Loop - 1] -match "(?<content>.*)/psu"
		$PSUReadable = $matches['content']
		$ServiceProfile = Get-UcsServiceProfile -Ucs $PSU.Ucs[$Loop - 1] -PnDn $PSUReadable
		if (($ServiceProfile -eq $null) -and ($PSUReadable.SubString(0,7) -eq "sys/rac"))
			{
				$SPArray += "<<UNASSOCIATED>>"
			}
		elseif ($ServiceProfile -ne $null)
			{
				$SPArray += $ServiceProfile.Name
			}
		else
			{
				$SPArray += "N/A"
			}
		Write-Host "." -NoNewline
		$Loop++
	}
while
	(
		$Loop -le $PSULoop
	)
$PSUTable = @{"UCS" = $PSU.Ucs ; "DN" = $PSU.Dn ; "Model" = $PSU.Model ; "Serial" = $PSU.Serial ; "ServiceProfile" = $SPArray ; "MfgTime" = $PSU.MfgTime}
$DiskLoop = ($Disk.Dn).count
$Loop = 1
$SPArray = @()
do
	{
		$DiskFull = $Disk.Dn[$Loop - 1] -match "(?<content>.*)/board"
		$DiskReadable = $matches['content']
		$ServiceProfile = Get-UcsServiceProfile -Ucs $Disk.Ucs[$Loop - 1] -PnDn $DiskReadable
		if ($ServiceProfile -eq $null)
			{
				$SPArray += "<<UNASSOCIATED>>"
			}
		elseif ($ServiceProfile -ne $null)
			{
				$SPArray += $ServiceProfile.Name
			}
		else
			{
				$SPArray += "N/A"
			}
		Write-Host "." -NoNewline
		$Loop++
	}
while
	(
		$Loop -le $DiskLoop
	)
$DiskTable = @{"UCS" = $Disk.Ucs ; "DN" = $Disk.Dn ; "Model" = $Disk.Model ; "Serial" = $Disk.Serial ; "ServiceProfile" = $SPArray ; "MfgTime" = "N/A"}
Write-Host "." -NoNewline
Write-Host ""

#Setting values for data formatting
$Loop = 1
$FILoop = $FITable.UCS.Count
$FIModuleLoop = $FIModuleTable.UCS.Count
$ChassisLoop = $ChassisTable.UCS.Count
$IOMLoop = $IOMTable.UCS.Count
$FEXLoop = $FEXTable.UCS.Count
$BladeLoop = $BladeTable.UCS.Count
$RackLoop = $RackTable.UCS.Count
$AdapterLoop = $AdapterTable.UCS.Count
$AdapterExpansionLoop = $AdapterExpansionTable.UCS.Count
$DIMMsLoop = $DIMMsTable.UCS.Count
$StorageControllerLoop = $StorageControllerTable.UCS.Count
$StorageLoop = $StorageTable.UCS.Count
$FanLoop = $FanTable.UCS.Count
$PSULoop = $PSUTable.UCS.Count
$LicenseLoop = $License.UCS.Count

#Writing Fabric Interconnect information to Excel
if ($FILoop -ne 0)
	{
		Write-Output ""
		Write-Output "Writing Fabric Interconnect information to Excel...(Please Wait)"
		$DoLoop = 1
		$Worksheet.Cells.item($Loop+1,1) = "'Fabric Interconnect(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($FILoop -eq 1)
					{
						$Worksheet.Cells.Item($Loop+1,2) = "'"+$FITable.UCS #UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$FITable.DN #Name
						$Worksheet.Cells.item($Loop+1,4) = "'N/A" #Service Profile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$FITable.Model #Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$FITable.Serial #Serial
						if (($FITable.MfgTime -eq "") -or ($FITable.MfgTime -eq $null) -or ($FITable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($FITable.MfgTime -replace "T", "  ") #Manufacturing Time
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$Worksheet.Cells.Item($Loop+1,2) = "'"+$FITable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$FITable.DN[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,4) = "'N/A" #Service Profile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$FITable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$FITable.Serial[$DoLoop-1]
						if ($FITable.MfgTime.Count -ne 0)
							{
								if (($FITable.MfgTime[$DoLoop-1] -eq "") -or ($FITable.MfgTime[$DoLoop-1] -eq $null) -or ($FITable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($FITable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $FILoop)
	}

#Writing Fabric Interconnect Module information to Excel
if ($FIModuleLoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing Fabric Interconnect Module information to Excel...(Please Wait)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'Fabric Interconnect Module(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($FIModuleLoop -eq 1)
					{
						$DNFull = $FIModuleTable.DN -match "/(?<content>.*)/"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$FIModuleTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'N/A" #Service Profile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$FIModuleTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$FIModuleTable.Serial
						if (($FIModuleTable.MfgTime -eq "") -or ($FIModuleTable.MfgTime -eq $null) -or ($FIModuleTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($FIModuleTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $FIModuleTable.DN[$DoLoop-1] -match "/(?<content>.*)/"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$FIModuleTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'N/A" #Service Profile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$FIModuleTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$FIModuleTable.Serial[$DoLoop-1]
						if ($FIModuleTable.MfgTime.Count -ne 0)
							{
								if (($FIModuleTable.MfgTime[$DoLoop-1] -eq "") -or ($FIModuleTable.MfgTime[$DoLoop-1] -eq $null) -or ($FIModuleTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($FIModuleTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $FIModuleLoop)
	}

#Writing Chassis information to Excel
if ($ChassisLoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing Chassis information to Excel...(Please Wait)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'Chassis(')"
		do
			{
				Write-Host "." -NoNewline
				if ($ChassisLoop -eq 1)
					{
						$DNFull = $ChassisTable.DN -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$ChassisTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'N/A" #Service Profile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$ChassisTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$ChassisTable.Serial
						if (($ChassisTable.MfgTime -eq "") -or ($ChassisTable.MfgTime -eq $null) -or ($ChassisTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($ChassisTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $ChassisTable.DN[$DoLoop-1] -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$ChassisTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'N/A" #Service Profile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$ChassisTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$ChassisTable.Serial[$DoLoop-1]
						if ($ChassisTable.MfgTime.Count -ne 0)
							{
								if (($ChassisTable.MfgTime[$DoLoop-1] -eq "") -or ($ChassisTable.MfgTime[$DoLoop-1] -eq $null) -or ($ChassisTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($ChassisTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $ChassisLoop)
	}

#Writing IOM information to Excel
if ($IOMLoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing IOM information to Excel...(Please Wait)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'IOM(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($IOMLoop -eq 1)
					{
						$DNFull = $IOMTable.DN -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$IOMTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'N/A" #Service Profile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$IOMTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$IOMTable.Serial
						if (($IOMTable.MfgTime -eq "") -or ($IOMTable.MfgTime -eq $null) -or ($IOMTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($IOMTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $IOMTable.DN[$DoLoop-1] -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$IOMTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'N/A" #Service Profile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$IOMTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$IOMTable.Serial[$DoLoop-1]
						if ($IOMTable.MfgTime.Count -ne 0)
							{
								if (($IOMTable.MfgTime[$DoLoop-1] -eq "") -or ($IOMTable.MfgTime[$DoLoop-1] -eq $null) -or ($IOMTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($IOMTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $IOMLoop)
	}

#Writing FEX information to Excel
if ($FEXLoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing FEX information to Excel...(Please Wait)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'FEX(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($FEXLoop -eq 1)
					{
						$DNFull = $FEXTable.DN -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$FEXTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'N/A" #Service Profile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$FEXTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$FEXTable.Serial
						if (($FEXTable.MfgTime -eq "") -or ($FEXTable.MfgTime -eq $null) -or ($FEXTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($FEXTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $FEXTable.DN[$DoLoop-1] -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$FEXTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'N/A" #Service Profile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$FEXTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$FEXTable.Serial[$DoLoop-1]
						if ($FEXTable.MfgTime.Count -ne 0)
							{
								if (($FEXTable.MfgTime[$DoLoop-1] -eq "") -or ($FEXTable.MfgTime[$DoLoop-1] -eq $null) -or ($FEXTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($FEXTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $FEXLoop)
	}

#Writing Blade Server information to Excel
if ($BladeLoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing Blade Server information to Excel...(Please Wait)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'Blade Server(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($BladeLoop -eq 1)
					{
						$DNFull = $BladeTable.DN -match "/(?<content>.*)"
						$DN = $matches['content']
						$SPFull = $BladeTable.ServiceProfile -match "/ls-(?<content2>.*)"
						$SP = $matches['content2']
						if ($SP -eq $null)
							{
								$SP = "<<UNASSOCIATED>>"
							}

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$BladeTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$SP
						$Worksheet.Cells.item($Loop+1,5) = "'"+$BladeTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$BladeTable.Serial
						if (($BladeTable.MfgTime -eq "") -or ($BladeTable.MfgTime -eq $null) -or ($BladeTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($BladeTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $BladeTable.DN[$DoLoop-1] -match "/(?<content>.*)"
						$DN = $matches['content']
						$SPFull = $BladeTable.ServiceProfile[$DoLoop-1] -match "/ls-(?<content2>.*)"
						$SP = $matches['content2']
						if ($SP -eq $null)
							{
								$SP = "<<UNASSOCIATED>>"
							}
						
						$Worksheet.Cells.Item($Loop+1,2) = "'"+$BladeTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$SP
						$Worksheet.Cells.item($Loop+1,5) = "'"+$BladeTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$BladeTable.Serial[$DoLoop-1]
						if ($BladeTable.MfgTime.Count -ne 0)
							{
								if (($BladeTable.MfgTime[$DoLoop-1] -eq "") -or ($BladeTable.MfgTime[$DoLoop-1] -eq $null) -or ($BladeTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($BladeTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $BladeLoop)
	}

#Writing Rack Server information to Excel
if ($RackLoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing Rack Server information to Excel...(Please Wait)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'Rack Server(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($RackLoop -eq 1)
					{
						$DNFull = $RackTable.DN -match "/(?<content>.*)"
						$DN = $matches['content']
						$SPFull = $RackTable.ServiceProfile -match "/ls-(?<content2>.*)"
						$SP = $matches['content2']
						if ($SP -eq $null)
							{
								$SP = "<<UNASSOCIATED>>"
							}

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$RackTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$SP
						$Worksheet.Cells.item($Loop+1,5) = "'"+$RackTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$RackTable.Serial
						if (($RackTable.MfgTime -eq "") -or ($RackTable.MfgTime -eq $null) -or ($RackTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($RackTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $RackTable.DN[$DoLoop-1] -match "/(?<content>.*)"
						$DN = $matches['content']
						$SPFull = $RackTable.ServiceProfile[$DoLoop-1] -match "/ls-(?<content2>.*)"
						$SP = $matches['content2']
						if ($SP -eq $null)
							{
								$SP = "<<UNASSOCIATED>>"
							}
						
						$Worksheet.Cells.Item($Loop+1,2) = "'"+$RackTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$SP
						$Worksheet.Cells.item($Loop+1,5) = "'"+$RackTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$RackTable.Serial[$DoLoop-1]
						if ($RackTable.MfgTime.Count -ne 0)
							{
								if (($RackTable.MfgTime[$DoLoop-1] -eq "") -or ($RackTable.MfgTime[$DoLoop-1] -eq $null) -or ($RackTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($RackTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $RackLoop)
	}

#Writing Adapter information to Excel
if ($AdapterLoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing Adapter information to Excel...(Please Wait)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'Adapter(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($AdapterLoop -eq 1)
					{
						$DNFull = $AdapterTable.DN -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$AdapterTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$AdapterTable.ServiceProfile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$AdapterTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$AdapterTable.Serial
						if (($AdapterTable.MfgTime -eq "") -or ($AdapterTable.MfgTime -eq $null) -or ($AdapterTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($AdapterTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $AdapterTable.DN[$DoLoop-1] -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$AdapterTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$AdapterTable.ServiceProfile[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,5) = "'"+$AdapterTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$AdapterTable.Serial[$DoLoop-1]
						if ($AdapterTable.MfgTime.Count -ne 0)
							{
								if (($AdapterTable.MfgTime[$DoLoop-1] -eq "") -or ($AdapterTable.MfgTime[$DoLoop-1] -eq $null) -or ($AdapterTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($AdapterTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $AdapterLoop)
	}

#Writing Adapter Expansion information to Excel
if ($AdapterExpansionLoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing Adapter Expansion information to Excel...(Please Wait)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'Adapter Expansion(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($AdapterLoop -eq 1)
					{
						$DNFull = $AdapterExpansionTable.DN -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$AdapterExpansionTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$AdapterExpansionTable.ServiceProfile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$AdapterExpansionTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$AdapterExpansionTable.Serial
						if (($AdapterExpansionTable.MfgTime -eq "") -or ($AdapterExpansionTable.MfgTime -eq $null) -or ($AdapterExpansionTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($AdapterExpansionTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $AdapterExpansionTable.DN[$DoLoop-1] -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$AdapterExpansionTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$AdapterExpansionTable.ServiceProfile[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,5) = "'"+$AdapterExpansionTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$AdapterExpansionTable.Serial[$DoLoop-1]
						if ($AdapterExpansionTable.MfgTime.Count -ne 0)
							{
								if (($AdapterExpansionTable.MfgTime[$DoLoop-1] -eq "") -or ($AdapterExpansionTable.MfgTime[$DoLoop-1] -eq $null) -or ($AdapterExpansionTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($AdapterExpansionTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $AdapterExpansionLoop)
	}

#Writing DIMM information to Excel
if ($DIMMsLoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing DIMM information to Excel...(Please Wait...a long time :-)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'DIMM(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($DIMMsLoop -eq 1)
					{
						$DNFull = $DIMMsTable.DN -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$DIMMsTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$DIMMsTable.ServiceProfile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$DIMMsTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$DIMMsTable.Serial
						if (($DIMMsTable.MfgTime -eq "") -or ($DIMMsTable.MfgTime -eq $null) -or ($DIMMsTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($DIMMsTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $DIMMsTable.DN[$DoLoop-1] -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$DIMMsTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$DIMMsTable.ServiceProfile[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,5) = "'"+$DIMMsTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$DIMMsTable.Serial[$DoLoop-1]
						if ($DIMMsTable.MfgTime.Count -ne 0)
							{
								if (($DIMMsTable.MfgTime[$DoLoop-1] -eq "") -or ($DIMMsTable.MfgTime[$DoLoop-1] -eq $null) -or ($DIMMsTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($DIMMsTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $DIMMsLoop)
	}

#Writing Storage Controller information to Excel
if ($StorageControllerLoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing Storage Controller information to Excel...(Please Wait)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'Storage Controller(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($StorageControllerLoop -eq 1)
					{
						$DNFull = $StorageControllerTable.DN -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$StorageControllerTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$StorageControllerTable.ServiceProfile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$StorageControllerTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$StorageControllerTable.Serial
						if (($StorageControllerTable.MfgTime -eq "") -or ($StorageControllerTable.MfgTime -eq $null) -or ($StorageControllerTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($StorageControllerTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $StorageControllerTable.DN[$DoLoop-1] -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$StorageControllerTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$StorageControllerTable.ServiceProfile[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,5) = "'"+$StorageControllerTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$StorageControllerTable.Serial[$DoLoop-1]
						if ($StorageControllerTable.MfgTime.Count -ne 0)
							{
								if (($StorageControllerTable.MfgTime[$DoLoop-1] -eq "") -or ($StorageControllerTable.MfgTime[$DoLoop-1] -eq $null) -or ($StorageControllerTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($StorageControllerTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $StorageControllerLoop)
	}

#Writing Local Disk information to Excel
if ($StorageLoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing Local Disk information to Excel...(Please Wait)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'Local Disk(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($StorageLoop -eq 1)
					{
						$DNFull = $StorageTable.DN -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$StorageTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$StorageTable.ServiceProfile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$StorageTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$StorageTable.Serial
						if (($StorageTable.MfgTime -eq "") -or ($StorageTable.MfgTime -eq $null) -or ($StorageTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($StorageTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $StorageTable.DN[$DoLoop-1] -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$StorageTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$StorageTable.ServiceProfile[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,5) = "'"+$StorageTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$StorageTable.Serial[$DoLoop-1]
						if ($StorageTable.MfgTime.Count -ne 0)
							{
								if (($StorageTable.MfgTime[$DoLoop-1] -eq "") -or ($StorageTable.MfgTime[$DoLoop-1] -eq $null) -or ($StorageTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($StorageTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $StorageLoop)
	}

#Writing Fan information to Excel
if ($FanLoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing Fan information to Excel...(Please Wait)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'Fan(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($FanLoop -eq 1)
					{
						$DNFull = $FanTable.DN -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$FanTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'N/A" #Service Profile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$FanTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$FanTable.Serial
						if (($FanTable.MfgTime -eq "") -or ($FanTable.MfgTime -eq $null) -or ($FanTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($FanTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $FanTable.DN[$DoLoop-1] -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$FanTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'N/A" #Service Profile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$FanTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$FanTable.Serial[$DoLoop-1]
						if ($FanTable.MfgTime.Count -ne 0)
							{
								if (($FanTable.MfgTime[$DoLoop-1] -eq "") -or ($FanTable.MfgTime[$DoLoop-1] -eq $null) -or ($FanTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($FanTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $FanLoop)
	}

#Writing PSU information to Excel
if ($PSULoop -ne 0)
	{
		Write-Output ""
		Write-Output ""
		Write-Output "Writing Power Supply information to Excel...(Please Wait)"
		$DoLoop = 1
		$Loop += 1
		$Worksheet.Cells.item($Loop+1,1) = "'PSU(s)"
		do
			{
				Write-Host "." -NoNewline
				if ($PSULoop -eq 1)
					{
						$DNFull = $PSUTable.DN -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$PSUTable.UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$PSUTable.ServiceProfile
						$Worksheet.Cells.item($Loop+1,5) = "'"+$PSUTable.Model
						$Worksheet.Cells.item($Loop+1,6) = "'"+$PSUTable.Serial
						if (($PSUTable.MfgTime -eq "") -or ($PSUTable.MfgTime -eq $null) -or ($PSUTable.MfgTime -eq "not-applicable"))
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+($PSUTable.MfgTime -replace "T", "  ")
							}
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$DNFull = $PSUTable.DN[$DoLoop-1] -match "/(?<content>.*)"
						$DN = $matches['content']

						$Worksheet.Cells.Item($Loop+1,2) = "'"+$PSUTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$DN
						$Worksheet.Cells.item($Loop+1,4) = "'"+$PSUTable.ServiceProfile[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,5) = "'"+$PSUTable.Model[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$PSUTable.Serial[$DoLoop-1]
						if ($PSUTable.MfgTime.Count -ne 0)
							{
								if (($PSUTable.MfgTime[$DoLoop-1] -eq "") -or ($PSUTable.MfgTime[$DoLoop-1] -eq $null) -or ($PSUTable.MfgTime[$DoLoop-1] -eq "not-applicable"))
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
									}
								else
									{
										$Worksheet.Cells.item($Loop+1,7) = "'"+($PSUTable.MfgTime[$DoLoop-1] -replace "T", "  ")
									}
							}
						else
							{
								$Worksheet.Cells.item($Loop+1,7) = "'"+"N/A"
							}
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $PSULoop)
	}

#Licensing Information
$Loop = $Loop + 2

#Create Excel headers
Write-Output ""
Write-Output ""
Write-Output "Setting worksheet headers for Licenses"
$Worksheet.Cells.Item($Loop,1) = "Category"
$Worksheet.Cells.Item($Loop,2) = "UCS"
$Worksheet.Cells.Item($Loop,3) = "Fabric Interconnecct"
$Worksheet.Cells.item($Loop,4) = "Absolute Quantity"
$Worksheet.Cells.item($Loop,5) = "Default Quantity"
$Worksheet.Cells.item($Loop,6) = "Used Quantity"
$Worksheet.Cells.item($Loop,7) = "Grace Period Used"
$Worksheet.Cells.item($Loop,8) = "Operational State"
$Worksheet.Cells.item($Loop,9) = "Peer Status"

#Format Excel cell headers
Write-Output "	Formatting cells"
$Worksheet.Cells.item($Loop,1).font.size=12
$Worksheet.Cells.item($Loop,1).font.bold=$true
$Worksheet.Cells.item($Loop,1).font.underline=$true
$Worksheet.Cells.item($Loop,2).font.size=12
$Worksheet.Cells.item($Loop,2).font.bold=$true
$Worksheet.Cells.item($Loop,2).font.underline=$true
$Worksheet.Cells.item($Loop,3).font.size=12
$Worksheet.Cells.item($Loop,3).font.bold=$true
$Worksheet.Cells.item($Loop,3).font.underline=$true
$Worksheet.Cells.item($Loop,4).font.size=12
$Worksheet.Cells.item($Loop,4).font.bold=$true
$Worksheet.Cells.item($Loop,4).font.underline=$true
$Worksheet.Cells.item($Loop,5).font.size=12
$Worksheet.Cells.item($Loop,5).font.bold=$true
$Worksheet.Cells.item($Loop,5).font.underline=$true
$Worksheet.Cells.item($Loop,6).font.size=12
$Worksheet.Cells.item($Loop,6).font.bold=$true
$Worksheet.Cells.item($Loop,6).font.underline=$true
$Worksheet.Cells.item($Loop,7).font.size=12
$Worksheet.Cells.item($Loop,7).font.bold=$true
$Worksheet.Cells.item($Loop,7).font.underline=$true
$Worksheet.Cells.item($Loop,8).font.size=12
$Worksheet.Cells.item($Loop,8).font.bold=$true
$Worksheet.Cells.item($Loop,8).font.underline=$true
$Worksheet.Cells.item($Loop,9).font.size=12
$Worksheet.Cells.item($Loop,9).font.bold=$true
$Worksheet.Cells.item($Loop,9).font.underline=$true
$Worksheet.columns.item(8).columnWidth = 25
$Worksheet.columns.item(9).columnWidth = 25

#Writing Fabric Interconnect Licensing information to Excel
if ($LicenseLoop -ne 0)
	{
		Write-Output "		Writing Fabric Interconnect Licensing information to Excel...(Please Wait)"
		$DoLoop = 1
		$Worksheet.Cells.item($Loop+1,1) = "'Fabric Interconnect(s) Licensing"
		do
			{
				Write-Host "." -NoNewline
				if ($LicenseLoop -eq 1)
					{
						$Worksheet.Cells.Item($Loop+1,2) = "'"+$LicenseTable.UCS #UCS
						$Worksheet.Cells.item($Loop+1,3) = "'"+$LicenseTable.FI #FI
						$Worksheet.Cells.item($Loop+1,4) = "'"+$licenseTable.Abs #Absolute License Quantity
						$Worksheet.Cells.item($Loop+1,5) = "'"+$LicenseTable.Def #Default License Quantity
						$Worksheet.Cells.item($Loop+1,6) = "'"+$LicenseTable.Used #Used License Quantity
						$Worksheet.Cells.item($Loop+1,7) = "'"+$LicenseTable.Grace #Grace Period Used
						$Worksheet.Cells.item($Loop+1,8) = "'"+$LicenseTable.Oper #Operational State
						$Worksheet.Cells.item($Loop+1,9) = "'"+$LicenseTable.Peer #Peer State
						$Loop += 1
						$DoLoop += 1
					}
				else
					{
						$Worksheet.Cells.Item($Loop+1,2) = "'"+$LicenseTable.UCS[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,3) = "'"+$LicenseTable.FI[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,4) = "'"+$licenseTable.Abs[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,5) = "'"+$LicenseTable.Def[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,6) = "'"+$LicenseTable.Used[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,7) = "'"+$LicenseTable.Grace[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,8) = "'"+$LicenseTable.Peer[$DoLoop-1]
						$Worksheet.Cells.item($Loop+1,9) = "'"+$LicenseTable.Oper[$DoLoop-1]
						$Loop += 1
						$DoLoop += 1
					}
			}
		while ($DoLoop -le $LicenseLoop)
	}

#Save the Excel file
Write-Output ""
Write-Output ""
[string]$ConnectedUCS = ($FIs | where {$_.Id -eq "A"}).ucs
if ($ConnectedUCS.Length -ge 232)
	{
		$UCSList = $ConnectedUCS.substring(0,220)
	}
else
	{
		$UCSList = $ConnectedUCS
	}
$date = Get-Date
$DateFormat = [string]$date.Month+"-"+[string]$Date.Day+"-"+[string]$date.year+"_"+[string]$date.Hour+"-"+[string]$date.Minute+"-"+[string]$date.Second
$file = $PSScriptRoot + "\UCS Serial Numbers for "+$UCSList+"_"+$DateFormat+".xlsx"
Write-Output "The Excel file will be created as:"
Write-Output $file
Write-Output "	Saving Excel File...(Please wait)"
$Workbook.SaveAs($file)
Write-Output "		Complete"

#Close the Excel file
Write-Output ""
Write-Output "Closing Excel Spreadsheet..."
$Workbook.Close()

#Exit Excel
Write-Output ""
Write-Output "Exiting Excel..."
$Excel.Quit()

#Exit the Script
Write-Output ""
Write-Output "Script Complete"
Disconnect-Ucs
exit