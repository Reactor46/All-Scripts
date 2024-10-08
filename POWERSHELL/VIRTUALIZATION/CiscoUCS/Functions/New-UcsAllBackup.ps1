﻿<#

.SYNOPSIS
	This script allows you to backup a single or multiple UCS domains.  It will create each type of backup available on UCS.

.DESCRIPTION
	This script allows you to backup a single or multiple UCS domains.  It will create each type of backup available on UCS.

.EXAMPLE
	New-UcsAllBackup.ps1
	This script can be run without any command line parameters.  User will be prompted for all parameters and options required

.EXAMPLE
	New-UcsAllBackup.ps1 -ucs "1.2.3.4" -ucred -folder "\\fileserver\fileshare\folder" -skiperrors
	-ucs -- UCS Manager IP address(s) or Host Name(s) -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-ucred -- UCS Manager Credential Switch -- Adding this switch will immediately prompt you for your UCSM username and password
	-folder -- Folder to save backup files -- You must have access to this folder for the files to be written
	-skiperrors -- Tells the script to skip any prompts for errors and continues with 'y'
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords

.EXAMPLE
	New-UcsAllBackup.ps1 -ucs "1.2.3.4" -saved "myucscred.csv" -folder "\\fileserver\fileshare\folder" -skiperrors
	-ucs -- UCS Manager IP address(s) or Host Name(s) -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-savedcred -- UCSM credentials file -- Example: -savedcred "myucscred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\myucscred.csv
		Make sure the password file is located in the same folder as the script
	-folder -- Folder to save backup files -- You must have access to this folder for the files to be written
	-skiperrors -- Tells the script to skip any prompts for errors and continues with 'y'
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords

.NOTES
	Author: Joe Martin
	Email: joemar@cisco.com
	Company: Cisco Systems, Inc.
	Version: v0.1.01
	Date: 12/2/2015
	Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
	UCSM IP Address(s) or Hostname(s)
	UCSM Username and Password
	UCSM Credentials Filename
	Backup files destination folder

.OUTPUTS
	backup xml files
	
.LINK
	http://communities.cisco.com/people/joemar/content

#>

#Command Line Parameters
param(
	[string]$UCSM,				# IP Address(s) or Hostname(s).  If multiple entries, separate by commas
	[switch]$UCREDENTIALS,		# UCSM Credentials (Username and Password).  Requires all domains to use the same credentials
	[string]$SAVEDCRED,			# Saved UCSM Credentials.  To create do: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} | Export-CSV -NoTypeInformation .\myucscred.csv
	[string]$FOLDER,			# Folder to save backup files to
	[switch]$SKIPERROR			# Skip any prompts for errors and continues with 'y'
)

#Clear the screen
clear-host

#Script kicking off
Write-Output "Script Running..."
Write-Output ""

#Tell the user what the script does
Write-Output "This script allows you to backup a single or multiple UCS domains."
Write-Output "It will create each type of backup available on UCS."
Write-Output ""
Write-Output "NOTE: UCS Platform Emulator cannot take a full-state backup"
Write-Output ""

#Gather credentials if command line flag is set
if ($UCREDENTIALS)
	{
		Write-Output "Enter UCSM Credentials"
		Write-Output ""
		$cred = Get-Credential -Message "Enter UCSM Credentials"
	}

#Change directory to the script root
cd $PSScriptRoot

#Select folder to save files to
if ($FOLDER)
	{
		$TestPath = Test-Path $FOLDER
		if ($TestPath -eq $true)
			{
				#Hold for future options
			}
		else
			{
				Write-Output ""
				Write-Output "The folder you specified either does not exist or you do not have access to it"
				Write-Output "	Exiting..."
				Disconnect-Ucs
				exit				
			}
	}
else
	{
		$FOLDER = $PSScriptRoot
	}
Write-Output "The files will be saved to folder: $FOLDER"
Write-Output ""

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
					}
				if ($Try -ieq "y")
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

#Function that removes existing backup configs, creates a new one and then removes it when complete
Function UCSBackup ($UCStoBackup, $UCSbackupType)
	{
		$UcsName = (Get-UcsTopSystem | where {$_.Address -eq $UCStoBackup}).Ucs
		Write-Output ""
		Write-Output "Removing any previous UCSM Backup configurations from UCS: $UcsName($UcsDomain)"
		$DontShow = Get-UcsMgmtBackup | Remove-UcsMgmtBackup -Ucs $UcsName -Force
		Write-Output "	Complete"
		Write-Output ""
		
		Write-Output "Backing up UCS: $UcsName($UCStoBackup), Backup Type: $UCSbackupType"
		$Date = Get-Date
		$DateFormat = [string]$Date.Month+"-"+[string]$Date.Day+"-"+[string]$Date.Year+"_"+[string]$Date.Hour+"-"+[string]$Date.Minute+"-"+[string]$Date.Second
		$BackupFile = $FOLDER+"\UCSMBackup_"+$UcsName+"_"+$UCSbackupType+"_"+$DateFormat+".xml"
		Try
			{
				$DontShow = Backup-Ucs -PreservePooledValues -Type $UCSbackupType -Ucs $UcsName -PathPattern $BackupFile -ErrorAction Stop
			}
		Catch
			{
				Write-Output "	***WARNING*** Error creating backup: $BackupFile"
				Write-Output "		NOTE: This is normal behavior for a full-state backup on a UCS Emulator"
				Write-Output ""
			}
		Finally
			{
				if (Test-Path $BackupFile -eq $true)
					{
						Write-Output "	Complete - $BackupFile"
						Write-Output ""
					}
			}

		Write-Output "Removing backup job from UCS: $UcsName($UCStoBackup)"
		$Hostname = ((Get-WmiObject -Class Win32_ComputerSystem).Name).ToLower()
		$dontshow = Start-UcsTransaction
			$mo = Get-UcsMgmtBackup -Hostname $Hostname | Remove-UcsMgmtBackup -Force
		$dontshow = Complete-UcsTransaction -Force
		Write-Output "	Complete"
	}

#Main part of script which calls the backup function for a UCS domain and a backup type
$UcsHandle = Get-UcsStatus | select -Property VirtualIpv4Address
foreach ($UcsDomain in $UcsHandle.VirtualIpv4Address)
	{
		UCSBackup $UcsDomain "config-all"
		UCSBackup $UcsDomain "config-logical"
		UCSBackup $UcsDomain "config-system"
		UCSBackup $UcsDomain "full-state" # Remember that full-state doesn't work on an emulator and it will be skip the file creation
	}

#Disconnect from UCSM(s)
Disconnect-Ucs

#Exit the Script
Write-Output ""
Write-Output "Script Complete"
exit