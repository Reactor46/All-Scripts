<#

.SYNOPSIS
	Access UCS KVM across multiple UCSM domains from a single interface

.DESCRIPTION
	This script allows you to connect to multiple UCSM domains anywhere in the world and select the Service Profile you wish to KVM to.

.EXAMPLE
	Invoke-UcsKvm.ps1
	This script can be run without any command line parameters.  User will be prompted for all parameters and options required

.EXAMPLE
	Invoke-UcsKvm.ps1 -ucs "1.2.3.4,5,6,7,8" -ucred
	-ucs -- UCS Manager IP or Host Name -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-ucred -- UCS Manager Credential Switch -- Adding this switch will immediately prompt you for your UCSM username and password
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords
	
.EXAMPLE
	Invoke-UcsKvm.ps1 -ucs "1.2.3.4,5,6,7,8" -saved "myucscred.csv" -serviceprofile "test1" -skiperrors
	-ucs -- UCS Manager IP or Host Name -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-savedcred -- UCSM credentials file -- Example: -savedcred "myucscred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\myucscred.csv
		Make sure the password file is located in the same folder as the script
	-serviceprofile -- pick a specific service profile to launch the KVM for
	-skiperrors -- Tells the script to skip any prompts for errors and continues with 'y'
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords

.NOTES
	Author: Joe Martin
	Email: joemar@cisco.com
	Company: Cisco Systems, Inc.
	Version: v0.1.11
	Date: 7/11/2014
	Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
	UCSM IP Address(s) or Hostname(s)
	UCSM Username and Password
	UCSM Credentials Filename
	Select Service Profiles to KVM to

.OUTPUTS
	UCSM KVM launches to selected Service Profile
	
.LINK
	http://communities.cisco.com/people/joemar/content

#>

#Command Line Parameters
param(
	[string]$UCSM,				# IP Address(s) or Hostname(s).  If multiple entries, separate by commas
	[switch]$UCREDENTIALS,		# UCSM Credentials (Username and Password)
	[string]$SAVEDCRED,			# Saved UCSM Credentials.  To create do: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} | Export-CSV -NoTypeInformation .\myucscred.csv
	[string]$SERVICEPROFILE,		# Service Profile to KVM to
	[switch]$SKIPERROR			# Skip any prompts for errors and continues with 'y'
)

#Clear the screen
clear-host

#Script kicking off
Write-Output "Script Running..."
Write-Output ""

#Gather any credentials requested from command line
if ($UCREDENTIALS)
	{
		Write-Output ""
		Write-Output "Enter UCSM Credentials"
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
		Write-Output "If you are running a version of Windows before 7 or Server 2008R2 you need to update to be supported"
		Write-Output "		Exiting..."
		Disconnect-Ucs
		exit
	}

#Tell the user what the script does
Write-Output "This script allows you to log into multiple UCS Domains and access any service profile"
Write-Output "for KVM access"
Write-Output ""
Write-Output "Prerequisites:"
Write-Output "	PowerShell V3 or above"
Write-Output "	PowerTool for PowerShell"
Write-Output "	Java"
Write-Output "	Valid UCSM login that allows KVM access"
Write-Output "	Network access to UCSM(s)"
Write-Output ""

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
				Write-Output "		PowerTool is now Loaded"
			}
	}
else
	{
		Write-Output "	PowerTool is already Loaded"
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
if ($myCon.count -eq 0)
	{
		Write-Output ""
		Write-Output "You are not logged into any UCSM systems"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Wait for a few seconds
Write-Output ""
Write-Output "Gathering UCSM Information...Please wait..."
sleep 3

##Show all Service Profiles on all UCMSs
if (!$SERVICEPROFILE)
	{
		Clear-Host
		Write-Output "UCS Service Profiles"
		$serviceprofiles = Get-UcsServiceProfile | Sort-Object name | where {$_.Uuid -ne "derived"}
		$serviceprofiles | sort-object name | select name, usrlbl, ucs, assocstate, operstate | Format-Table -AutoSize | Out-Host
		Write-Output "Select the UCS Service Profile from the pulldown or select EXIT to quit"
	}
	
##Prompt for Service Profile to log into
foreach ($sp in $serviceprofiles)
	{
		[array]$DropDownArray += $sp.Name+" == "+$sp.Ucs
	}
[array]$DropDownArray += "EXIT"

##This Function Returns the Selected Value (Service Profiles or EXIT) and Closes the Form
function Return-DropDown 
	{
		$script:Choice = ($DropDown.SelectedItem.ToString()).split(" ==")[0]
		$script:FrameTitle = ($DropDown.SelectedItem.ToString()).split(" ==")[0] + "/" + ($DropDown.SelectedItem.ToString()).split(" ==")[4]
		$Form.Close()
		if ($script:Choice -ne "EXIT")
			{
				Get-UcsServiceProfile –Name $script:Choice -LimitScope | Start-UcsKvmSession -FrameTitle $script:FrameTitle
				sleep -Seconds 1
				Disconnect-Ucs
			}
	}
	
##GUI to list service profiles and select option
if (!$SERVICEPROFILE)
	{
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
		$Form = New-Object System.Windows.Forms.Form
		$Form.width = 400
		$Form.height = 150
		$Form.StartPosition = "CenterScreen"
		$Form.Text = ”Cisco UCS KVM Launcher”
		$DropDown = new-object System.Windows.Forms.ComboBox
		$DropDown.Location = new-object System.Drawing.Size(100,10)
		$DropDown.Size = new-object System.Drawing.Size(250,30)
		ForEach ($Item in $DropDownArray) 
			{
				$DropDown.Items.Add($Item) | Out-Null
			}
		$Form.Controls.Add($DropDown)
		$DropDownLabel = new-object System.Windows.Forms.Label
		$DropDownLabel.Location = new-object System.Drawing.Size(1,10)
		$DropDownLabel.size = new-object System.Drawing.Size(255,20)
		$DropDownLabel.Text = "Service Profiles"
		$Form.Controls.Add($DropDownLabel)
		$Button = new-object System.Windows.Forms.Button
		$Button.Location = new-object System.Drawing.Size(100,50)
		$Button.Size = new-object System.Drawing.Size(255,25)
		$Button.Text = "Select"
		$Button.Add_Click({Return-DropDown})
		$form.Controls.Add($Button)
		$Form.Add_Shown({$Form.Activate()})
		$Form.ShowDialog() | Out-Null
	}
else
	{
		$valid = Get-UcsServiceProfile | where {$_.name -eq $SERVICEPROFILE}
		if (!$valid)
			{
				Write-Output ""
				Write-Output "This is not a valid Service Profile Name"
				Write-Output "	Exiting..."
				Write-Output ""
				Disconnect-Ucs
				exit
			}
		else
			{
				Write-Output ""
				Write-Output "Launching KVM for Service Profile: $SERVICEPROFILE"
				$Choice = $SERVICEPROFILE
				$FrameTitle = $SERVICEPROFILE + "/" + $valid.ucs
				if ($Choice -ne "EXIT")
					{
						Get-UcsServiceProfile –Name $Choice -LimitScope | Start-UcsKvmSession -FrameTitle $FrameTitle
						sleep -Seconds 10
						Disconnect-Ucs
					}

			}
	}