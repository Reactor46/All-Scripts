<#

.SYNOPSIS
	This script add VLANs to a single or multiple UCS domains from a CSV

.DESCRIPTION
	This script add VLANs to a single or multiple UCS domains from a CSV.  The CSV file must be formatted as follows:
		Name,VLAN
		MGMT,100
		Data,101

.EXAMPLE
	Add-UcsVlans.ps1
	This script can be run without any command line parameters.  User will be prompted for all parameters and options required

.EXAMPLE
	Add-UcsVlans.ps1 -ucs "1.2.3.4" -file "List_of_VLANs.csv" -ucred
	-ucs -- UCS Manager IP address(s) or Host Name(s) -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-file -- CSV file to import VLAN names and VLAN numbers
	-ucred -- UCS Manager Credential Switch -- Adding this switch will immediately prompt you for your UCSM username and password
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords

.EXAMPLE
	Add-UcsVlans.ps1 -ucs "1.2.3.4" -saved "myucscred.csv" -file "List_of_VLANs.csv" -skiperror
	-ucs -- UCS Manager IP address(s) or Host Name(s) -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-savedcred -- UCSM credentials file -- Example: -savedcred "myucscred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\myucscred.csv
		Make sure the password file is located in the same folder as the script
	-file -- CSV file to import VLAN names and VLAN numbers
	-skiperrors -- Tells the script to skip any prompts for errors and continues with 'y'
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords

.NOTES
	Author: Joe Martin
	Email: joemar@cisco.com
	Company: Cisco Systems, Inc.
	Version: v0.1.03
	Date: 11/13/2015
	Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
	***Provide any additional inputs here
	UCSM IP Address(s) or Hostname(s)
	UCSM Credentials Filename
	UCSM Username and Password
	CSV File

.OUTPUTS
	None
	
.LINK
	http://communities.cisco.com/people/joemar/content

#>

#Command Line Parameters
param(
	[string]$UCSM,				# IP Address(s) or Hostname(s).  If multiple entries, separate by commas
	[switch]$UCREDENTIALS,		# UCSM Credentials (Username and Password).  Requires all domains to use the same credentials
	[string]$SAVEDCRED,			# Saved UCSM Credentials.  To create do: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} | Export-CSV -NoTypeInformation .\myucscred.csv
	[string]$FILE,				# VLANs CSV file.  Must be located in the same folder as this script
	[switch]$SKIPERROR			# Skip any prompts for errors and continues with 'y'
)

#Clear the screen
clear-host

#Script kicking off
Write-Output "Script Running..."
Write-Output ""

#Tell the user what the script does
Write-Output "This script will import a list of VLAN names and VLAN number from a CSV and add those VLANs to a single or multiple UCS domains"
Write-Output "The CSV file must be formatted as:"
Write-Output "		Name,VLAN"
Write-Output "		MGMT,100"
Write-Output "		Data,101"
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

#Get Data File
cd $PSScriptRoot
$CsvFiles = dir "*.csv"

#Load Data File
if ($FILE)
	{
		if (Test-Path $FILE)
			{
				#Set the data configuration file
				$CsvFile = $FILE
			}
		else
			{
				Write-Host ""
				Write-Host -ForegroundColor Red "The file you specified does not exist"
				Write-Host -ForegroundColor Red "	Exiting..."
				Disconnect-Ucs
				exit
			}
	}
else
	{
		#Create list of CSV Files
		[array]$DropDownArray = $null
		foreach ($CFs in $CsvFiles)
			{
				[array]$DropDownArray += $CFs.Name
			}
		
		#Menu Function
		function Return-DropDown 
			{
				$Choice = $DropDown.SelectedItem.ToString()
				$Form.Close()
			}
		[array]$DropDownArray += "EXIT"

		#Generate GUI input box for CSV File Selection
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
		$Form = New-Object System.Windows.Forms.Form
		$Form.width = 700
		$Form.height = 150
		$Form.StartPosition = "CenterScreen"
		$Form.Text = ”VLAN CSV File to use”
		$DropDown = new-object System.Windows.Forms.ComboBox
		$DropDown.Location = new-object System.Drawing.Size(100,10)
		$DropDown.Size = new-object System.Drawing.Size(550,30)
		ForEach ($Item in $DropDownArray) 
			{
				$DropDown.Items.Add($Item) | Out-Null
			}
		$Form.Controls.Add($DropDown)
		$DropDownLabel = new-object System.Windows.Forms.Label
		$DropDownLabel.Location = new-object System.Drawing.Size(1,10)
		$DropDownLabel.size = new-object System.Drawing.Size(255,20)
		$DropDownLabel.Text = "VLAN CSV File"
		$Form.Controls.Add($DropDownLabel)
		$Button = new-object System.Windows.Forms.Button
		$Button.Location = new-object System.Drawing.Size(300,50)
		$Button.Size = new-object System.Drawing.Size(75,25)
		$Button.Text = "Select"
		$Button.Add_Click({Return-DropDown})
		$form.Controls.Add($Button)
		$Form.Add_Shown({$Form.Activate()})
		$Form.ShowDialog() | Out-Null
		
		#Check for valid entry
		if ($DropDown.SelectedItem -eq $null)
			{
				Write-Host -ForegroundColor Red "Nothing Selected"
				Disconnect-Ucs
				exit
			}
		
		#Check to see if EXIT selected
		if ($DropDown.SelectedItem -eq "EXIT")
			{
				Write-Host -ForegroundColor DarkBlue "You have chosen to EXIT the script"
				Disconnect-Ucs
				exit
			}
		#Set the data configuration file
		$CsvFile = $DropDown.SelectedItem
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

$VLANList = Import-Csv $CsvFile
if (!$VLANList)
	{
		Write-Output "The file import failed"
		Write-Output "	Exiting..."
		exit
	}
else
	{
		Write-Output ""
		$VLANCloud = Get-UcsLanCloud
			if (!$VLANCloud)
				{
					Write-Output "Could not access VLAN Cloud"
					Write-Output "	Exiting..."
					exit
				}
			else
				{
					Write-Output "Validating VLAN Names and Numbers"
					foreach ($Item in $VLANList)
						{
							$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
							if ((($Item.Name).Length -le 32) -and (($CharacterTest.Match($Item.Name).Success)))
								{
								}
							else
								{
									Write-Host -ForegroundColor Red 'Invalid entry in VLAN Name:'$Item.Name
									Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
									Write-Host -ForegroundColor Red "		Exiting..."
									Disconnect-Ucs
									exit
								}
			
							if (([int]$Item.VLAN -le 3967) -or ($Item.VLAN -match "^404[8-9]") -or ($Item.VLAN -match "^40[5-8][0-9]") -or ($Item.VLAN -match "^409[0-3]"))
								{
								}
							else
								{
									Write-Host -ForegroundColor Red 'Invalid entry in VLAN Number:'$Item.VLAN
									Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
									Write-Host -ForegroundColor Red "		Exiting..."
									Disconnect-Ucs
									exit
								}
						}
					Write-Output "	Validated Successfully"
					Write-Output ""
					Write-Output "Checking VLAN Names and Numbers against existing UCS entries"
					$ExistingVlans = Get-UcsVlan -Cloud ethlan
					foreach ($Item in $ExistingVlans)
						{
							if ($VLANList.Name -contains $Item.Name)
								{
									Write-Host '	VLAN Name already exists:'$Item.Name
									Write-Host -ForegroundColor Red "	Please correct this issue in the data file or UCS and try again"
									Write-Host -ForegroundColor Red "		Exiting..."
									Disconnect-Ucs
									exit
								}							
							if ($VLANList.VLAN -contains $Item.Id)
								{
									Write-Host '	VLAN Number already exists:'$Item.Id
									Write-Host -ForegroundColor Red "	Please correct this issue in the data file or UCS and try again"
									Write-Host -ForegroundColor Red "		Exiting..."
									Disconnect-Ucs
									exit
								}							
						}
					Write-Output "	No overlaps exist"
					Write-Output ""
					$ErrorActionPreference = "Continue"
					Write-Output "Adding VLANs from CSV File, please wait..."
					Start-UcsTransaction
						foreach ($Item in $VLANList)
							{
								$VLANCloud | Add-UcsVlan -Id $Item.VLAN -Name $Item.Name
							}
					Complete-UcsTransaction
					Write-Output "Done adding VLANs"
				}
	}

#Disconnect from UCSM(s)
Disconnect-Ucs

#Exit the Script
Write-Output ""
Write-Output "Script Complete"
exit