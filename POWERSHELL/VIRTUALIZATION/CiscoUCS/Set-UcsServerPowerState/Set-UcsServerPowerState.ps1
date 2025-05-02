<#

.SYNOPSIS
	This script allows you to power ON or OFF any of the servers in a single or multiple UCS domains.

.DESCRIPTION
	This script allows you to power ON or OFF any of the servers in a single or multiple UCS domains.  Be very careful as this could be a career limiting move!

.EXAMPLE
	Set-UcsServerPowerState.ps1
	This script can be run without any command line parameters.  User will be prompted for all parameters and options required

.EXAMPLE
	Set-UcsServerPowerState.ps1 -ucs "1.2.3.4" -ucred -serviceprofile "Test1" -state "on"
	-ucs -- UCS Manager IP or Host Name -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-ucred -- UCS Manager Credential Switch -- Adding this switch will immediately prompt you for your UCSM username and password
	-serviceprofiles -- A single or multiple service profiles.  If multiple, separate them by a comma.
	-state -- Power State -- Options: on, off
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords

.EXAMPLE
	Set-UcsServerPowerState.ps1 -ucs "1.2.3.4" -saved "myucscred.csv" -serviceprofile "Test1" -state "on" -skiperror
	-ucs -- UCS Manager IP address(s) or Host Name(s) -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-savedcred -- UCSM credentials file -- Example: -savedcred "myucscred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\myucscred.csv
		Make sure the password file is located in the same folder as the script
	-serviceprofiles -- A single or multiple service profiles.  If multiple, separate them by a comma.
	-state -- Power State -- Options: on, off
	-skiperrors -- Tells the script to skip any prompts for errors and continues with 'y'
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords

.NOTES
	Author: Joe Martin
	Email: joemar@cisco.com
	Company: Cisco Systems, Inc.
	Version: v0.6.00
	Date: 4/24/2014
	Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
	UCSM IP Address(s) or Hostname(s)
	UCSM Username and Password
	UCSM Credentials file
	Server Power State

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
	[string]$SERVICEPROFILES,	# Single or multiple service profiles.  If multiple separate by a comma
	[string]$STATE,				# Power State -- Options: on, off
	[switch]$SKIPERROR			# Skip any prompts for errors and continues with 'y'
)

#Clear the screen
clear-host

#Script kicking off
Write-Output "Script Running..."
Write-Output ""

#Tell the user what the script does
Write-Output "This script allows you to power ON or OFF some or all the blades"
Write-Output "in multiple UCS domains.  Be very careful as this could be a"
Write-Output "career limiting move!"
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
				Disconnect-Ucs
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

$SPs = Get-UcsServiceProfile | where {$_.Type -eq "instance"}

if (!$SERVICEPROFILES)
	{
		#Offer list of SPs to control power
		Write-Output ""
		Write-Output "Select the service profile(s) you wish to power ON or OFF"
		Write-Output "	Press CANCEL or hit Esc to exit the script"
		Write-Output "	You can select multiple entries by holding the control key to pick"
		Write-Output "	individuals or hold the Shift key to select a range, or any combination."
		#Multi-Select routine example provided at: http://technet.microsoft.com/en-us/library/ff730950.aspx
		$Script:SelectedObjects = @()
		$Script:Exit = 'n'
		
		[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
		
		$objForm = New-Object System.Windows.Forms.Form 
		$objForm.Text = "Service Profiles"
		$objForm.Size = New-Object System.Drawing.Size(300,600) 
		$objForm.StartPosition = "CenterScreen"
		
		$objForm.KeyPreview = $True

		$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    			{
        			foreach ($objItem in $objListbox.SelectedItems)
          		{$Script:SelectedObjects += $objItem}
				$objForm.Close()
			}
		})

		$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
		{$objForm.Close(); Write-Output "" ; Write-Output "You pressed Escape"; Write-Output "	exiting..."; Disconnect-Ucs; $Script:Exit = "y"}})

		$OKButton = New-Object System.Windows.Forms.Button
		$OKButton.Location = New-Object System.Drawing.Size(75,500)
		$OKButton.Size = New-Object System.Drawing.Size(75,23)
		$OKButton.Text = "OK"

		$OKButton.Add_Click(
   			{
        			foreach ($objItem in $objListbox.SelectedItems)
            		{$Script:SelectedObjects += $objItem}
        			$objForm.Close()
   			})

		$objForm.Controls.Add($OKButton)

		$CancelButton = New-Object System.Windows.Forms.Button
		$CancelButton.Location = New-Object System.Drawing.Size(150,500)
		$CancelButton.Size = New-Object System.Drawing.Size(75,23)
		$CancelButton.Text = "Cancel"
		$CancelButton.Add_Click({$objForm.Close(); Write-Output ""; Write-Output "You pressed Cancel"; Disconnect-Ucs; $Script:Exit = "y"})
		$objForm.Controls.Add($CancelButton)
	
		$objLabel = New-Object System.Windows.Forms.Label
		$objLabel.Location = New-Object System.Drawing.Size(10,20) 
		$objLabel.Size = New-Object System.Drawing.Size(280,20) 
		$objLabel.Text = "Select from below. SHIFT or CNTRL for multi-select:"
		$objForm.Controls.Add($objLabel) 
		
		$objListbox = New-Object System.Windows.Forms.Listbox 
		$objListbox.Location = New-Object System.Drawing.Size(10,40) 
		$objListbox.Size = New-Object System.Drawing.Size(260,20)
		$objListBox.Sorted = $True

		$objListbox.SelectionMode = "MultiExtended"

		[void] $objListbox.Items.Add("--ALL--")
		foreach ($ServiceProfile in $SPs)
			{
				[void] $objListbox.Items.Add($ServiceProfile.Name)
			}

		$objListbox.Height = 450
		$objForm.Controls.Add($objListbox) 
		$objForm.Topmost = $True
	
		$objForm.Add_Shown({$objForm.Activate()})
		[void] $objForm.ShowDialog()
	}
else
	{
		if ($SERVICEPROFILES -ieq "all")
			{
				$Script:SelectedObjects = "--ALL--"
			}
		else
			{
				[array]$SPArray = ($SERVICEPROFILES.split(",")).trim()
				$Script:SelectedObjects = $SPArray
			}
	}

if (($Script:SelectedObjects.Count -ne 1) -and ($Script:SelectedObjects -eq "--ALL--"))
	{
		Write-Output ""
		Write-Output "ERROR.  If selecting ALL, can only select ALL"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}
else
	{
		if ($Script:SelectedObjects -eq "--ALL--")
			{

				foreach ($Item in $SPs)
					{
						[array]$SPList = $SPList + $Item.Name
					}
				[array]$Script:SelectedObjects = $SPList
			}
	}
if ($Script:Exit -eq "y")
	{
		Write-Output ""
		Write-Output "You have chosen to Exit"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

if ($Script:SelectedObjects.Count -eq 0)
	{
		Write-Output ""
		Write-Output "You didn't select anything"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Validate Service Profiles are valid
foreach ($Item in $Script:SelectedObjects)
	{
		$SpExists = Get-UcsServiceProfile -Name $Item
		if (!$SpExists)
			{
				Write-Output ""
				Write-Output "Service Profile $SpExists is not valid"
				Write-Output "	Exiting..."
				Disconnect-Ucs
				exit
			}
	}

$SpList = $Script:SelectedObjects

if (!$STATE)
	{
		#Turn servers on, off-graceful
		Write-Output ""
		Write-Output "Select the power state for the servers"
		Write-Output "	Press CANCEL or hit Esc to exit the script"
		#Multi-Select routine example provided at: http://technet.microsoft.com/en-us/library/ff730950.aspx
		$Script:SelectedObjects = @()
		$Script:Exit = 'n'
		
		[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
		
		$objForm = New-Object System.Windows.Forms.Form 
		$objForm.Text = "Power State"
		$objForm.Size = New-Object System.Drawing.Size(300,600) 
		$objForm.StartPosition = "CenterScreen"
		
		$objForm.KeyPreview = $True

		$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    			{
        			foreach ($objItem in $objListbox.SelectedItems)
          		{$Script:SelectedObjects += $objItem}
				$objForm.Close()
			}
		})

		$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
		{$objForm.Close(); Write-Output "" ; Write-Output "You pressed Escape"; Write-Output "	exiting..."; Disconnect-Ucs; $Script:Exit = "y"}})

		$OKButton = New-Object System.Windows.Forms.Button
		$OKButton.Location = New-Object System.Drawing.Size(75,500)
		$OKButton.Size = New-Object System.Drawing.Size(75,23)
		$OKButton.Text = "OK"

		$OKButton.Add_Click(
   			{
        			foreach ($objItem in $objListbox.SelectedItems)
            		{$Script:SelectedObjects += $objItem}
        			$objForm.Close()
   			})

		$objForm.Controls.Add($OKButton)

		$CancelButton = New-Object System.Windows.Forms.Button
		$CancelButton.Location = New-Object System.Drawing.Size(150,500)
		$CancelButton.Size = New-Object System.Drawing.Size(75,23)
		$CancelButton.Text = "Cancel"
		$CancelButton.Add_Click({$objForm.Close(); Write-Output ""; Write-Output "You pressed Cancel"; Disconnect-Ucs; $Script:Exit = "y"})
		$objForm.Controls.Add($CancelButton)
	
		$objLabel = New-Object System.Windows.Forms.Label
		$objLabel.Location = New-Object System.Drawing.Size(10,20) 
		$objLabel.Size = New-Object System.Drawing.Size(280,20) 
		$objLabel.Text = "Select from below. SHIFT or CNTRL for multi-select:"
		$objForm.Controls.Add($objLabel) 
		
		$objListbox = New-Object System.Windows.Forms.Listbox 
		$objListbox.Location = New-Object System.Drawing.Size(10,40) 
		$objListbox.Size = New-Object System.Drawing.Size(260,20)
		$objListBox.Sorted = $False

		#$objListbox.SelectionMode = "MultiExtended"

		[void] $objListbox.Items.Add("ON")
		[void] $objListbox.Items.Add("OFF")

		$objListbox.Height = 450
		$objForm.Controls.Add($objListbox) 
		$objForm.Topmost = $True
	
		$objForm.Add_Shown({$objForm.Activate()})
		[void] $objForm.ShowDialog()
	}
else
	{
		if ($STATE -ieq "on")
			{
				$Script:SelectedObjects = "on"
			}
		elseif ($STATE -ieq "off")
			{
				$Script:SelectedObjects = "off"
			}
		else
			{
				Write-Output ""
				Write-Output "You have selected an invalid power state"
				Write-Output "	Exiting..."
				Disconnect-Ucs
				exit

			}
				
	}
if ($Script:Exit -eq "y")
	{
		Write-Output ""
		Write-Output "You have chosen to Exit"
		Write-Output "	Exiting..."
		exit
	}

if ($Script:SelectedObjects.Count -eq 0)
	{
		Write-Output ""
		Write-Output "You didn't select anything"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

if ($Script:SelectedObjects -ieq "on")
	{
		$PowerValue = "admin-up"
	}
if ($Script:SelectedObjects -ieq "off")
	{
		$PowerValue = "soft-shut-down"
	}


#Set the power state on the servers
Write-Output ""
Write-Output "Setting Power States"
foreach ($Item in $SpList)
	{
		$Dontshow = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name $item -LimitScope | Get-UcsServerPower | Set-UcsServerPower -State $PowerValue -Force
		Write-Output "	Turning $Item $Script:SelectedObjects"
	}

#Disconnect from UCSM(s)
Disconnect-Ucs

#Exit the Script
Write-Output ""
Write-Output "Script Complete"
exit