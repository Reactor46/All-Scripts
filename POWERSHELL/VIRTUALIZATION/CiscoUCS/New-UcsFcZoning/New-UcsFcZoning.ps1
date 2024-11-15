﻿<#

.SYNOPSIS
	Generates SAN zoning information for UCS managed servers connected to Cisco or Brocade fabrics

.DESCRIPTION
	This script will take a CSV file of SAN Information and use information gathered from UCSM to generate Fibre Channel Zones and Zonesets for either Cisco or Brocade SAN Switches.  It allows you to save the configuration files to your drive or to write them directly to your equipment.

.EXAMPLE
	New-UcsFcZoning.ps1
	This script can be run without any command line parameters.  User will be prompted for all parameters and options required

.EXAMPLE
	New-UcsFcZoning.ps1 -req "y" -ucs "1.2.3.4" -ucred -serviceprofile "one, two, etc" -manufacture "Cisco" -wwpn "WWPN xxxx.csv" -output "Equipment" -fabrica "2.3.4.5" -acred -fabricb "3.4.5.6" -bcred
	-req -- Acknowledge that required prerequisites have been met -- Valid options are: Y or N
	-ucs -- UCS Manager IP or Host Name -- Example: 1.2.3.4 or myucs or myucs.domain.local
	-ucred -- UCS Manager Credential Switch -- Adding this switch will immediately prompt you for your UCSM username and password
	-serviceprofile -- Service Profile name or names to create zones for -- Valid options are: SingleName or NameOne,NameTwo,ETC or All
	-manufacture -- SAN fabric manufacture -- Valid options are: Cisco or Brocade
	-wwpn -- WWPN targets CSV file -- Example: WWPN xxxxxx.csv (File name must start with: WWPN)
	-output -- Destination of configuration -- Valid options are: File or Equipment
	-fabrica -- SAN Fabric A IP or Host Name -- Example: 2.3.4.5 or mysana or mysana.domain.local
	-acred -- SAN Fabric A Credential Switch -- Adding this switch will immediately prompt you for your SAN Fabric A username and password
	-fabricb -- SAN Fabric B IP or Host Name -- Example: 3.4.5.6 or mysanb or mysana.domain.local
	-bcred -- SAN Fabric B Credential Switch -- Adding this switch will immediately prompt you for your SAN Fabric B username and password
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords

.EXAMPLE
	New-UcsFcZoning.ps1 -req "y" -ucs "1.2.3.4" -ucred -serviceprofile "All" -manufacture "Brocade" -wwpn "WWPN xxxx.csv" -output "File"
	-req -- Acknowledge that required prerequisites have been met -- Valid options are: Y or N
	-ucs -- UCS Manager IP or Host Name -- Example: 1.2.3.4 or myucs or myucs.domain.local
	-ucred -- UCS Manager Credential Switch -- Adding this switch will immediately prompt you for your UCSM username and password
	-serviceprofile -- Service Profile name or names to create zones for -- Valid options are: SingleName or NameOne,NameTwo,ETC or All
	-manufacture -- SAN fabric manufacture -- Valid options are: Cisco or Brocade
	-wwpn -- WWPN targets CSV file -- Example: WWPN xxxxxx.csv (File name must start with: WWPN)
	-output -- Destination of configuration -- Valid options are: File or Equipment
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords
	
.EXAMPLE
	New-UcsFcZoning.ps1 -req "y" -ucs "1.2.3.4" -usaved "myucscred.csv" -serviceprofile "All" -manufacture "Cisco" -wwpn "WWPN xxxx.csv" -output "Equipment" -fabrica "2.3.4.5" -asaved "myacred.csv" -fabricb "3.4.5.6" -bsaved "mybcred.csv" -skiperrors
	-req -- Acknowledge that required prerequisites have been met -- Valid options are: Y or N
	-ucs -- UCS Manager IP or Host Name -- Example: 1.2.3.4 or myucs or myucs.domain.local
	-usavedcred -- UCSM credentials file -- Example: -usavedcred "myucscred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\myucscred.csv
		Make sure the password file is located in the same folder as the script
	-serviceprofile -- Service Profile name or names to create zones for -- Valid options are: SingleName or NameOne,NameTwo,ETC or All
	-manufacture -- SAN fabric manufacture -- Valid options are: Cisco or Brocade
	-wwpn -- WWPN targets CSV file -- Example: WWPN xxxxxx.csv (File name must start with: WWPN)
	-output -- Destination of configuration -- Valid options are: File or Equipment
	-fabrica -- SAN Fabric A IP or Host Name -- Example: 2.3.4.5 or mysana or mysana.domain.local
	-asavedcred -- Fabric A credentials file -- Example: -asavedcred "myacred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\myacred.csv
		Make sure the password file is located in the same folder as the script
	-fabricb -- SAN Fabric B IP or Host Name -- Example: 3.4.5.6 or mysanb or mysana.domain.local
	-bsavedcred -- Fabric B credentials file -- Example: -bsavedcred "mybcred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\mybcred.csv
		Make sure the password file is located in the same folder as the script	
	-skiperrors -- Tells the script to skip any prompts for errors and continues with 'y'
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords
	
.NOTES
	Author: Joe Martin
	Email: joemar@cisco.com
	Company: Cisco Systems, Inc.
	Version: v0.5.08
	Date: 2/12/2015
	Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
	UCSM IP Address or Hostname
	UCSM Username and Password
	UCSM Credentials File
	Select Service Profiles to Zone
	Select Cisco or Brocade for Zoning
	Select WWPN targets CSV file
	Select output to File or Equipment to SAN Switches
	If Equipment selected, Fabric A IP Address or Hostname
	If Equipment selected, Fabric A Username and Password
	If Equipment selected, Fabric B IP Address or Hostname
	If Equipment selected, Fabric B Username and Password

.OUTPUTS
	If File selected for output, two files will be created in the same directory as the script resides.
	File format are .TXT
	
.LINK
	http://communities.cisco.com/people/joemar/content

#>

#Command Line Parameters
param(
	[string]$REQUIREMENTSMET,	# Y or N
	[string]$UCSM,				# IP Address or Hostname
	[switch]$UCREDENTIALS,		# UCSM Credentials (Username and Password)
	[string]$USAVEDCRED,		# Saved UCSM Credentials.  To create do: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} | Export-CSV -NoTypeInformation .\myucscred.csv
	[string]$SERVICEPROFILES,	# ALL or Service Profile Name or list of Service Profile Names separated by commas
	[string]$MANUFACTURE,		# Cisco or Brocade
	[string]$WWPNCSV,			# WWPN xxxx.csv file located in the same directory as the script
	[string]$OUTPUT,			# File or Equipment
	[string]$FABRICA,			# IP Address or Hostname of SAN Fabric A
	[string]$ASAVEDCRED,		# Saved Fabric A Credentials.  To create do: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} | Export-CSV -NoTypeInformation .\myacred.csv
	[switch]$ACREDENTIALS,		# Fabric A Credentials (Username and Password)
	[string]$FABRICB,			# IP Address or Hostname of SAN Fabric B
	[string]$BSAVEDCRED,		# Saved Fabric B Credentials.  To create do: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} | Export-CSV -NoTypeInformation .\mybcred.csv
	[switch]$BCREDENTIALS,		# Fabric B Credentials (Username and Password)
	[switch]$SKIPERROR			# Skip any prompts for errors and continues with 'y'
)

#Clear the screen
clear-host

#Show user that script has started
Write-Output "Script Running..."

#Gather any credentials requested from command line
if ($UCREDENTIALS)
	{
		Write-Output ""
		Write-Output "Enter UCSM Credentials"
		$credu = Get-Credential -Message "Enter UCSM Credentials"
	}
if ($ACREDENTIALS)
	{
		Write-Output ""
		Write-Output "Enter Fabric A Credentials"
		$creda = Get-Credential -Message "Enter SAN Fabric A Credentials"
	}
if ($BCREDENTIALS)
	{
		Write-Output ""
		Write-Output "Enter Fabric B Credentials"
		$credb = Get-Credential -Message "Enter SAN Fabric B Credentials"
	}

#Change directory to the script root
cd $PSScriptRoot

#Check to see if credential files exists
if ($USAVEDCRED)
	{
		if ((Test-Path $USAVEDCRED) -eq $false)
			{
				Write-Output ""
				Write-Output "Your credentials file $USAVEDCRED does not exist in the script directory"
				Write-Output "	Exiting..."
				Disconnect-Ucs
				exit
			}
	}

if ($ASAVEDCRED)
	{
		if ((Test-Path $ASAVEDCRED) -eq $false)
			{
				Write-Output ""
				Write-Output "Your credentials file $ASAVEDCRED does not exist in the script directory"
				Write-Output "	Exiting..."
				Disconnect-Ucs
				exit
			}
	}

if ($BSAVEDCRED)
	{
		if ((Test-Path $BSAVEDCRED) -eq $false)
			{
				Write-Output ""
				Write-Output "Your credentials file $BSAVEDCRED does not exist in the script directory"
				Write-Output "	Exiting..."
				Disconnect-Ucs
				exit
			}
	}
	
#Tell user what the script does
Write-Output ""
Write-Output "Overview:"
Write-Output "	This script will generate the Aliases, Zones and Zone Sets for either"
Write-Output "	Brocade or Cisco MDS/Nexus."
Write-Output "		The script builds Single Initiator/Single Target Zones as this"
Write-Output "		is the standard recommended and supported by most array"
Write-Output "		manufactures."
Write-Output "	It uses a pre-defined CSV file of SAN target information along with"
Write-Output "	information gathered by UCSM."
Write-Output "	It allows you to save the configuration files to your drive or to write"
Write-Output "	directly to your equipment with SSH."
Write-Output "	It will generate the information for all Service Profiles with vHBAs or"
Write-Output "	selected ones."
Write-Output ""
Write-Output "Prerequisites for this script are:"
Write-Output "	PowerShell must be enabled on your client computer."
Write-Output "		Example: set-executionpolicy unrestricted -force"
Write-Output "	You must be running PowerShell version 3 or above."
Write-Output "	You must download and install Cisco PowerTool for PowerShell from:"
Write-Output "		http://www.cisco.com"
Write-Output "			A CCO Login is required."
Write-Output "	You must download plink.exe and place it in the same folder as this"
Write-Output "	script and CSV targets file."
Write-Output "		You can download plink.exe here:"
Write-Output "		http://www.chiark.greenend.org.uk/~sgtatham/putty/download.html"
Write-Output "	plink.exe license can be found here:"
Write-Output "		http://www.chiark.greenend.org.uk/~sgtatham/putty/licence.html"
Write-Output "		It is assumed plink will be located in this folder:"
Write-Output "			$PSScriptRoot"
Write-Output "				But you can change this in the script."
write-output '				 Search for: $PlinkAndPath'
Write-Output "				 and edit to match the location of your choice."
Write-Output "	You must be network connected and have reachability via SSH TCP Port 22"
Write-Output "	to your SAN Fabric Switches/Directors."
Write-Output "	You must be network connected and have reachability via SSL TCP Port"
Write-Output "	443 to your Cisco UCS Domain."
Write-Output "	You must have a login into your UCSM domain with appropriate rights."
Write-Output "	You must have a login into your SAN Fabric Switches/Directors with"
Write-Output "	appropriate rights."
Write-Output "	You must have this script and the required WWPN targets CSV file in the"
Write-Output "	same folder."
Write-Output "	If outputting configurations to a file they will also be saved in the"
Write-Output "	same folder as this script and WWPN targets CSV file."
Write-Output ""
Write-Output "The script supports Service Profiles with vHBAs for:"
Write-Output "	Two vHBAs: A & B Fabric."
Write-Output "	One vHBA: A Fabric."
Write-Output "	One vHBA: B Fabric."
Write-Output "		It does NOT support having some Service Profiles with only an A"
Write-Output "		fabric vHBA and some with only a B fabric vHBA or where some"
Write-Output "		have two vHBAs and others only have one."
Write-Output ""
Write-Output "Scalability and Performance:"
Write-Output "	The scripts tested limits are building Single Initiator/Single Target"
Write-Output "	Zones is 160 Service Profiles against a 32 Port SAN Array (16 A and 16"
Write-Output "	B controller ports).  This is 2560 Aliases, Zones and Zone Set members "
Write-Output "	per fabric."
Write-Output "		This is not a maximum, but a maximum tested."
Write-Output "	It takes only a few seconds to build configurations written to file."
Write-Output "	It took ~1 second to build each zone when pushing to an MDS switch and"
Write-Output "	~4 seconds to build each zone when pushing to a Brocade switch."
Write-Output "		For 2560 zones it took 21 minutes to the MDS and 114 minutes to"
Write-Output "		the brocade for each fabric so be patient when writing to your"
Write-Output "		equipment."
Write-Output ""
Write-Output "How it works:"
Write-Output "	The script logs into a UCSM Domain and collects information about the"
Write-Output "	VSANS in use, The Service Profiles and their attached vHBA WWPNs."
Write-Output "	The script then reads the WWPN CSV file which contains the name of"
Write-Output "	the zoneset to be used in the fabric along with the names and WWPNs of"
Write-Output "	the Array targets."
Write-Output "	The script then builds aliases, zones and zonesets based on the"
Write-Output "	commands for either Cisco or Brocade."
Write-Output "	Finally the script outputs the configurations to either file or will"
Write-Output "	SSH directly into your SAN fabric and writes the configuration to your"
Write-Output "	fabric."
Write-Output ""

#Have you met all the prerequisites and want to proceed
if (($REQUIREMENTSMET -ieq "y") -or ($REQUIREMENTSMET -ieq "n"))
	{
		$Choice = $REQUIREMENTSMET
	}
else
	{
		$Choice = Read-Host "Have you met the above prerequisites? (Y/N)"
	}
if ($Choice -ieq "y")
	{
		Write-Output ""
	}
elseif ($Choice -ieq "n")
	{
		Write-Output "You have chosen to exit"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}
else
	{
		Write-Output "You have selected an invalid option"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

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

#Set error action preference
#$ErrorActionPreference = "SilentlyContinue"
#$ErrorActionPreference = "Stop"
#$ErrorActionPreference = "Continue"
#$ErrorActionPreference = "Inquire"
$ErrorLevel = "SilentlyContinue"
$ErrorActionPreference = $ErrorLevel

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
				Write-Output "Cisco UCS PowerTool Module did not load.  Please correct his issue and try again"
				Write-Output "	Exiting..."
				Disconnect-Ucs
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

#Location of plink.exe application
#Make sure to place it in the following folder or update the path below
$PlinkAndPath = "$PSScriptRoot\plink.exe"

#Validate plink installed in identified folder
Write-Output ""
Write-Output "Validating that plink.exe is located in the correct folder"
if (Test-Path $PlinkAndPath)
	{
		Write-Output "	plink.exe is located in the specified folder:"
		Write-Output "		$PlinkAndPath"
	}
else
	{
		Write-Output "	plink.exe is MISSING.  Please download and place in the specified folder: $PlinkAndPath"
		Write-Output "	You can download from: You can download plink here: http://www.chiark.greenend.org.uk/~sgtatham/putty/download.html"
		Write-Output "		Exiting..."
		Disconnect-Ucs
		exit
	}

#Define UCS Domain(s)
Write-Output ""
Write-Output "Connecting to UCSM"
if ($UCSM -ne "")
	{
		$myucs = $UCSM
	}
else
	{
		$myucs = Read-Host "Enter UCS system IP or Hostname"
	}
if (($myucs -eq "") -or ($myucs -eq $null) -or ($Error[0] -match "PromptingException"))
	{
		Write-Output ""
		Write-Output "You have provided invalid input."
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}
else
	{
		Disconnect-Ucs
	}

#Test that UCSM is IP Reachable via Ping
Write-Output ""
Write-Output "Testing reachability to UCSM"
$ping = new-object system.net.networkinformation.ping
$results = $ping.send($myucs)
if ($results.Status -ne "Success")
	{
		Write-Output "	Can not access UCSM $myucs by Ping"
		Write-Output ""
		Write-Output "It is possible that a firewall is blocking ICMP (PING) Access."
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
				Write-Output "You have provided invalid input.  Please enter (Y/N) only."
				Write-Output "	Exiting..."
				Disconnect-Ucs
				exit
			}			
	}
else
	{
		Write-Output "	Successfully pinged UCSM: $myucs"
	}
	
#Allow Logins to single or multiple UCSM systems
$multilogin = Set-UcsPowerToolConfiguration -SupportMultipleDefaultUcs $false

#Log into UCSM
Write-Output ""
Write-Output "Logging into UCSM"

#Verify PowerShell Version to pick prompt type
if (!$UCREDENTIALS)
	{
		if (!$USAVEDCRED)
			{
				if ($PSMinimum -ge "3")
					{
						Write-Output "	Enter your UCSM credentials"
						$credu = Get-Credential -Message "UCSM(s) Login Credentials" -UserName "admin"
					}
				else
					{
						Write-Output "	Enter your UCSM credentials"
						$credu = Get-Credential
					}
			}
		else
			{
				$CredFile = import-csv $USAVEDCRED
				$Username = $credfile.UserName
				$Password = $credfile.EncryptedPassword
				$credu = New-Object System.Management.Automation.PsCredential $Username,(ConvertTo-SecureString $Password)			
			}
	}
#Log into UCSM
$myCon = Connect-Ucs $myucs -Credential $credu

#Check to see if log in was successful
if (($myucs | Measure-Object).count -ne ($myCon | Measure-Object).count) 
	{
		Write-Output "		Error Logging into UCS."
		Write-Output "		Make sure your user has login rights the UCS system and has the"
		Write-Output "		proper role/privledges to use this tool..."
		Write-Output "			Exiting..."
		Disconnect-Ucs
		exit
	}
else
	{
		if (!$UCREDENTIALS)
			{
				Write-Output "		Login Successful"
			}
		else
			{
				Write-Output "	Login Successful"
			}
	}

#Gather vHBA Information
Write-Output ""
Write-Output "Gathering vHBA information from UCSM"
$AllvHBAsA = Get-UcsVhba  | where {($_.Addr -ine "derived") -and ($_.SwitchID -eq "A")}
$AllvHBAsB = Get-UcsVhba  | where {($_.Addr -ine "derived") -and ($_.SwitchID -eq "B")}

#Put vHBA Info into a Hash Table
if ($AllvHBAsA.count -ne 0)
	{
		$vHBAInfo = @{"ServiceProfile" = $AllvHBAsA.Dn; "WWPNa" = $AllvHBAsA.Addr; "WWPNb" = $AllvHBAsB.Addr; "WWNN" = $AllvHBAsA.NodeAddr}
	}
elseif ($AllvHBAsB.Count -ne 0)
	{
		$vHBAInfo = @{"ServiceProfile" = $AllvHBAsB.Dn; "WWPNa" = $AllvHBAsA.Addr; "WWPNb" = $AllvHBAsB.Addr; "WWNN" = $AllvHBAsB.NodeAddr}
	}
	
#Check to see if any service profiles have vHBAs
if ($vHBAInfo.ServiceProfile -eq $null)
	{
		Write-Output ""
		Write-Output "	No Service Profiles configured with vHBAs"
		Write-Output "		Please correct and run this script again"
		Write-Output "			Exiting..."
		Disconnect-Ucs
		exit
	}
else
	{
		Write-Output "	Information collected"
	}

if ($SERVICEPROFILES -eq "")
	{
		#Offer user to build all or selected zones
		Write-Output ""
		Write-Output "Do you wish to create zoning information for ALL service profiles or selected?"
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
		if ($vHBAInfo.serviceprofile.count -ne 1)
			{
				[void] $objListbox.Items.Add("--ALL--")
			}
		foreach ($SP in $vHBAInfo.ServiceProfile)
			{
				$ServiceProfileFull = $SP -match "/ls-(?<content>.*)/fc-"
				$ServiceProfile = $matches['content']
				[void] $objListbox.Items.Add($ServiceProfile)
			}

		$objListbox.Height = 450
		$objForm.Controls.Add($objListbox) 
		$objForm.Topmost = $True
	
		$objForm.Add_Shown({$objForm.Activate()})
		[void] $objForm.ShowDialog()
	}
elseif ($SERVICEPROFILES -ieq "all")
	{
		$Script:SelectedObjects = "--ALL--"
	}
else
	{
		[array]$SPArray = ($SERVICEPROFILES.split(",")).trim()
		$Script:SelectedObjects = $SPArray
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
				#Hold for future use
			}
		else
			{
				$TempServiceProfile = @()
				$TempWWPNa = @()
				$TempWWPNb = @()
				$TempWWNN = @()
				$Count = $vHBAInfo.ServiceProfile.Count
				if ($Count -eq 1)
					{
						$ServiceProfileFull = $vHBAInfo.ServiceProfile -match "/ls-(?<content>.*)/fc-"
						$ServiceProfile = $matches['content']
						$TempServiceProfile += $vHBAInfo.ServiceProfile
						$TempWWPNa += $vHBAInfo.WWPNa
						$TempWWPNb += $vHBAInfo.WWPNb
						$TempWWNN += $vHBAInfo.WWNN
					}
				else
					{
						foreach ($Item in $Script:SelectedObjects)
							{
								$LoopCount = 0
								do
									{
										$ServiceProfileFull = $vHBAInfo.ServiceProfile[$LoopCount] -match "/ls-(?<content>.*)/fc-"
										$ServiceProfile = $matches['content']
										if ($ServiceProfile -eq $Item)
											{
												$TempServiceProfile += $vHBAInfo.ServiceProfile[$LoopCount]
												$TempWWPNa += $vHBAInfo.WWPNa[$LoopCount]
												$TempWWPNb += $vHBAInfo.WWPNb[$LoopCount]
												$TempWWNN += $vHBAInfo.WWNN[$LoopCount]
											}
										$LoopCount += 1
									}
								while ($LoopCount -le $Count)
							}
					}
				$vHBAInfo = $null
				$vHBAInfo = @{"ServiceProfile" = $TempServiceProfile; "WWPNa" = $TempWWPNa; "WWPNb" = $TempWWPNb; "WWNN" = $TempWWNN}
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

#Gather vSAN Information
Write-Output ""
Write-Output "Collecting vSAN information from UCSM"
$vSANa = Get-UcsFiSanCloud -Id "A" | get-UcsVsan
$vSANaID = $vSANa.id
$vSANb = Get-UcsFiSanCloud -Id "B" | get-UcsVsan
$vSANbID = $vSANb.id

#Check to see if vSANs are configured in the SAN Cloud
if (($vSANaID -eq $null) -and ($vSANbID -eq $null))
	{
		Write-Output ""
		Write-Output "	No vSAN(s) configured in the SAN Cloud of your UCS"
		Write-Output "		Please correct and run this script again"
		Write-Output "			Exiting..."
		Disconnect-Ucs
		exit
	}
else
	{
		Write-Output "	Information collected"
	}

#Create list of Config Files
if ($MANUFACTURE -eq "")
	{
		Write-Output ""
		Write-Output "Select Manufacturer from pulldown (Cisco / Brocade / Exit)"
		[array]$DropDownArray = $null
		foreach ($CFs in $ConfigFiles)
		{
		[array]$DropDownArray += $CFs.Name
		}

		#Menu Function for SAN Fabric - Brocade or Cisco
		function Return-DropDown 
			{
				$Choice = $DropDown.SelectedItem.ToString()
				$Form.Close()
			}

		#Generate GUI input box for Config File Selection
		# Script examples provided at: http://technet.microsoft.com/en-us/library/ff730949.aspx
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
		$Form = New-Object System.Windows.Forms.Form
		$Form.width = 700
		$Form.height = 150
		$Form.StartPosition = "CenterScreen"
		$Form.Text = ”SAN Fabric Manufacturer to Configure”
		$DropDown = new-object System.Windows.Forms.ComboBox
		$DropDown.Location = new-object System.Drawing.Size(100,10)
		$DropDown.Size = new-object System.Drawing.Size(550,30)
		$DropDown.Items.Add("Cisco") | Out-Null
		$DropDown.Items.Add("Brocade") | Out-Null
		$DropDown.Items.Add("EXIT") | Out-Null
		$Form.Controls.Add($DropDown)
		$DropDownLabel = new-object System.Windows.Forms.Label
		$DropDownLabel.Location = new-object System.Drawing.Size(1,10)
		$DropDownLabel.size = new-object System.Drawing.Size(255,20)
		$DropDownLabel.Text = "Manufacturer"
		$Form.Controls.Add($DropDownLabel)
		$Button = new-object System.Windows.Forms.Button
		$Button.Location = new-object System.Drawing.Size(300,50)
		$Button.Size = new-object System.Drawing.Size(75,25)
		$Button.Text = "Select"
		$Button.Add_Click({Return-DropDown})
		$form.Controls.Add($Button)
		$Form.Add_Shown({$Form.Activate()})
		$Form.ShowDialog() | Out-Null
	}
else
	{
		$Dropdown = @{"SelectedItem" = $MANUFACTURE}
	}
#Check for valid entry
if ($DropDown.SelectedItem -eq $null)
	{
		Write-Output ""
		Write-Output "Nothing Selected"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Check to see if EXIT selected
if ($DropDown.SelectedItem -eq "EXIT")
	{
		Write-Output ""
		Write-Output "You have chosen to EXIT the script"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Set SAN Fabric Manufacture
$SANFabric = $DropDown.SelectedItem

Write-Output "	Generating configuration for: $SANFabric"

#Get Data File
$ConfigFiles = dir "WWPN*.csv"
if ($ConfigFiles -eq "")
	{
		Write-Output ""
		Write-Output "Input CSV file must start with WWPN as in WWPN_List.csv and be located in: $PSScriptRoot"
		Write-Output ""
		Write-Output "Input CSV file must be formatted as below:"
		Write-Output "ZoneSetName		,	NameA		,	WWPN_A			,		NameB		,	WWPN_B"
		Write-Output "LabFabric		,	Cntrl-A-1	,	50:00:00:00:00:00:AA:11	,		Cntrl-A-2	,	50:00:00:00:00:00:AB:22"
		Write-Output "			,	Cntrl-A-3	,	50:00:00:00:00:00:AA:33	,		Cntrl-A-4	,	50:00:00:00:00:00:AB:44"
		Write-Output "			,	Cntrl-B-1	,	50:00:00:00:00:00:BA:15	,		Cntrl-B-2	,	50:00:00:00:00:00:BB:26"
		Write-Output "			,	Cntrl-B-3	,	50:00:00:00:00:00:BA:37	,		Cntrl-B-4	,	50:00:00:00:00:00:BB:48"
		Write-Output ""
		Write-Output "ZoneSetName is the name of the Zone Set"
		Write-Output "NameA is the name of the array controller port connecting for SAN fabric A"
		Write-Output "WWPN_A is the WWPN of the array controller port connecting to SAN fabric A"
		Write-Output "NameB is the name of the array controller port connecting for SAN fabric B"
		Write-Output "WWPN_B is the WWPN of the array controller port connecting to SAN fabric B"
		Write-Output ""
		Write-Output "	Please correct this issue and try again"
		Write-Output "		Exiting..."
		Disconnect-Ucs
		exit
	}
elseif ($WWPNCSV -eq "")
	{
		Write-Output ""
		Write-Output "Select CSV File from pulldown (CSV files or EXIT)"
	}

<# Sample CSV file and format.  Save this as WWPNtest.csv and then open in Excel to see the format and then create your own real version.
ZoneSetName,NameA,WWPN_A,NameB,WWPN_B
LabFabric,Cntrl-A-1,50:00:00:00:00:00:AA:11,Cntrl-A-2,50:00:00:00:00:00:AB:22
,Cntrl-A-3,50:00:00:00:00:00:AA:33,Cntrl-A-4,50:00:00:00:00:00:AB:44
,Cntrl-B-1,50:00:00:00:00:00:BA:15,Cntrl-B-2,50:00:00:00:00:00:BB:26
,Cntrl-B-3,50:00:00:00:00:00:BA:37,Cntrl-B-4,50:00:00:00:00:00:BB:48
#>

#Create list of Config Files
if ($WWPNCSV -eq "")
	{
		[array]$DropDownArray = $null
		foreach ($CFs in $ConfigFiles)
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
		
		#Generate GUI input box for Config File Selection
		# Script examples provided at: http://technet.microsoft.com/en-us/library/ff730949.aspx
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
		$Form = New-Object System.Windows.Forms.Form
		$Form.width = 700
		$Form.height = 150
		$Form.StartPosition = "CenterScreen"
		$Form.Text = ”WWPN Targets List CSV to use”
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
		$DropDownLabel.Text = "CSV File"
		$Form.Controls.Add($DropDownLabel)
		$Button = new-object System.Windows.Forms.Button
		$Button.Location = new-object System.Drawing.Size(300,50)
		$Button.Size = new-object System.Drawing.Size(75,25)
		$Button.Text = "Select"
		$Button.Add_Click({Return-DropDown})
		$form.Controls.Add($Button)
		$Form.Add_Shown({$Form.Activate()})
		$Form.ShowDialog() | Out-Null
	}
else
	{
		$Dropdown = @{"SelectedItem" = $PSScriptRoot+"\"+$WWPNCSV}
	}

#Check for valid entry
if ($DropDown.SelectedItem -eq $null)
	{
		Write-Output ""
		Write-Output "Nothing Selected"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Check to see if EXIT selected
if ($DropDown.SelectedItem -eq "EXIT")
	{
		Write-Output ""
		Write-Output "You have chosen to EXIT the script"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Load the data configuration file
$CSVFile = $DropDown.SelectedItem
$CSVFileDir = $PSScriptRoot
cd $CSVFileDir
$TargetInfo = Import-Csv $CSVFile

#Validating data file
Write-Output ""
Write-Output "Validating Data File"
$FileThere = Test-Path $CSVFile
if ($FileThere -eq $false)
	{
		Write-Output ""
		Write-Output "The WWPN Targets file specified does not exist"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}
$CharacterTest = [regex]"^[A-Za-z0-9_-]*$"
$WWPNtest = [regex]"^[a-fA-F0-9:]*$"
if ((($TargetInfo.ZoneSetName[0]).Length -le 64) -and (($CharacterTest.Match($TargetInfo.ZoneSetName[0]).Success)))
	{
		#Hold for future option
	}
else
	{
		Write-Output "	Invalid entry: ZoneSetName: $TargetInfo.ZoneSetName[0]"
		Write-Output "		Please correct this error in the data file and try again"
		Write-Output "			Exiting..."
		Disconnect-Ucs
		exit
	}
$TestLoop = 0
foreach ($TI in $TargetInfo)
	{
		if ($AllvHBAsA.Count -ne 0)
			{
				if ((($TI.NameA[$TestLoop]).Length -le 64) -and (($CharacterTest.Match($TargetInfo.NameA[$TestLoop]).Success)))
					{
						#Hold for future option
					}
				else
					{
						Write-Output "	Invalid entry: NameA: $TI.NameA[$TestLoop]"
						Write-Output "		Please correct this error in the data file and try again"
						Write-Output "			Exiting..."
						Disconnect-Ucs
						exit
					}
				if ((($TI.WWPN_A[$TestLoop]) -match "\w{2}\:\w{2}\:\w{2}\:\w{2}\:\w{2}\:\w{2}\:\w{2}\:\w{2}") -and (($WWPNtest.Match(($TI.WWPN_A[$TestLoop])).Success)))
					{
						Write-Output "	Invalid entry: WWPN_A: $TI.WWPN_A[$TestLoop]"
						Write-Output "		Please correct this error in the data file and try again"
						Write-Output "			Exiting..."
						Disconnect-Ucs
						exit
					}
			}
		if ($AllvHBAsB.Count -ne 0)
			{
				if ((($TI.NameB[$TestLoop]).Length -le 64) -and (($CharacterTest.Match($TI.NameB[$TestLoop]).Success)))
					{
						#Hold for future option
					}
				else
					{
						Write-Output "	Invalid entry: NameB: $TI.NameB[$TestLoop]"
						Write-Output "		Please correct this error in the data file and try again"
						Write-Output "			Exiting..."
						Disconnect-Ucs
						exit
					}
				if ((($TI.WWPN_B[$TestLoop]) -match "\w{2}\:\w{2}\:\w{2}\:\w{2}\:\w{2}\:\w{2}\:\w{2}\:\w{2}") -and (($WWPNtest.Match(($TI.WWPN_B[$TestLoop])).Success)))
					{
						Write-Output "	Invalid entry: WWPN_B: $TI.WWPN_B[$TestLoop]"
						Write-Output "		Please correct this error in the data file and try again"
						Write-Output "			Exiting..."
						Disconnect-Ucs
						exit
					}
			}
	}
Write-Output "	CSV file is valid"

Write-Output "		You will be using the following CSV file:"
Write-Output "			$CSVFile"

#Menu Function for Destination - File or Equipment
if ($OUTPUT -eq "")
	{
		Write-Output ""
		Write-Output "Select destination of configuration from pulldown (File / Equipment / EXIT)"
		function Return-DropDown 
			{
				$Choice = $DropDown.SelectedItem.ToString()
				$Form.Close()
			}
		
		#Generate GUI input box for Config File Selection
		# Script examples provided at: http://technet.microsoft.com/en-us/library/ff730949.aspx
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
		$Form = New-Object System.Windows.Forms.Form
		$Form.width = 700
		$Form.height = 150
		$Form.StartPosition = "CenterScreen"
		$Form.Text = ”Destination of Configuration”
		$DropDown = new-object System.Windows.Forms.ComboBox
		$DropDown.Location = new-object System.Drawing.Size(100,10)
		$DropDown.Size = new-object System.Drawing.Size(550,30)
		$DropDown.Items.Add("File") | Out-Null
		$DropDown.Items.Add("Equipment") | Out-Null
		$DropDown.Items.Add("EXIT") | Out-Null
		$Form.Controls.Add($DropDown)
		$DropDownLabel = new-object System.Windows.Forms.Label
		$DropDownLabel.Location = new-object System.Drawing.Size(1,10)
		$DropDownLabel.size = new-object System.Drawing.Size(255,20)
		$DropDownLabel.Text = "Destination"
		$Form.Controls.Add($DropDownLabel)
		$Button = new-object System.Windows.Forms.Button
		$Button.Location = new-object System.Drawing.Size(300,50)
		$Button.Size = new-object System.Drawing.Size(75,25)
		$Button.Text = "Select"
		$Button.Add_Click({Return-DropDown})
		$form.Controls.Add($Button)
		$Form.Add_Shown({$Form.Activate()})
		$Form.ShowDialog() | Out-Null
	}
else
	{
		$Dropdown = @{"SelectedItem" = $OUTPUT}
	}

#Check for valid entry
if ($DropDown.SelectedItem -eq $null)
	{
		Write-Output ""
		Write-Output "Nothing Selected"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Check to see if EXIT selected
if ($DropDown.SelectedItem -eq "EXIT")
	{
		Write-Output ""
		Write-Output "You have chosen to EXIT the script"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Set Output Location (File or Equipment)
$Destination = $DropDown.SelectedItem

Write-Output ""
Write-Output "Sending configuration to: $Destination"

#Set column number
$ColumnLoop = 1

#Reset Loop Counter
$Loop = 0

#ZoneSet Creation Flag
$ZoneSetCreate = 1

#Cisco Config T entry
$ConfigT = 0

#Determine how many Service Profiles there are
$LoopMax = ($vHBAInfo.ServiceProfile | where {$_ -ne $null}).Count

#Create Empty Arrays for Commands
$FabricACommands = @()
$FabricBCommands = @()

Write-Output ""
Write-Output "Generating Configuration"

#Fabric A
if ($AllvHBAsA.Count -ne 0)
	{
		if ($LoopMax -ne 0)
			{
				do 
					{	
						$ServiceProfileFull = $vHBAInfo.ServiceProfile[$Loop] -match "/ls-(?<content>.*)/fc-"
						if ($ServiceProfileFull -eq $true)
							{
								$ServiceProfile = $matches['content']
							}
						else
							{
								$ServiceProfile = " "
							} 
						$WWPNa = $vHBAInfo.WWPNa[$Loop]
						$WWNN = $vHBAInfo.WWNN[$Loop]

						#Generate Configuration Files
						#For Cisco
						if ($SANFabric -eq "Cisco")
							{
								$ZoneMember = 1
								if ($ConfigT -eq 0)
									{
										if ($Destination -eq "File")
											{
												$FabricACommands += "configure terminal"
											}
										if ($Destination -eq "File")
											{
												$FabricACommands += ""
											}
									}
								$ConfigT += 1
								foreach ($TI in $TargetInfo)
									{
										if ((($ZoneMember -eq 1) -or ($Loop -eq 0)) -and ($TI.WWPN_A -ne ""))
											{
												$FabricACommands += "device-alias mode enhanced"
												$FabricACommands += "device-alias database"
											}
										if (($ZoneMember -eq 1) -and ($TI.WWPN_A -ne ""))
											{								
												$FabricACommands += " device-alias name $ServiceProfile pwwn $WWPNa"
											}
										$TIName = $TI.NameA
										$TIWWPN = $TI.WWPN_A
										$TIZoneSet = $TargetInfo.ZoneSetName[0]
										if (($Loop -eq 0) -and ($TI.WWPN_A -ne ""))
											{
												$FabricACommands += " device-alias name $TIName pwwn $TIWWPN"
											}
										if ((($ZoneMember -eq 1) -or ($Loop -eq 0)) -and ($TI.WWPN_A -ne ""))
											{
												$FabricACommands += "device-alias commit"
												if ($Destination -eq "File")
													{
														$FabricACommands += ""
													}
											}
										if ($TI.WWPN_A -ne "")
											{
												$FabricACommands += "zone name $ServiceProfile-$ZoneMember vsan $vSANaID"
												$FabricACommands += " member device-alias $ServiceProfile"
												$FabricACommands += " member device-alias $TIName"
												if ($Destination -eq "File")
													{
														$FabricACommands += ""
													}
												$FabricACommands += "zoneset name $TIZoneSet vsan $vSANaID"
												$FabricACommands += " member $ServiceProfile-$ZoneMember"
											}
										$ZoneMember += 1
										if (($Destination -eq "File") -and ($TI.WWPN_A -ne ""))
											{
												$FabricACommands += ""
											}
									}
								if ($Destination -eq "Equipment")
									{
										$FabricACommands += "!FLAG!"
									}
							}
						#For Brocade
						else
							{
								$ZoneMember = 1
								foreach ($TI in $TargetInfo)
									{
										if (($ZoneMember -eq 1) -and ($TI.WWPN_A -ne ""))
											{								
												$FabricACommands += "alicreate $ServiceProfile, $WWPNa"
											}
										$TIName = ($TI.NameA).replace("-","_")
										$TIWWPN = $TI.WWPN_A
										$TIZoneSet = $TargetInfo.ZoneSetName[0]
										if (($Loop -eq 0) -and ($TI.WWPN_A -ne ""))
											{
												$FabricACommands += "alicreate $TIName, $TIWWPN"
											}
										if ($TI.WWPN_A -ne "")
											{
												$FabricACommands += "zonecreate $ServiceProfile"+"__"+"$ZoneMember, $ServiceProfile"
												$FabricACommands += "zoneadd $ServiceProfile"+"__"+"$ZoneMember, $TIName"
											}
										if (($ZoneSetCreate -eq 1) -and ($TI.WWPN_A -ne ""))
											{								
												$FabricACommands += "cfgcreate $TIZoneSet, $ServiceProfile"+"__"+"$ZoneMember"
												if ($Destination -eq "File")
													{
														$FabricACommands += ""
													}
												$ZoneSetCreate += 1
											}
										elseif ($TI.WWPN_A -ne "")
											{
												$FabricACommands += "cfgadd $TIZoneSet, $ServiceProfile"+"__"+"$ZoneMember"
												if ($Destination -eq "File")
													{
														$FabricACommands += ""
													}
											}
										$ZoneMember += 1
									}
								if ($Destination -eq "Equipment")
									{
										$FabricACommands += "!FLAG!"
									}
							}
				
						#Increment Counters
						$Loop += 1
						$ColumnLoop += 1
					}
				while ($Loop -lt $LoopMax)
			}
		
		if ($SANFabric -eq "Cisco")
			{
				if ($Destination -eq "File")
					{
						$FabricACommands += "zoneset activate name $TIZoneSet vsan $vSANaID"
						$FabricACommands += ""
						$FabricACommands += "copy running-config startup-config"
					}
			}
		else
			{
				if ($Destination -eq "File")
					{
						$FabricACommands += "echo y | cfgsave"
						$FabricACommands += ""
						$FabricACommands += "echo y | cfgenable $TIZoneSet"
					}
			}
		
		#Reset Loop Counter
		$Loop = 0
		
		#ZoneSet Creation Flag
		$ZoneSetCreate = 1
		
		#Cisco Config T entry
		$ConfigT = 0
	}
	
#Fabric B
if ($AllvHBAsB.Count -ne 0)
	{
		if ($LoopMax -ne 0)
			{
				do 
					{
						$ServiceProfileFull = $vHBAInfo.ServiceProfile[$Loop] -match "/ls-(?<content>.*)/fc-"
						if ($ServiceProfileFull -eq $true)
							{
								$ServiceProfile = $matches['content']
							}
						else
							{
								$ServiceProfile = " "
							} 
						$WWPNb = $vHBAInfo.WWPNb[$Loop]
						$WWNN = $vHBAInfo.WWNN[$Loop]
										
						#Generate Configuration Files
						#For Cisco
						if ($SANFabric -eq "Cisco")
							{
								$ZoneMember = 1
								if ($ConfigT -eq 0)
									{
										if ($Destination -eq "File")
											{
												$FabricBCommands += "configure terminal"
												$FabricBCommands += ""
											}
									}
								$ConfigT += 1
								foreach ($TI in $TargetInfo)
									{					
										if ((($ZoneMember -eq 1) -or ($Loop -eq 0)) -and ($TI.WWPN_B -ne ""))
											{
												$FabricBCommands += "device-alias mode enhanced"
												$FabricBCommands += "device-alias database"
											}
										if (($ZoneMember -eq 1) -and ($TI.WWPN_B -ne ""))
											{								
												$FabricBCommands += " device-alias name $ServiceProfile pwwn $WWPNb"
											}
										$TIName = $TI.NameB
										$TIWWPN = $TI.WWPN_B
										$TIZoneSet = $TargetInfo.ZoneSetName[0]
										if (($Loop -eq 0) -and ($TI.WWPN_B -ne ""))
											{								
												$FabricBCommands += " device-alias name $TIName pwwn $TIWWPN"
											}
										if ((($ZoneMember -eq 1) -or ($Loop -eq 0)) -and ($TI.WWPN_B -ne ""))
											{
												$FabricBCommands += "device-alias commit"
												if ($Destination -eq "File")
													{
														$FabricBCommands += ""
													}
											}
										if ($TI.WWPN_B -ne "")
											{
												$FabricBCommands += "zone name $ServiceProfile-$ZoneMember vsan $vSANbID"
												$FabricBCommands += " member device-alias $ServiceProfile"
												$FabricBCommands += " member device-alias $TIName"
												if ($Destination -eq "File")
													{
														$FabricBCommands += ""
													}
												$FabricBCommands += "zoneset name $TIZoneSet vsan $vSANbID"
												$FabricBCommands += " member $ServiceProfile-$ZoneMember"
											}
										$ZoneMember += 1
										if (($Destination -eq "File") -and ($TI.WWPN_B -ne ""))
											{
												$FabricBCommands += ""
											}
									}
								if ($Destination -eq "Equipment")
									{
										$FabricBCommands += "!FLAG!"
									}
							}
						#For Brocade
						else
							{
								$ZoneMember = 1
								foreach ($TI in $TargetInfo)
									{
										if (($ZoneMember -eq 1) -and ($TI.WWPN_B -ne ""))
											{								
												$FabricBCommands += "alicreate $ServiceProfile, $WWPNb"
											}
										$TIName = ($TI.NameB).replace("-","_")
										$TIWWPN = $TI.WWPN_B
										$TIZoneSet = $TargetInfo.ZoneSetName[0]
										if (($Loop -eq 0) -and ($TI.WWPN_B -ne ""))
											{
												$FabricBCommands += "alicreate $TIName, $TIWWPN"
											}
										if ($TI.WWPN_B -ne "")
											{
												$FabricBCommands += "zonecreate $ServiceProfile"+"__"+"$ZoneMember, $ServiceProfile"
												$FabricBCommands += "zoneadd $ServiceProfile"+"__"+"$ZoneMember, $TIName"
											}
										if (($ZoneSetCreate -eq 1) -and ($TI.WWPN_B -ne ""))
											{								
												$FabricBCommands += "cfgcreate $TIZoneSet, $ServiceProfile"+"__"+"$ZoneMember"
												if ($Destination -eq "File")
													{
														$FabricBCommands += ""
													}
												$ZoneSetCreate += 1
											}
										elseif ($TI.WWPN_B -ne "")
											{
												$FabricBCommands += "cfgadd $TIZoneSet, $ServiceProfile"+"__"+"$ZoneMember"
												if ($Destination -eq "File")
													{
														$FabricBCommands += ""
													}
											}
										$ZoneMember += 1
									}
								if ($Destination -eq "Equipment")
									{
										$FabricBCommands += "!FLAG!"
									}
							}
				
						#Increment Counters
						$Loop += 1
						$ColumnLoop += 1
					}
				while ($Loop -lt $LoopMax)
			}

		if ($SANFabric -eq "Cisco")
			{				
				if ($Destination -eq "File")
					{
						$FabricBCommands += "zoneset activate name $TIZoneSet vsan $vSANbID"
						$FabricBCommands += ""
						$FabricBCommands += "copy running-config startup-config"
					}
			}
		else
			{				
				if ($Destination -eq "File")
					{
						$FabricBCommands += "echo y | cfgsave"
						$FabricBCommands += ""
						$FabricBCommands += "echo y | cfgenable $TIZoneSet"
					}
			}
	}	
Write-Output "	Configuration Created"

#Run PUTTY/PLINK commands in powershell
#http://www.christowles.com/2011/06/how-to-ssh-from-powershell-using.html
Function Invoke-SSHCommands 
	{
		Param($Hostname,$Username,$Password, $CommandArray, $PlinkAndPath, $ConnectOnceToAcceptHostKey)

		$Target = $Username + '@' + $Hostname
 		$plinkoptions = "-ssh $Target -pw $Password"
 
 		#Build ssh Commands
 		$remoteCommand = ""
 		$CommandArray | % {$remoteCommand += [string]::Format('{0}; ', $_)}
		$CommandArray = "'"+$CommandArray+"'"
				
 		#plink prompts to accept client host key. This section will login and accept the host key then logout.
		if($ConnectOnceToAcceptHostKey -eq 1)
			{
				$PlinkCommand  = [string]::Format('echo y | & "{0}" {1} exit',
				$PlinkAndPath, $plinkoptions )
				$ErrorActionPreference = "SilentlyContinue"
  				$msg = Invoke-Expression $PlinkCommand  -ErrorAction SilentlyContinue | Out-Null
				Start-Sleep -Seconds 2
				$ErrorActionPreference = $ErrorLevel
 			}
 
		#Format plink command
		$PlinkCommand = [string]::Format('& "{0}" {1} {2}', $PlinkAndPath, $plinkoptions, $CommandArray)

		#Ready to run the following command
		$ErrorActionPreference = "SilentlyContinue"		
		$msg = Invoke-Expression $PlinkCommand  -ErrorAction Inquire | Out-Null
		$ErrorActionPreference = $ErrorLevel
	}
 		
#Convert secure password to plain text
#http://www.powershelladmin.com/wiki/Powershell_prompt_for_password_convert_securestring_to_plain_text
function ConvertFrom-SecureToPlain
	{
		param([Parameter(Mandatory=$true)][System.Security.SecureString] $SecurePassword)
  
		# Create a "password pointer"
		$PasswordPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    
		# Get the plain text version of the password
		$PlainTextPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto($PasswordPointer)
    
		# Free the pointer
		[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($PasswordPointer)
   
		# Return the plain text password
		$PlainTextPassword
	}

#Save commands to file
if ($Destination -ieq "File")
	{
		Write-Output ""
		$UCSName = $myCon.Ucs
		$FileNameA = $PSScriptRoot+"\FabricA_"+$TargetInfo.ZoneSetName[0]+"_"+$UCSName+".txt"
		$FileNameB = $PSScriptRoot+"\FabricB_"+$TargetInfo.ZoneSetName[0]+"_"+$UCSName+".txt"
		Write-Output "Saving configuration files to:"
		if ($AllvHBAsA.Count -ne 0)
			{
				Write-Output "	$FileNameA"
				$FabricACommands | Out-File $FileNameA
			}
		if ($AllvHBAsB.Count -ne 0)
			{
				Write-Output "	$FileNameB"
				$FabricBCommands | Out-File $FileNameB
			}
		Write-Output "		Files Saved..."
	}
#Save commands to equipment
elseif ($Destination -ieq "Equipment")
	{
		Write-Output ""
		Write-Output "Sending configuration to your equipment"
		Write-Output "	Sending commands to your equipment over SSH can be very slow."
		Write-Output "		Please wait..."
		
		#Fabric A Access
		if ($AllvHBAsA.Count -ne 0)
			{
				Write-Output ""
				Write-Output "Connecting to SAN Fabric A."
				if ($FABRICA -eq "")
					{
						Write-Output "	Enter IP Address or Hostname."
						Write-Output "	Just press OK if you don't want to configure SAN Fabric A."
						Write-Output "	Pressing CANCEL will exit the script."
						$Hostname = ""
						$Hostname = Read-Host "IP Address or Hostname of Fabric Switch A."
					}
				else
					{
						$Hostname = $FABRICA
					}
				if (($Hostname -eq "") -or ($Hostname -eq $null))
					{
						Write-Output ""
						Write-Output "	You did not enter anything..."
						Write-Output "		Skipping the configuration of SAN Fabric A"
					}
				else
					{
						#Test that Fabric A is IP Reachable via Ping
						Write-Output "	Testing reachability to SAN Fabric A: $Hostname"
						$ping = new-object system.net.networkinformation.ping
						$results = $ping.send($Hostname)
						if ($results.Status -ne "Success")
							{
								Write-Output "		Can not access SAN Fabric A $Hostname by Ping"
								Write-Output ""
								Write-Output "It is possible that a firewall is blocking ICMP (PING) Access."
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
										Write-Output "You have provided invalid input.  Please enter (Y/N) only."
										Write-Output "	Exiting..."
										Disconnect-Ucs
										exit
									}			
							}
						else
							{
								Write-Output "		Successfully pinged SAN Fabric A: $Hostname"
								Write-Output ""
							}
						if (!$ASAVEDCRED)
							{
								if (!$ACREDENTIALS)
									{
										Write-Output "Enter SAN Fabric A Switch Credentials"
										$creda = Get-Credential -Message "Enter SAN Fabric A Credentials"
									}
							}
						else
							{
								$CredFile = import-csv $ASAVEDCRED
								$Username = $credfile.UserName
								$Password = $credfile.EncryptedPassword
								$creda = New-Object System.Management.Automation.PsCredential $Username,(ConvertTo-SecureString $Password)			
							}
						$Username = $creda.UserName.TrimStart("\")
						$Password = ConvertFrom-SecureToPlain -SecurePassword $creda.Password
						if (($Username -eq "") -or ($Password -eq ""))
							{
								Write-Output ""
								Write-Output "	You did not enter a username and/or password"
								Write-Output "		Exiting..."
								Disconnect-Ucs
								exit
							}
						else
							{
								Write-Output "Connected to SAN Fabric Switch A: $Hostname"
							}

						#Fabric A Commands
 						Write-Output "	Please wait while we send commands to: $Hostname"
						Write-Output ""
						Write-Output "***If this is the first time connecting to this switch then you will see a"
						Write-Output "security message.  Ignore it and continue to wait."
						Write-Output "Script will auto accept***"
						Write-Output ""
						#For Cisco
						if ($SANFabric -eq "Cisco")
							{
								$AcceptHostKey = 1
								$Commands = '"configure terminal" ; "'
								foreach ($Item in $FabricACommands)
									{
										if ($Item -ne "!FLAG!")
											{
												$Commands += $Item+'" ; "'
											}
										else
											{
												$Commands += 'exit"'
												#Use plink to configure your settings
												Invoke-SSHCommands -User $Username  `
								 				-Hostname $Hostname `
								 				-Password $Password `
								 				-PlinkAndPath $PlinkAndPath `
								 				-CommandArray $Commands `
												-ConnectOnceToAcceptHostKey $AcceptHostKey
												$AcceptHostKey += 1
												$Commands = '"configure terminal" ; "'
											}
									}
								$Commands += 'zoneset activate name '+$TIZoneSet+' vsan '+$vSANaID+'" ; "copy running-config startup-config" ; "exit"'
								Write-Output "Fabric A Configuration Complete...Please Wait..."
								#Use plink to configure your settings
								Invoke-SSHCommands -User $Username  `
		 						-Hostname $Hostname `
		 						-Password $Password `
		 						-PlinkAndPath $PlinkAndPath `
		 						-CommandArray $Commands `
								-ConnectOnceToAcceptHostKey $AcceptHostKey
								$AcceptHostKey += 1
							}
						#For Brocade
						else
							{
								$AcceptHostKey = 1
								$Commands = '"'
								foreach ($Item in $FabricACommands)
									{
										if ($Item -ne "!FLAG!")
											{
												$Commands += $Item+'" ; "'
											}
										else
											{
												$Commands += 'echo y | cfgsave" ; "exit"'
												#Use plink to configure your settings
												Invoke-SSHCommands -User $Username  `
								 				-Hostname $Hostname `
								 				-Password $Password `
								 				-PlinkAndPath $PlinkAndPath `
								 				-CommandArray $Commands `
												-ConnectOnceToAcceptHostKey $AcceptHostKey
												$AcceptHostKey += 1
												$Commands = '"'
											}
									}
								$Commands += 'echo y | cfgenable '+$TIZoneSet+'" ; "echo y | cfgsave" ; "exit"'
								Write-Output "Fabric A Configuration Complete"
								#Use plink to configure your settings
								Invoke-SSHCommands -User $Username  `
		 						-Hostname $Hostname `
		 						-Password $Password `
		 						-PlinkAndPath $PlinkAndPath `
		 						-CommandArray $Commands `
								-ConnectOnceToAcceptHostKey $AcceptHostKey
								$AcceptHostKey += 1
							}
					}
			}	

		#Fabric B Access
		if ($AllvHBAsB.Count -ne 0)
			{
				Write-Output ""
				Write-Output "Connecting to SAN Fabric B."
				if ($FABRICB -eq "")
					{
						Write-Output "	Enter IP Address or Hostname."
						Write-Output "	Just press OK if you don't want to configure SAN Fabric B."
						Write-Output "	Pressing CANCEL will exit the script."
						$Hostname = ""
						$Hostname = Read-Host "IP Address or Hostname of Fabric Switch B."
					}
				else
					{
						$Hostname = $FABRICB
					}
				if (($Hostname -eq "") -or ($Hostname -eq $null))
					{
						Write-Output ""
						Write-Output "	You did not enter anything..."
						Write-Output "		Skipping the configuration of SAN Fabric B"
					}
				else
					{
						#Test that Fabric B is IP Reachable via Ping
						Write-Output "	Testing reachability to SAN Fabric B: $Hostname"
						$ping = new-object system.net.networkinformation.ping
						$results = $ping.send($Hostname)
						if ($results.Status -ne "Success")
							{
								Write-Output "	Can not access SAN Fabric B $Hostname by Ping"
								Write-Output ""
								Write-Output "It is possible that a firewall is blocking ICMP (PING) Access."
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
										Write-Output "You have provided invalid input.  Please enter (Y/N) only."
										Write-Output "	Exiting..."
										Disconnect-Ucs
										exit
									}			
							}
						else
							{
								Write-Output "		Successfully pinged SAN Fabric B: $Hostname"
								Write-Output ""
							}
						if (!$BSAVEDCRED)
							{
								if (!$BCREDENTIALS)
									{
										Write-Output "Enter SAN Fabric B Switch Credentials"
										$credb = Get-Credential -Message "Enter SAN Fabric B Credentials"
									}
							}
						else
							{
								$CredFile = import-csv $BSAVEDCRED
								$Username = $credfile.UserName
								$Password = $credfile.EncryptedPassword
								$credb = New-Object System.Management.Automation.PsCredential $Username,(ConvertTo-SecureString $Password)			
							}
						$Username = $credb.UserName.TrimStart("\")
						$Password = ConvertFrom-SecureToPlain -SecurePassword $credb.Password
						if (($Username -eq "") -or ($Password -eq ""))
							{
								Write-Output ""
								Write-Output "	You did not enter a username and/or password"
								Write-Output "		Exiting..."
								Disconnect-Ucs
								exit
							}
						else
							{
								Write-Output "Connected to SAN Fabric Switch B: $Hostname"
							}

						#Fabric B Commands
 						Write-Output "	Please wait while we send commands to: $Hostname"
						Write-Output ""
						Write-Output "***If this is the first time connecting to this switch then you will see a"
						Write-Output "security message.  Ignore it and continue to wait."
						Write-Output "Script will auto accept***"
						Write-Output ""
						#For Cisco
						if ($SANFabric -eq "Cisco")
							{
								$AcceptHostKey = 1
								$Commands = '"configure terminal" ; "'
								foreach ($Item in $FabricBCommands)
									{
										if ($Item -ne "!FLAG!")
											{
												$Commands += $Item+'" ; "'
											}
										else
											{
												$Commands += 'exit"'
												#Use plink to configure your settings
												Invoke-SSHCommands -User $Username  `
								 				-Hostname $Hostname `
								 				-Password $Password `
								 				-PlinkAndPath $PlinkAndPath `
								 				-CommandArray $Commands `
												-ConnectOnceToAcceptHostKey $AcceptHostKey
												$AcceptHostKey += 1
												$Commands = '"configure terminal" ; "'
											}
									}
								$Commands += 'zoneset activate name '+$TIZoneSet+' vsan '+$vSANbID+'" ; "copy running-config startup-config" ; "exit"'
								Write-Output "Fabric B Configuration Complete...Please Wait..."
								#Use plink to configure your settings
								Invoke-SSHCommands -User $Username  `
		 						-Hostname $Hostname `
		 						-Password $Password `
		 						-PlinkAndPath $PlinkAndPath `
		 						-CommandArray $Commands `
								-ConnectOnceToAcceptHostKey $AcceptHostKey
								$AcceptHostKey += 1
							}
						#For Brocade
						else
							{
								$AcceptHostKey = 1
								$Commands = '"'
								foreach ($Item in $FabricBCommands)
									{
										if ($Item -ne "!FLAG!")
											{
												$Commands += $Item+'" ; "'
											}
										else
											{
												$Commands += 'echo y | cfgsave" ; "exit"'
												#Use plink to configure your settings
												Invoke-SSHCommands -User $Username  `
								 				-Hostname $Hostname `
								 				-Password $Password `
								 				-PlinkAndPath $PlinkAndPath `
								 				-CommandArray $Commands `
												-ConnectOnceToAcceptHostKey $AcceptHostKey
												$AcceptHostKey += 1
												$Commands = '"'
											}
									}
								$Commands += 'echo y | cfgenable '+$TIZoneSet+'" ; "echo y | cfgsave" ; "exit"'
								Write-Output "Fabric B Configuration Complete"
								#Use plink to configure your settings
								Invoke-SSHCommands -User $Username  `
		 						-Hostname $Hostname `
		 						-Password $Password `
		 						-PlinkAndPath $PlinkAndPath `
		 						-CommandArray $Commands `
								-ConnectOnceToAcceptHostKey $AcceptHostKey
								$AcceptHostKey += 1
							}
					}
			}
	}

#Clear username and password info from prying eyes
$credu = $null
$creda = $null
$credb = $null
$Username = $null
$Password = $null

#Exit Script
Write-Output ""
Write-Output "Script Complete"
Write-Output "	Exiting..."
Disconnect-Ucs
exit