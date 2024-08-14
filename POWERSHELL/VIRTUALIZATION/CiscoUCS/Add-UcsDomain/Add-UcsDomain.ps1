<#

.SYNOPSIS
	This script builds a new UCS system up with a standard configuration based on naming and features used by me.

.DESCRIPTION
	This script will completely build a UCS system from the ground up.
	Once you have physically installed, cabled and done the initial steps via the serial port to assign base access information running this script will remove all ucs default pools, policies and templates.  It will then build all pools, policies and templates for your UCSM based on the information configured in an Answer File.

.EXAMPLE
	Add-UcsDomain.ps1
	This script can only be run without command line parameters.  User will be prompted for all parameters and options required

.EXAMPLE
	Add-UcsDomain.ps1 -file "Data-Laptop Emulator.ps1" -ucred
	-file -- Answer file.  Must be located in the same folder as this script
	-ucred -- UCS Manager Credential Switch -- Adding this switch will immediately prompt you for your UCSM username and password

.NOTES
	Author: Joe Martin
	Email: joemar@cisco.com
	Company: Cisco Systems, Inc.
	Version: v0.9.22
	Date: 12/7/2015
	Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
	Select configuration Data file to use 
	UCSM Username and Password
	Y or N to continue to configure the displayed system

.OUTPUTS
	Only display information is output to the host display for visual confirmation of configuration functions.
	
.LINK
	http://communities.cisco.com/people/joemar/content

#>

#Command Line Parameters
param(
	[string]$FILE,				# Answer file.  Must be located in the same folder as this script
	[switch]$UCREDENTIALS		# UCSM Credentials (Username and Password)
)

Try
{
#Unload all Data- modules
Remove-Module Data-*

#Setup PowerShell console colors for compatibility with my script colors
$PowerShellWindow = (Get-Host).UI.RawUI
$PowerShellWindow.BackgroundColor = "White"
$PowerShellWindow.ForegroundColor = "Black"

#Clear the screen
clear-host

Write-Host -ForegroundColor White -BackgroundColor DarkBlue "Script Running..."
Write-Host ""

#Gather any credentials requested from command line
if ($UCREDENTIALS)
	{
		$cred = Get-Credential -Message "Enter UCSM Credentials"
	}

#Verify PowerShell Version for script support
$PSVersion = $psversiontable.psversion
$PSMinimum = $PSVersion.Major
if ($PSMinimum -ge "3")
	{
	}
else
	{
		Write-Host -ForegroundColor Red "This script requires PowerShell version 3 or above"
		Write-Host -ForegroundColor Red "Please update your system and try again."
		Write-Host -ForegroundColor Red "You can download PowerShell updates here:"
		Write-Host -ForegroundColor Red "	http://search.microsoft.com/en-us/DownloadResults.aspx?rf=sp&q=powershell+4.0+download"
		Write-Host -ForegroundColor Red "If you are running a version of Windows before 7 or Server 2008R2 you need to update to be supported"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

#Tell the user what this script does
Write-Host -ForegroundColor DarkBlue "This script will completely build a UCS system from the ground up."
Write-Host -ForegroundColor DarkBlue "Once you have physically installed, cabled and done the initial steps via the serial port to assign base access information"
Write-Host -ForegroundColor DarkBlue "running this script will remove all ucs default pools, policies and templates.  It will then build all pools, policies and templates"
Write-Host -ForegroundColor DarkBlue "for your UCSM based on the information configured in an Answer File."
Write-Host ""
Write-Host -ForegroundColor DarkBlue "Prerequisites:"
Write-Host -ForegroundColor DarkBlue "	Place this script and your answer file into the same directory"
Write-Host -ForegroundColor DarkBlue "	Make sure your system is running PowerShell and PowerTool"
Write-Host -ForegroundColor DarkBlue "	Make sure you have an administrative login into UCSM"
Write-Host -ForegroundColor DarkBlue "	Make sure the client running this script has IP access to UCSM"
Write-Host ""
Write-Host -ForegroundColor DarkBlue "NOTE:"
Write-Host -ForegroundColor DarkBlue "	Many of the settings are automatic and are based on my personal best practices."
Write-Host -ForegroundColor DarkBlue "		By default I turn on VLAN Compression"
Write-Host -ForegroundColor DarkBlue "		I set many policies with full range of options"
Write-Host -ForegroundColor DarkBlue "	You can always change the answer file or this script to adjust those parameters or names"
Write-Host ""
Write-Host -ForegroundColor DarkBlue "The DATA-Blank answer file is very well documented to help you setup your own answer files"
Write-Host -ForegroundColor DarkBlue "I have also included a sample answer file that you can run on the UCSM Emulator to see possible settings"
Write-Host -ForegroundColor DarkBlue "Your answer files must be named DATA-<something>.ps1 to be found by this script."
Write-Host ""
Write-Host -ForegroundColor DarkBlue "There is also a sample file:Custom UCS Settings v0.1 - BLANK.ps1 included."
Write-Host -ForegroundColor DarkBlue "	There is an option to run a custom script at the end of your UCS build to add any custom capabilities"
Write-Host -ForegroundColor DarkBlue "	This is an advanced feature required deep knowledge of PowerShell and PowerTool."
Write-Host -ForegroundColor DarkBlue "		An example of how I have used this is a customer who had 1Gb on their northbound LAN so I sent commands"
Write-Host -ForegroundColor DarkBlue "		to the uplinks to change the port speeds from 10Gb down to 1Gb"
Write-Host ""
Write-Host -ForegroundColor DarkBlue "Once the script has completed you should be ready to create Service Profiles from the created Service Profile Templates"
Write-Host ""

#Do not show errors in script
## Default is Continue.  Options are Inquire and Stop
#$ErrorActionPreference = "Stop"
#$ErrorActionPreference = "Inquire"
#$ErrorActionPreference = "Continue"
$ErrorActionPreference = "SilentlyContinue"

#Load the UCS PowerTool
Write-Host -ForegroundColor DarkBlue "Checking Cisco PowerTool"
$PowerToolLoaded = $null
$Modules = Get-Module
$PowerToolLoaded = $modules.name
if ( -not ($Modules -like "ciscoUcsPs"))
	{
		Write-Host -ForegroundColor DarkBlue "	Loading Module: Cisco UCS PowerTool Module"
		Import-Module ciscoUcsPs
		$Modules = Get-Module
		if ( -not ($Modules -like "ciscoUcsPs"))
			{
				Write-Host ""
				Write-Host -ForegroundColor Red "Cisco UCS PowerTool Module did not load.  Please correct his issue and try again"
				Write-Host -ForegroundColor Red "	Exiting..."
				exit
			}
		else
			{
				Write-Host -ForegroundColor DarkGreen "	PowerTool is Loaded"
			}
	}
else
	{
		Write-Host -ForegroundColor DarkGreen "	PowerTool is Loaded"
	}

#Validate correct version of PowerTool
$PowerTool = Get-UcsPowerToolConfiguration
if (([int]$PowerTool.Version.Major -ge 1) -and ([int]$PowerTool.Version.Minor -ge 0) -and ([int]$PowerTool.Version.Build -ge 0))
	{
		Write-Host ""
		Write-Host -ForegroundColor DarkBlue "PowerTool " $PowerTool.Version "is loaded" -NoNewline
		$PTUCSM = 2.1
	}
else
	{
		Write-Host ""
		Write-Host -ForegroundColor Red "Your version of PowerTool is too old for this script"
		Write-Host -ForegroundColor Red "	Exiting..."
		Disconnect-Ucs
		exit
	}
if ([int]$PowerTool.Version.Minor -eq 0)
	{
		Write-Host -ForegroundColor DarkBlue " for version 2.1 of UCSM"
	}
if ([int]$PowerTool.Version.Minor -eq 1)
	{
		Write-Host -ForegroundColor DarkBlue " for version 2.1 and 2.2 of UCSM"
	}

#Make sure I am disconnected from the UCS system to be configured
Disconnect-Ucs

#Get Data File
cd $PSScriptRoot
$ConfigFiles = dir "Data-*.ps1"

#Load Data File
if ($FILE)
	{
		if (Test-Path $FILE)
			{
				#Set the data configuration file
				$ConfigFile = $FILE
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
		#Create list of Config Files
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
		[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
		[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
		$Form = New-Object System.Windows.Forms.Form
		$Form.width = 700
		$Form.height = 150
		$Form.StartPosition = "CenterScreen"
		$Form.Text = ”Config File to use”
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
		$DropDownLabel.Text = "Config File"
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
		$ConfigFile = $DropDown.SelectedItem
		$ConfigFileDir = $CFs.Directory
		cd $ConfigFileDir
	}

#Load the data configuration file
Write-Host ""
Import-Module .\$ConfigFile -Force
	
#Validate data file
Write-Host ""
Write-Host -ForegroundColor DarkBlue "Validating Data File..."

$CharacterTest = [regex]"^[YyNn]"
if ($CharacterTest.Match($UCSEmulator).success)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $UCSEmulator'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Disconnect-Ucs
		exit
	}

#Test that UCSM is IP Reachable via Ping
$ping = new-object system.net.networkinformation.ping
$results = $ping.send($myucs)
if ($results.Status -ne "Success")
	{
		Write-host -ForegroundColor Red "	Can not access UCSM $myucs by Ping"
		Write-Host ""
		Write-Host -ForegroundColor DarkCyan "It is possible that a firewall is blocking ICMP (PING) Access.  Would you like to try to log in anyway?"
		$Try = Read-Host "Would you like to try to log in anyway? (Y/N)"
		if ($Try -ieq "y")
			{
				Write-Host ""
				Write-Host -ForegroundColor DarkBlue "Trying to log in anyway!"
				Write-Host ""
			}
		elseif ($Try -ieq "n")
			{
				Write-Host ""
				Write-Host -ForegroundColor DarkBlue "You have chosen to exit"
				Disconnect-Ucs
				exit
			}
		else
			{
				Write-Host ""
				Write-Host -ForegroundColor Red "You have provided invalid input"
				Write-Host ""
				Disconnect-Ucs
				exit
			}			
	}

#Validate data file
$CharacterTest = [regex]"^[0-9a-fA-F]*$"
if (($UCSDomain.Length -eq 2) -and (($CharacterTest.Match($UCSDomain).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $UCSDomain'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[YyNn]"
if ($CharacterTest.match($BootFromHD).Success)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $BootFromHD'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[YyNn]"
if ($CharacterTest.match($BootFromSAN).Success)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $BootFromSAN'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[YyNn]"
if ($CharacterTest.match($BootFromiSCSI).Success)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $BootFromiSCSI'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (($ChassisPower -ieq "n+1") -or ($ChassisPower -ieq "grid") -or ($ChassisPower -ieq "non-redundant"))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $ChassisPower'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

foreach ($SP in $ServerPort)
	{
		$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
		$NumberTest = [regex]"^[0-9]*$"
		if (($NumberTest.Match($SP["Port"]).Success) -and ($NumberTest.Match($SP["Slot"]).Success) -and ((($SP["LabelA"]).Length -le 16) -and (($CharacterTest.Match(($SP["LabelA"])).Success))) -and ((($SP["LabelB"]).Length -le 16) -and (($CharacterTest.Match(($SP["LabelB"])).Success))))
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $ServerPort Array'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}
	}		

foreach ($UP in $UplinkPort)
	{
		$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
		$NumberTest = [regex]"^[0-9]*$"
		if (($NumberTest.Match($UP["Port"]).Success) -and ($NumberTest.Match($UP["Slot"]).Success) -and ((($UP["LabelA"]).Length -le 16) -and (($CharacterTest.Match(($UP["LabelA"])).Success))) -and ((($UP["LabelB"]).Length -le 16) -and (($CharacterTest.Match(($UP["LabelB"])).Success))))
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $UplinkPort Array'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}
	}		


$CharacterTest = [regex]"^[YyNn]"
if ($CharacterTest.match($LANPortChannels).Success)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $LANPortChannels'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}
	
$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if (($LANPortChannelAName.Length -le 16) -and (($CharacterTest.Match($LANPortChannelAName).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $LANPortChannelAName'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ([int]$LANPortChannelANumber -le 256)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $LANPortChannelANumber'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if (($LANPortChannelBName.Length -le 16) -and (($CharacterTest.Match($LANPortChannelBName).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $LANPortChannelBName'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ([int]$LANPortChannelANumber -le 256)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $LANPortChannelBNumber'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}
	
foreach ($FC in $FCPort)
	{
		$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
		$NumberTest = [regex]"^[0-9]*$"
		if (($NumberTest.Match($FC["Port"]).Success) -and ($NumberTest.Match($FC["Slot"]).Success) -and ((($FC["LabelA"]).Length -le 16) -and (($CharacterTest.Match(($FC["LabelA"])).Success))) -and ((($FC["LabelB"]).Length -le 16) -and (($CharacterTest.Match(($FC["LabelB"])).Success))))
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $FCPort Array'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}
	}

$CharacterTest = [regex]"^[YyNn]"
if ($CharacterTest.match($SANPortChannels).Success)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $SANPortChannels'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}
	
$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if (($SANPortChannelAName.Length -le 16) -and (($CharacterTest.Match($SANPortChannelAName).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $SANPortChannelAName'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ($SANPortChannelANumber -le 256)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $LANPortChannelANumber'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if (($SANPortChannelBName.Length -le 16) -and (($CharacterTest.Match($SANPortChannelBName).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $SANPortChannelBName'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ($SANPortChannelANumber -le 256)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $SANPortChannelBNumber'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (([int]$VSANidA -le 4078) -or ($VSANidA -match "^40[8-9][0-3]") -or ($VSANidA -eq ""))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $VSANidA'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (([int]$FcoeVlanA -le 4093) -and ([int]$FcoeVlanA -ne 4048) -or ($FcoeVlanA -eq ""))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $FcoeVlanA'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (([int]$VSANidB -le 4078) -or ($VSANidB -match "^40[8-9][0-3]") -or ($VSANidB -eq ""))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $VSANidB'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (([int]$FcoeVlanB -le 4093) -and ([int]$FcoeVlanB -ne 4048) -or ($FcoeVlanB -eq ""))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $FcoeVlanB'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if ((($VHBAnameA.Length -le 16) -and (($CharacterTest.Match($VHBAnameA).Success))) -or ($VHBAnameA -eq ""))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $VHBAnameA'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if ((($VHBAnameB.Length -le 16) -and (($CharacterTest.Match($VHBAnameB).Success))) -or ($VHBAnameB -eq ""))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $VHBAnameB'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ($ArrayPort[0].Controller -ne "")
	{
		foreach ($AP in $ArrayPort)
			{
				$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
				$NumberTest = [regex]"^[0-9]*$"
				$ABtest = [regex]"^[a-zA-Z]*$"
				$WWPNtest = [regex]"^[a-fA-F0-9:]*$"
				if (($ABtest.Match($AP["Controller"]).Success) -and ($NumberTest.Match($AP["Port"]).Success) -and ($NumberTest.Match($AP["Count"]).Success) -and ($ABtest.Match($AP["Fabric"]).Success) -and ((($AP["Name"]).Length -le 16) -and (($CharacterTest.Match(($AP["Name"])).Success))) -and (($AP["WWPN"] -match "\w{2}\:\w{2}\:\w{2}\:\w{2}\:\w{2}\:\w{2}\:\w{2}\:\w{2}") -and (($WWPNtest.Match(($AP["WWPN"])).Success))))
					{
					}
				else
					{
						Write-Host -ForegroundColor Red 'Invalid entry: $ArrayPort Array'
						Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
						Write-Host -ForegroundColor Red "		Exiting..."
						Disconnect-Ucs
						exit
					}
			}
	}

#Test Boot Matrix $BootMatrix1 = @{Name = "FC1526";	APrimary = "1";	ASecondary = "5";	BPrimary = "2";	BSecondary = "6" }
if ($BootMatrix1[0].Name -ne "")
	{
		foreach ($BM in $BootMatrix)
			{
				$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
				$NumberTest = [regex]"^[0-9]*$"
				if (((($BM["Name"]).Length -le 6) -and (($CharacterTest.Match(($BM["Name"])).Success))) -and ($NumberTest.Match($BM["APrimary"]).Success) -and ($NumberTest.Match($BM["ASecondary"]).Success) -and ($NumberTest.Match($BM["BPrimary"]).Success) -and ($NumberTest.Match($BM["BSecondary"]).Success))
					{
					}
				else
					{
						Write-Host -ForegroundColor Red 'Invalid entry: $BootMatrix Array'
						Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
						Write-Host -ForegroundColor Red "		Exiting..."
						Disconnect-Ucs
						exit
					}
			}
	}

$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if (($iSCSINicNameA.Length -le 16) -and (($CharacterTest.Match($iSCSINicNameA).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $iSCSINicNameA'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if (($iSCSINicNameB.Length -le 16) -and (($CharacterTest.Match($iSCSINicNameB).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $iSCSINicNameB'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (($iSCSIPool -ieq "Pool") -or ($iSCSIPool -ieq "DHCP") -or ($iSCSIPool -ieq ""))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $iSCSIPool'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if (($DefaultiSCSIInitiatorPoolA.Length -le 32) -and (($CharacterTest.Match($DefaultiSCSIInitiatorPoolA).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultiSCSIInitiatorPoolA'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if (($DefaultiSCSIInitiatorPoolB.Length -le 32) -and (($CharacterTest.Match($DefaultiSCSIInitiatorPoolB).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultiSCSIInitiatorPoolB'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if (($iSCSIiqn.Length -le 256) -and (($CharacterTest.Match($iSCSIiqn).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $iSCSIiqn'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if (($IQNPrefix.Length -le 150) -and (($CharacterTest.Match($IQNPrefix).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $IQNPrefix'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
if (($IQNSuffix.Length -le 64) -and (($CharacterTest.Match($IQNSuffix).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $IQNSuffix'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (([int]$IQNFrom -le 1000) -and ([int]$IQNFrom -le [int]$IQNTo))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $IQNFrom'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (([int]$IQNTo -le 1000) -and ([int]$IQNTo -ge [int]$IQNFrom))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $IQNTo'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[YyNn]"
if (($CharacterTest.match($BronzeQoSEnabled).Success) -and ($CharacterTest.match($SilverQoSEnabled).Success) -and ($CharacterTest.match($GoldQoSEnabled).Success) -and ($CharacterTest.match($PlatinumQoSEnabled).Success))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: QoS Option - Enabled'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CoStest = [regex]"^[1-6]"
if (($CoStest.match($FibreChannelQoSCoS).Success) -and ($CoStest.match($BronzeQoSCoS).Success) -and ($CoStest.match($SilverQoSCoS).Success) -and ($CoStest.match($GoldQoSCoS).Success) -and ($CoStest.match($PlatinumQoSCoS).Success))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: QoS Option - CoS Value'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CoSList = @()
$CoSList += $FibreChannelQoSCoS
$CoSList += $BronzeQoSCoS
$CoSList += $SilverQoSCoS
$CoSList += $GoldQoSCoS
$CoSList += $PlatinumQoSCoS
$CoSunique = $CoSList | select –unique
if ($CoSList.Count -eq $CoSunique.Count)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: QoS Option - CoS Overlap'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ((($BronzeQoSPacketDrop -ieq "drop") -or ($BronzeQoSHostControl -ieq "no-drop")) -and (($SilverQoSPacketDrop -ieq "drop") -or ($SilverQoSPacketDrop -ieq "no-drop")) -and (($GoldQoSPacketDrop -ieq "drop") -or ($GoldQoSHostPacketDrop -ieq "no-drop")) -and (($PlatinumQoSPacketDrop -ieq "drop") -or ($PlatinumQoSPacketDrop -ieq "no-drop")))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: QoS Option - Packet Drop'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$DropList = @()
$DropList += $BronzeQoSPacketDrop
$DropList += $SilverQoSPacketDrop
$DropList += $GoldQoSPacketDrop
$DropList += $PlatinumQoSPacketDrop
$DropOverlap = $DropList | where {$_ -ieq "no-drop"}
if ($DropOverlap.Count -le 1)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: QoS Option - Packet Drop Overlap'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$QoSWeight = @("none", "best-effort", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
if (($QoSWeight | where {$_ -ieq $FibreChannelQoSWeight}) -and ($QoSWeight | where {$_ -ieq $BestEffortQoSWeight}) -and ($QoSWeight | where {$_ -ieq $BronzeQoSWeight}) -and ($QoSWeight | where {$_ -ieq $SilverQoSWeight}) -and ($QoSWeight | where {$_ -ieq $GoldQoSWeight}) -and ($QoSWeight | where {$_ -ieq $PlatinumQoSWeight}))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: QoS Option - Weight'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ((($FibreChannelQoSHostControl -ieq "none") -or ($FibreChannelQoSHostControl -ieq "full")) -and (($BestEffortQoSHostControl -ieq "none") -or ($BestEffortQoSHostControl -ieq "full")) -and (($BronzeQoSHostControl -ieq "none") -or ($BronzeQoSHostControl -ieq "full")) -and (($SilverQoSHostControl -ieq "none") -or ($SilverQoSHostControl -ieq "full")) -and (($GoldQoSHostControl -ieq "none") -or ($GoldQoSHostControl -ieq "full")) -and (($PlatinumQoSHostControl -ieq "none") -or ($PlatinumQoSHostControl -ieq "full")))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: QoS Option - Host Control'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (($BestEffortQoSMTU -ieq "normal") -or ($BestEffortQoSMTU -ieq "normal") -or (([int]$BestEffortQoSMTU -ge 1500) -and ([int]$BestEffortQoSMTU -le 9216)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: QoS Option - MTU'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ((($BestEffortQoSMulticastOptimized -ieq "yes") -or ($BestEffortQoSMulticastOptimized -ieq "no")) -and (($BronzeQoSMulticastOptimized -ieq "yes") -or ($BronzeQoSMulticastOptimized -ieq "no")) -and (($SilverQoSMulticastOptimized -ieq "yes") -or ($SilverQoSMulticastOptimized -ieq "no")) -and (($GoldQoSMulticastOptimized -ieq "yes") -or ($GoldQoSMulticastOptimized -ieq "no")) -and (($PlatinumQoSMulticastOptimized -ieq "yes") -or ($PlatinumQoSMulticastOptimized -ieq "no")))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: QoS Option - Multicast Optimized'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$MCOList = @()
$MCOList += $BestEffortQoSMulticastOptimized
$MCOList += $BronzeQoSMulticastOptimized
$MCOList += $SilverQoSMulticastOptimized
$MCOList += $GoldQoSMulticastOptimized
$MCOList += $PlatinumQoSMulticastOptimized
$MCOoverlap = $MCOList | where {$_ -ieq "yes"}
if ($MCOoverlap.Count -le 1)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: QoS Option - Multicast Optimized Overlap'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ((([int]$FibreChannelQoSBurst -ge 1) -and ([int]$FibreChannelQoSBurst -le 65535)) -and (([int]$BestEffortQoSBurst -ge 1) -and ([int]$BestEffortQoSBurst -le 65535)) -and (([int]$BronzeQoSBurst -ge 1) -and ([int]$BronzeQoSBurst -le 65535)) -and (([int]$SilverQoSBurst -ge 1) -and ([int]$SilverQoSBurst -le 65535)) -and (([int]$GoldQoSBurst -ge 1) -and ([int]$GoldQoSBurst -le 65535)) -and (([int]$PlatinumQoSBurst -ge 1) -and ([int]$PlatinumQoSBurst -le 65535)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: QoS Option - Burst'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ((($FibreChannelQoSRate -ge "1") -and ($FibreChannelQoSRate -le "9999999") -or ($FibreChannelQoSRate -ieq "line-rate")) -and (($BestEffortQoSRate -ge "1") -and ($BestEffortQoSRate -le "9999999") -or ($BestEffortQoSRate -ieq "line-rate")) -and (($BronzeQoSRate -ge "1") -and ($BronzeQoSRate -le "9999999") -or ($BronzeQoSRate -ieq "line-rate")) -and (($SilverQoSRate -ge "1") -and ($SilverQoSRate -le "9999999") -or ($SilverQoSRate -ieq "line-rate")) -and (($GoldQoSRate -ge "1") -and ($GoldQoSRate -le "9999999") -or ($GoldQoSRate -ieq "line-rate")) -and (($PlatinumQoSRate -ge "1") -and ($PlatinumQoSRate -le "9999999") -or ($PlatinumQoSRate -ieq "line-rate")))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: QoS Option - Rate'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

################################################################
# Figure out how to validate VLANS
################################################################

foreach ($NW in $network)
	{
		$CharacterTest = [regex]"^[A-Za-z0-9_:.-]*$"
		if ((($NW["vlanname"]).Length -le 32) -and (($CharacterTest.Match($NW["vlanname"]).Success)))
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $network["vlanname"]'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}
			
		if (([int]$NW["vlannumber"] -le 3967) -or ($NW["vlannumber"] -match "^404[8-9]") -or ($NW["vlannumber"] -match "^40[5-8][0-9]") -or ($NW["vlannumber"] -match "^409[0-3]"))
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $network["vlannumber"]'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}

		$CharacterTest = [regex]"^[0-9a-fA-F]*$"
		if ((($nw["macid"]).Length -eq 2) -and (($CharacterTest.Match($nw["macid"]).Success)) -or ($nw["macid"] -ieq ""))
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $network["macid"]'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}

		if (([int]$nw["mtu"] -ge 1500) -and ([int]$nw["mtu"] -le 9000) -or ($nw["mtu"] -eq ""))
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $network["mtu"]'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}
		
		if (($NW["fabric"] -ieq "A-B") -or ($NW["fabric"] -ieq "B-A") -or ($NW["fabric"] -ieq "A") -or ($NW["fabric"] -ieq "B") -or ($NW["fabric"] -ieq "NONE"))
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $network["fabric"]'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}
			
		if (($NW["QoSPolicy"] -ieq "BestEffort") -or ($NW["QoSPolicy"] -ieq "Bronze") -or ($NW["QoSPolicy"] -ieq "Silver") -or ($NW["QoSPolicy"] -ieq "Gold") -or ($NW["QoSPolicy"] -ieq "Platinum") -or ($NW["QoSPolicy"] -ieq $null) -or ($NW["QoSPolicy"] -ieq ""))
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $network["QoSPolicy"]'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}

		$CharacterTest = [regex]"^[A-za-z0-9_:.-]*$"
		if ((($NW["trunknic"]).Length -le 32) -and (($CharacterTest.Match($NW["trunknic"]).Success)))
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $network["trunknic"]'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}
			
		if (($NW["defaultvlan"] -ieq "y") -or ($NW["defaultvlan"] -ieq "n") -or ($NW["defaultvlan"] -ieq $null) -or ($NW["defaultvlan"] -ieq ""))
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $network["defaultvlan"]'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}
	}

$CharacterTest = [regex]"^[A-za-z0-9_:.-]*$"
if (($NativeVLANname.Length -le 32) -and (($CharacterTest.Match($NativeVLANname).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $NativeVLANname'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (([int]$NativeVLANnumber -le 3967) -or ($NativeVLANnumber -match "^404[8-9]") -or ($NativeVLANnumber -match "^40[5-8][0-9]") -or ($NativeVLANnumber -match "^409[0-3]"))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $NativeVLANnumber'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-za-z0-9_:.-]*$"
if (($PXENIC.Length -le 16) -and (($CharacterTest.Match($PXENIC).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $PXENIC'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-za-z0-9_:.-]*$"
if (($IPMIpassword.Length -le 20) -and (($CharacterTest.Match($IPMIpassword).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $IPMIpassword'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

try 
	{
		$Address = [System.Net.IPAddress]::parse($QuerierIpAddr)
		$Valid = $True -f $Address.IPaddressToString
	}
catch 
	{
		$Valid = $False -f $QuerierIpAddr
	}
if ($valid -eq $true)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $QuerierIpAddr'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

try 
	{
		$Address = [System.Net.IPAddress]::parse($DefGw)
		$Valid = $True -f $Address.IPaddressToString
	}
catch 
	{
		$Valid = $False -f $DefGw
	}
if ($valid -eq $true)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefGw'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ($PriDNS -ne "")
	{
		try 
			{
				$Address = [System.Net.IPAddress]::parse($PriDNS)
				$Valid = $True -f $Address.IPaddressToString
			}
		catch 
			{
				$Valid = $False -f $PriDNS
			}
		if ($valid -eq $true)
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $PriDNS'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}
	}

if ($SecDNS -ne "")
	{
		try 
			{
				$Address = [System.Net.IPAddress]::parse($SecDNS)
				$Valid = $True -f $Address.IPaddressToString
			}
		catch 
			{
				$Valid = $False -f $SecDNS
			}
		if ($valid -eq $true)
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $SecDNS'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}
	}

try 
	{
		$Address = [System.Net.IPAddress]::parse($MgmtIPstart)
		$Valid = $True -f $Address.IPaddressToString
	}
catch 
	{
		$Valid = $False -f $MgmtIPstart
	}
if ($valid -eq $true)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $MgmtIPstart'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

try 
	{
		$Address = [System.Net.IPAddress]::parse($MgmtIPend)
		$Valid = $True -f $Address.IPaddressToString
	}
catch 
	{
		$Valid = $False -f $MgmtIPend
	}
if ($valid -eq $true)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $MgmtIPend'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ($UCSDesc.Length -le 256)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $UCSDesc'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ($UCSOwner.Length -le 32)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $UCSOwner'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ($UCSSite.Length -le 32)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $UCSSite'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-za-z0-9_.-]*$"
if (($UCSDNSDomain.Length -le 256) -and (($CharacterTest.Match($UCSDNSDomain).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $UCSDNSDomain'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-za-z0-9_:.-]*$"
if (($SystemName.Length -le 30) -and (($CharacterTest.Match($SystemName).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $SystemName'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

#TimeZone
foreach ($NTP in $NTPName)
	{
		if ($NTP.length -le 64)
			{
			}
		else
			{
				Write-Host -ForegroundColor Red 'Invalid entry: $NTPName'
				Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "		Exiting..."
				Disconnect-Ucs
				exit
			}
	}

$ScrubOptions = @("No_Scrub", "BIOS_Scrub", "Disk_Scrub", "Full_Scrub")
if (($ScrubOptions | where {$_ -ieq $DefaultScrub}) -ieq $DefaultScrub)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $ScrubOptions'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (($DefaultRackServerDiscovery -ieq "user-acknowledged") -or ($DefaultRackServerDiscovery -ieq "immediate"))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultRackServerDiscovery'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (($DefaultRackManagement -ieq "user-acknowledged") -or ($DefaultRackManagement -ieq "auto-acknowledged"))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultRackManagement'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (($DefaultDiscoveryAction -ieq "1-link") -or ($DefaultDiscoveryAction -ieq "2-link") -or ($DefaultDiscoveryAction -ieq "4-link") -or ($DefaultDiscoveryAction -ieq "8-link") -or ($DefaultDiscoveryAction -ieq "platform-max"))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultDiscoveryAction'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (($DefaultLinkGrouping -ieq "none") -or ($DefaultLinkGrouping -ieq "port-channel"))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultLinkGrouping'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (($DefaultPowerControl -ieq "default") -or ($DefaultPowerControl -ieq "No_Cap"))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultPowerControl'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (($DefaultSoL -ieq "No_SoL") -or ($DefaultSoL -ieq "SoL_9600") -or ($DefaultSoL -ieq "SoL_19200") -or ($DefaultSoL -ieq "SoL_38400") -or ($DefaultSoL -ieq "SoL_57600") -or ($DefaultSoL -ieq "SoL_115200"))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultSoL'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-za-z0-9_:.-]*$"
if (($DefaultLANConnectivity.Length -le 16) -and (($CharacterTest.Match($DefaultLANConnectivity).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultLANConnectivity'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-za-z0-9_:.-]*$"
if (($DefaultLANwiSCSIConnectivity.Length -le 16) -and (($CharacterTest.Match($DefaultLANwiSCSIConnectivity).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultLANwiSCSIConnectivity'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if (($DefaultLANAdapter -ieq "") -or ($DefaultLANAdapter -ieq "Linux") -or ($DefaultLANAdapter -ieq "SRIOV") -or ($DefaultLANAdapter -ieq "Solaris") -or ($DefaultLANAdapter -ieq "VMWare") -or ($DefaultLANAdapter -ieq "VMWarePassThru") -or ($DefaultLANAdapter -ieq "Windows") -or ($DefaultLANAdapter -ieq "default"))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultLANAdapter'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-za-z0-9_:.-]*$"
if (($DefaultvSANName.Length -le 32) -and (($CharacterTest.Match($DefaultvSANName).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultvSANName'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-za-z0-9_:.-]*$"
if (($DefaultHBAConnectivity.Length -le 16) -and (($CharacterTest.Match($DefaultHBAConnectivity).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultHBAConnectivity'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CharacterTest = [regex]"^[A-za-z0-9_:.-]*$"
if (($DefaultWWPNPool.Length -le 30) -and (($CharacterTest.Match($DefaultWWPNPool).Success)))
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $DefaultWWPNPool'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CheckUUID = $UUIDfrom.Split("-")
$CharacterTest = [regex]"^[A-Fa-f0-9]*$"
if (($CheckUUID[0].Length -eq 4) -and (($CharacterTest.Match($CheckUUID[0]).Success)) -and ($CheckUUID[1].Length -eq 12) -and (($CharacterTest.Match($CheckUUID[1]).Success)) -and ($UUIDfrom.Substring(4,1) -eq "-") -and $UUIDfrom.Length -eq 17)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $UUIDfrom'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CheckUUID = $UUIDto.Split("-")
$CharacterTest = [regex]"^[A-Fa-f0-9]*$"
if (($CheckUUID[0].Length -eq 4) -and (($CharacterTest.Match($CheckUUID[0]).Success)) -and ($CheckUUID[1].Length -eq 12) -and (($CharacterTest.Match($CheckUUID[1]).Success)) -and ($UUIDto.Substring(4,1) -eq "-") -and $UUIDto.Length -eq 17)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $UUIDto'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CheckMAC = $MACfrom.Split(":")
$CharacterTest = [regex]"^[A-Fa-f0-9IDidIdiD]*$"
if (($CheckMAC[0].Length -eq 2) -and (($CharacterTest.Match($CheckMAC[0]).Success)) -and ($CheckMAC[1].Length -eq 2) -and (($CharacterTest.Match($CheckMAC[1]).Success)) -and ($CheckMAC[2].Length -eq 2) -and (($CharacterTest.Match($CheckMAC[2]).Success)) -and ($CheckMAC[3].Length -eq 2) -and (($CharacterTest.Match($CheckMAC[3]).Success)) -and ($CheckMAC[4].Length -eq 2) -and (($CharacterTest.Match($CheckMAC[4]).Success)) -and ($CheckMAC[5].Length -eq 2) -and (($CharacterTest.Match($CheckMAC[5]).Success)) -and $MACfrom.Length -eq 17)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $MACfrom'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CheckMAC = $MACto.Split(":")
$CharacterTest = [regex]"^[A-Fa-f0-9IDidIdiD]*$"
if (($CheckMAC[0].Length -eq 2) -and (($CharacterTest.Match($CheckMAC[0]).Success)) -and ($CheckMAC[1].Length -eq 2) -and (($CharacterTest.Match($CheckMAC[1]).Success)) -and ($CheckMAC[2].Length -eq 2) -and (($CharacterTest.Match($CheckMAC[2]).Success)) -and ($CheckMAC[3].Length -eq 2) -and (($CharacterTest.Match($CheckMAC[3]).Success)) -and ($CheckMAC[4].Length -eq 2) -and (($CharacterTest.Match($CheckMAC[4]).Success)) -and ($CheckMAC[5].Length -eq 2) -and (($CharacterTest.Match($CheckMAC[5]).Success)) -and $MACto.Length -eq 17)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $MACto'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CheckWWN = $WWNNfrom.Split(":")
$CharacterTest = [regex]"^[A-Fa-f0-9]*$"
if (($CheckWWN[0].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[0]).Success)) -and ($CheckWWN[1].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[1]).Success)) -and ($CheckWWN[2].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[2]).Success)) -and ($CheckWWN[3].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[3]).Success)) -and ($CheckWWN[4].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[4]).Success)) -and ($CheckWWN[5].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[5]).Success)) -and ($CheckWWN[6].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[6]).Success)) -and ($CheckWWN[7].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[7]).Success)) -and $WWNNfrom.Length -eq 23)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $WWNNfrom'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}
	
$CheckWWN = $WWNNto.Split(":")
$CharacterTest = [regex]"^[A-Fa-f0-9]*$"
if (($CheckWWN[0].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[0]).Success)) -and ($CheckWWN[1].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[1]).Success)) -and ($CheckWWN[2].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[2]).Success)) -and ($CheckWWN[3].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[3]).Success)) -and ($CheckWWN[4].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[4]).Success)) -and ($CheckWWN[5].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[5]).Success)) -and ($CheckWWN[6].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[6]).Success)) -and ($CheckWWN[7].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[7]).Success)) -and $WWNNto.Length -eq 23)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $WWNNto'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CheckWWN = $WWPNaFrom.Split(":")
$CharacterTest = [regex]"^[A-Fa-f0-9]*$"
if (($CheckWWN[0].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[0]).Success)) -and ($CheckWWN[1].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[1]).Success)) -and ($CheckWWN[2].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[2]).Success)) -and ($CheckWWN[3].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[3]).Success)) -and ($CheckWWN[4].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[4]).Success)) -and ($CheckWWN[5].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[5]).Success)) -and ($CheckWWN[6].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[6]).Success)) -and ($CheckWWN[7].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[7]).Success)) -and $WWPNaFrom.Length -eq 23)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $WWPNaFrom'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CheckWWN = $WWPNaTo.Split(":")
$CharacterTest = [regex]"^[A-Fa-f0-9]*$"
if (($CheckWWN[0].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[0]).Success)) -and ($CheckWWN[1].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[1]).Success)) -and ($CheckWWN[2].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[2]).Success)) -and ($CheckWWN[3].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[3]).Success)) -and ($CheckWWN[4].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[4]).Success)) -and ($CheckWWN[5].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[5]).Success)) -and ($CheckWWN[6].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[6]).Success)) -and ($CheckWWN[7].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[7]).Success)) -and $WWPNaTo.Length -eq 23)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $WWPNaTo'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CheckWWN = $WWPNbFrom.Split(":")
$CharacterTest = [regex]"^[A-Fa-f0-9]*$"
if (($CheckWWN[0].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[0]).Success)) -and ($CheckWWN[1].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[1]).Success)) -and ($CheckWWN[2].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[2]).Success)) -and ($CheckWWN[3].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[3]).Success)) -and ($CheckWWN[4].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[4]).Success)) -and ($CheckWWN[5].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[5]).Success)) -and ($CheckWWN[6].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[6]).Success)) -and ($CheckWWN[7].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[7]).Success)) -and $WWPNbFrom.Length -eq 23)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $WWPNbFrom'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

$CheckWWN = $WWPNbTo.Split(":")
$CharacterTest = [regex]"^[A-Fa-f0-9]*$"
if (($CheckWWN[0].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[0]).Success)) -and ($CheckWWN[1].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[1]).Success)) -and ($CheckWWN[2].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[2]).Success)) -and ($CheckWWN[3].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[3]).Success)) -and ($CheckWWN[4].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[4]).Success)) -and ($CheckWWN[5].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[5]).Success)) -and ($CheckWWN[6].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[6]).Success)) -and ($CheckWWN[7].Length -eq 2) -and (($CharacterTest.Match($CheckWWN[7]).Success)) -and $WWPNbTo.Length -eq 23)
	{
	}
else
	{
		Write-Host -ForegroundColor Red 'Invalid entry: $WWPNbTo'
		Write-Host -ForegroundColor Red "	Please correct this error in the data file and try again"
		Write-Host -ForegroundColor Red "		Exiting..."
		Disconnect-Ucs
		exit
	}

if ($CustomScript -ne $null)
	{
		cd $PSScriptRoot
		$TestCustomScript = Test-Path $CustomScript
		if ($TestCustomScript)
			{
			}
		else
			{
				Write-Host -ForegroundColor Red "	Script:"$CustomScript" does not exist in:"$PSScriptRoot
				Write-Host -ForegroundColor Red "		Please correct this error in the data file and try again"
				Write-Host -ForegroundColor Red "			Exiting..."
				Disconnect-Ucs
				exit
			}
	}

if (($DefaultiSCSIPoolSubnet -eq $null) -or ($DefaultiSCSIPoolSubnet -eq ""))
	{
		$DefaultiSCSIPoolSubnet = "255.255.255.0"
	}

if (($iSCSISubnetA -eq $null) -or ($iSCSISubnetA -eq ""))
	{
		$iSCSISubnetA = "255.255.255.0"
	}

if (($iSCSISubnetB -eq $null) -or ($iSCSISubnetB -eq ""))
	{
		$iSCSISubnetB = "255.255.255.0"
	}

Write-Host -ForegroundColor DarkGreen "	Data file validated"

#Log into the UCS System
$multilogin = Set-UcsPowerToolConfiguration -SupportMultipleDefaultUcs $false
Write-Host ""
Write-Host -ForegroundColor DarkBlue "Logging into UCS"
Write-Host -ForegroundColor DarkCyan "	Enter your UCSM credentials"

#Verify PowerShell Version to pick prompt type
$PSVersion = $psversiontable.psversion
$PSMinimum = $PSVersion.Major
if (!$UCREDENTIALS)
	{
		Write-Output "	Provide UCSM login credentials"
		if ($PSMinimum -ge "3")
			{
				$cred = Get-Credential -Message "UCSM Login Credentials" -UserName "admin"
			}
		else
			{
				$cred = Get-Credential
			}
	}
$myCon = Connect-Ucs $myucs -Credential $cred
if (($myucs | Measure-Object).count -ne ($myCon | Measure-Object).count) 
	{
		#Exit Script
		Write-Host -ForegroundColor Red "     Error Logging into UCS....Exiting"
		Disconnect-Ucs
		exit
	}
else
	{
		Write-Host -ForegroundColor DarkGreen "     Login Successful"
	}

#Make sure the user is positive that this is the system they want to modify/build
$UCSMajorVersion = $myCon.Version.Major
$UCSMinorVersion = $myCon.Version.Minor
$TopSystem=Get-UcsTopSystem -Address $myucs
$CurrentSystemName = $TopSystem.Name
$CurrentSystemIP = $TopSystem.Address
$CurrentSystemDesc = $TopSystem.Descr
$CurrentSystemOwner = $TopSystem.Owner
$CurrentSystemSite = $TopSystem.Site
Write-Host ""
Write-Host -ForegroundColor DarkBlue "You are logged into a UCS system with:"
Write-Host -ForegroundColor DarkBlue "	IP/Host:		" -NoNewline
Write-Host -ForegroundColor White -BackgroundColor DarkBlue $CurrentSystemIP
Write-Host -ForegroundColor DarkBlue "	Name: 		" -NoNewline
Write-Host -ForegroundColor White -BackgroundColor DarkBlue $CurrentSystemName
Write-Host -ForegroundColor DarkBlue "	Site: 		" -NoNewline
Write-Host -ForegroundColor White -BackgroundColor DarkBlue $CurrentSystemSite
Write-Host -ForegroundColor DarkBlue "	Owner: 		" -NoNewline
Write-Host -ForegroundColor White -BackgroundColor DarkBlue $CurrentSystemOwner
Write-Host -ForegroundColor DarkBlue "	Description: 	" -NoNewline
Write-Host -ForegroundColor White -BackgroundColor DarkBlue $CurrentSystemDesc
Write-Host ""
if (!$FILE)
	{
		Write-Host -ForegroundColor DarkCyan "Are you sure you want to continue (Y/N)"
		$AreYouSure = $null
		$AreYouSure = Read-Host "Are you sure you want to continue (Y/N)"
		if ($AreYouSure -ine "Y")
			{
				Write-Host -ForegroundColor DarkBlue "You have choosen to exit"
				Write-Host -ForegroundColor DarkBlue "	Exiting..."
				Disconnect-Ucs
				exit
			}
	}

#Test for single FI
$IsRedundantFI = Get-UcsFiModule | where {$_.Dn -eq "sys/switch-B/slot-1"}
	
#remove default pools - From real systems and the emulator
Write-Host ""
Write-Host -ForegroundColor DarkBlue "Removing UCS defaults"
Get-UcsOrg -Level root | Get-UcsServerPool -Name "blade-pool-2" -LimitScope | Remove-UcsServerPool -force
Get-UcsOrg -Level root | Get-UcsServerPool -Name "default" -LimitScope | Remove-UcsServerPool -force

Get-UcsOrg -Level root | Get-UcsIpPool -Name "ip-pool-1" -LimitScope | Remove-UcsIpPool -force

Get-UcsOrg -Level root | Get-UcsIqnPoolPool -Name "default" -LimitScope | Remove-UcsIqnPoolPool -force

Get-UcsOrg -Level root | Get-UcsUuidSuffixPool -Name "default" -LimitScope | Remove-UcsUuidSuffixPool -force

Get-UcsOrg -Level root | Get-UcsMacPool -Name "mac-pool-1" -LimitScope | Remove-UcsMacPool -force
Get-UcsOrg -Level root | Get-UcsMacPool -Name "default" -LimitScope | Remove-UcsMacPool -force

Get-UcsOrg -Level root | Get-UcsIqnPoolPool -Name "default" -LimitScope | Remove-UcsIqnPoolPool

Get-UcsOrg -Level root | Get-UcsWwnPool -Name "node-default" -LimitScope | Remove-UcsWwnPool -force
Get-UcsOrg -Level root | Get-UcsWwnPool -Name "default" -LimitScope | Remove-UcsWwnPool -force

Get-UcsOrg -Level root | Get-UcsLocalDiskConfigPolicy -Name "default" -LimitScope | Remove-UcsLocalDiskConfigPolicy -force

Get-UcsOrg -Level root | Get-UcsScrubPolicy -Name "default" -LimitScope | Remove-UcsScrubPolicy -force

Get-UcsOrg -Level root | Get-UcsServerPoolQualification -Name "all-chassis" -LimitScope | Remove-UcsServerPoolQualification  -force

Get-UcsOrg -Level root | Get-UcsFabricMulticastPolicy -Name "default" -LimitScope | Remove-UcsFabricMulticastPolicy -force

Get-UcsOrg -Level root | Get-UcsQosPolicy -Name "qos-1" -LimitScope | Remove-UcsQosPolicy -force

#Servers - Remove Sub-Organizations - From Emulator
Get-UcsOrg -Level root | Get-UcsOrg -Name "Finance" -LimitScope | Remove-UcsOrg -Force

#Admin - Remove default DNS Entry - From Emulator
Start-UcsTransaction
$mo = Get-UcsDns | Set-UcsDns -AdminState "enabled" -Descr "" -Domain "localdomain" -PolicyOwner "local" -Port 0 -Force
$mo_1 = Get-UcsDnsServer -Name "172.16.104.2" | Remove-UcsDnsServer -Force
Complete-UcsTransaction

#Set Equipment Policies
Start-UcsTransaction
Get-UcsOrg -Level root | Get-UcsRackServerDiscPolicy | Set-UcsRackServerDiscPolicy -Action $DefaultRackServerDiscovery -Descr "" -Name "default" -PolicyOwner "local" -Qualifier "" -ScrubPolicyName $DefaultScrub -Force
Get-UcsOrg -Level root | Get-UcsChassisDiscoveryPolicy | Set-UcsChassisDiscoveryPolicy -Action $DefaultDiscoveryAction -Descr "" -LinkAggregationPref $DefaultLinkGrouping -Name "" -PolicyOwner "local" -Rebalance "user-acknowledged" -Force
Get-UcsOrg -Level root | Get-UcsComputeServerMgmtPolicy | Set-UcsComputeServerMgmtPolicy -Action $DefaultRackManagement -Descr "" -Name "default" -PolicyOwner "local" -Qualifier "" -Force
if ($ChassisPower -ieq "grid")
	{
		Get-UcsOrg -Level root | Get-UcsPowerControlPolicy | Set-UcsPowerControlPolicy -Descr "" -PolicyOwner "local" -Redundancy "grid" -Force
	}
else
	{
		Get-UcsOrg -Level root | Get-UcsPowerControlPolicy | Set-UcsPowerControlPolicy -Descr "" -PolicyOwner "local" -Redundancy "n+1" -Force
	}
Complete-UcsTransaction

#Configure Server Ports
Write-Host -ForegroundColor DarkBlue "Configuring Server Ports"
if ($UCSEmulator -ieq "y")
	{
		Write-Host -ForegroundColor White -BackgroundColor DarkBlue "Since this is a UCS Emulator, we are skipping the configuration of server ports"
		sleep 5
	}
else
	{
		foreach ($sp in $ServerPort)
			{
				Start-UcsTransaction
				Get-UcsFabricServerCloud -Id "A" | Add-UcsServerPort -AdminState "enabled" -Name "" -PortId $sp["Port"] -SlotId $sp["Slot"] -UsrLbl $sp["LabelA"]
				Get-UcsFabricServerCloud -Id "B" | Add-UcsServerPort -AdminState "enabled" -Name "" -PortId $sp["Port"] -SlotId $sp["Slot"] -UsrLbl $sp["LabelB"]
				Complete-UcsTransaction
			}
	}
#Configure Uplink Ports
Write-Host -ForegroundColor DarkBlue "Configuring Uplink Ports"
foreach ($up in $UplinkPort)
	{
		Start-UcsTransaction
		Get-UcsFiLanCloud -Id "A" | Add-UcsUplinkPort -AdminSpeed "10gbps" -AdminState "enabled" -FlowCtrlPolicy "default" -Name "" -PortId $up["Port"] -SlotId $up["Slot"] -UsrLbl $up["LabelA"]
		Get-UcsFiLanCloud -Id "B" | Add-UcsUplinkPort -AdminSpeed "10gbps" -AdminState "enabled" -FlowCtrlPolicy "default" -Name "" -PortId $up["Port"] -SlotId $up["Slot"] -UsrLbl $up["LabelB"]
		Complete-UcsTransaction
	}

#Set Global QoS Settings
Write-Host -ForegroundColor DarkBlue "Setting Global QoS settings"
if (($FibreChannelQoSCoS -ine "3") -or ($FibreChannelQoSWeight -ine "5"))
	{
		Start-UcsTransaction
		$mo = Get-UcsQosclassDefinition | Set-UcsManagedObject -PropertyMap @{Descr=""; PolicyOwner="local"; } -Force
		$mo_1 = Get-UcsFcQosClass | Set-UcsFcQosClass -Cos $FibreChannelQoSCoS -Name "" -Weight $FibreChannelQoSWeight -Force
		Complete-UcsTransaction
	}
if (($BestEffortQoSCoS -ine "5") -or ($BestEffortQoSMTU -ne "normal") -or ($BestEffortQoSMulticastOptimized -ine "no") -or ($BestEffortQoSWeight -ine "5"))
	{
		Start-UcsTransaction
		$mo = Get-UcsQosclassDefinition | Set-UcsManagedObject -PropertyMap @{Descr=""; PolicyOwner="local"; } -Force
		$mo_1 = Get-UcsBestEffortQosClass | Set-UcsBestEffortQosClass -Mtu $BestEffortQoSMTU -MulticastOptimize $BestEffortQoSMulticastOptimized -Name "" -Weight $BestEffortQoSWeight -Force
		Complete-UcsTransaction
	}
if ($BronzeQoSEnabled -ieq "y")
	{
		Start-UcsTransaction
		$mo = Get-UcsQosclassDefinition | Set-UcsManagedObject -PropertyMap @{Descr=""; PolicyOwner="local"; } -Force
		$mo_1 = Get-UcsQosClass -Priority "bronze" | Set-UcsQosClass -AdminState "enabled" -Cos $BronzeQoSCoS -Drop $BronzeQoSPacketDrop -Mtu $BronzeQoSMTU -MulticastOptimize $BronzeQoSMulticastOptimized -Name "" -Weight $BronzeQoSWeight -Force
		Complete-UcsTransaction
	}
if ($SilverQoSEnabled -ieq "y")
	{
		Start-UcsTransaction
		$mo = Get-UcsQosclassDefinition | Set-UcsManagedObject -PropertyMap @{Descr=""; PolicyOwner="local"; } -Force
		$mo_1 = Get-UcsQosClass -Priority "silver" | Set-UcsQosClass -AdminState "enabled" -Cos $SilverQoSCoS -Drop $SilverQoSPacketDrop -Mtu $SilverQoSMTU -MulticastOptimize $SilverQoSMulticastOptimized -Name "" -Weight $SilverQoSWeight -Force
		Complete-UcsTransaction
	}
if ($GoldQoSEnabled -ieq "y")
	{
		Start-UcsTransaction
		$mo = Get-UcsQosclassDefinition | Set-UcsManagedObject -PropertyMap @{Descr=""; PolicyOwner="local"; } -Force
		$mo_1 = Get-UcsQosClass -Priority "gold" | Set-UcsQosClass -AdminState "enabled" -Cos $GoldQoSCoS -Drop $GoldQoSPacketDrop -Mtu $GoldQoSMTU -MulticastOptimize $GoldQoSMulticastOptimized -Name "" -Weight $GoldQoSWeight -Force
		Complete-UcsTransaction
	}
if ($PlatinumQoSEnabled -ieq "y")
	{
		Start-UcsTransaction
		$mo = Get-UcsQosclassDefinition | Set-UcsManagedObject -PropertyMap @{Descr=""; PolicyOwner="local"; } -Force
		$mo_1 = Get-UcsQosClass -Priority "platinum" | Set-UcsQosClass -AdminState "enabled" -Cos $PlatinumQoSCoS -Drop $PlatinumQoSPacketDrop -Mtu $PlatinumQoSMTU -MulticastOptimize $PlatinumQoSMulticastOptimized -Name "" -Weight $PlatinumQoSWeight -Force
		Complete-UcsTransaction
	}

#Servers - Create BIOS Policy
Write-Host -ForegroundColor DarkBlue "Setting BIOS Policies"
Write-Host -ForegroundColor DarkBlue "	Basic Policy"
Start-UcsTransaction
$mo = Get-UcsOrg -Level root  | Add-UcsBiosPolicy -Descr "" -Name "Basic" -PolicyOwner "local" -RebootOnUpdate "no"
$mo_1 = $mo | Set-UcsBiosVfQuietBoot -VpQuietBoot "disabled" -Force
Complete-UcsTransaction
Write-Host -ForegroundColor DarkBlue "	Performance Policy"
Start-UcsTransaction
$mo = Get-UcsOrg -Level root  | Add-UcsBiosPolicy -Descr "" -Name "Performance" -PolicyOwner "local" -RebootOnUpdate "no"
$mo_1 = $mo | Set-UcsBiosVfCPUPerformance -VpCPUPerformance "hpc" -Force
$mo_2 = $mo | Add-UcsManagedObject -XmlTag biosVfDramRefreshRate -ModifyPresent -PropertyMap @{VpDramRefreshRate="4x"; Dn="org-root/bios-prof-Performance/Dram-Refresh-Rate"; }
$mo_3 = $mo | Set-UcsBiosEnhancedIntelSpeedStep -VpEnhancedIntelSpeedStepTech "disabled" -Force
$mo_4 = $mo | Set-UcsBiosLvDdrMode -VpLvDDRMode "performance-mode" -Force
$mo_5 = $mo | Set-UcsBiosVfProcessorCState -VpProcessorCState "disabled" -Force
$mo_6 = $mo | Set-UcsBiosVfProcessorC1E -VpProcessorC1E "disabled" -Force
$mo_7 = $mo | Set-UcsBiosVfProcessorC3Report -VpProcessorC3Report "disabled" -Force
$mo_8 = $mo | Set-UcsBiosVfProcessorC6Report -VpProcessorC6Report "disabled" -Force
$mo_9 = $mo | Set-UcsBiosVfProcessorC7Report -VpProcessorC7Report "disabled" -Force
$mo_10 = $mo | Set-UcsBiosVfQuietBoot -VpQuietBoot "disabled" -Force
$mo_11 = $mo | Set-UcsBiosVfSelectMemoryRASConfiguration -VpSelectMemoryRASConfiguration "maximum-performance" -Force
Complete-UcsTransaction

if (([int]$UCSMajorVersion -ge 2) -and ([int]$UCSMinorVersion -ge 2))
	{
		Write-Host -ForegroundColor DarkBlue "Configuring DIMM Blacklisting"
		Write-Host ""
		#Enable DIMM Blacklisting
		Get-UcsManagedObject -Dn org-root/memory-config-default | Set-UcsManagedObject -PropertyMap @{PolicyOwner="local"; BlackListing="enabled"; Descr=""; } -Force
	}
else
	{
		Write-Host ""
		Write-Host -ForegroundColor DarkBlue "Skipping configuration of DIMM Blacklisting due to UCSM or PowerTool Version"
	}

if (([int]$UCSMajorVersion -ge 2) -and ([int]$UCSMinorVersion -ge 2) -and ($PTUCSM -ge 2.2))
	{
		#UDLD System Setting
		Write-Host -ForegroundColor DarkBlue "Configuring UDLD"
		Write-Host ""
		Get-UcsManagedObject -Dn org-root/udld-policy | Set-UcsManagedObject -PropertyMap @{PolicyOwner="local"; MsgInterval="15"; Descr=""; Name="default"; RecoveryAction="reset"; } -Force
	}
else
	{
		Write-Host ""
		Write-Host -ForegroundColor DarkBlue "Skipping configuration of UDLD due to UCSM or PowerTool Version"
	}

#Servers - Create Boot Policies
Write-Host -ForegroundColor DarkBlue "Setting Boot Policies"
##Boot from Local HD
If ($BootFromHD -ieq "y")
	{
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsBootPolicy -Descr "Boot from local hard drive" -EnforceVnicName "yes" -Name "Boot_From_HD" -PolicyOwner "local" -RebootOnUpdate "no"
		if ($PXENIC -ne "")
			{
				$mo_1 = $mo | Add-UcsLsbootLan -ModifyPresent -Order "3" -Prot "pxe" 
				$mo_1_1 = $mo_1 | Add-UcsLsbootLanImagePath -BootIpPolicyName "" -ISCSIVnicName "" -ImgPolicyName "" -ImgSecPolicyName "" -ProvSrvPolicyName "" -Type "primary" -VnicName $PXENIC
			}
		$mo_2 = $mo | Add-UcsLsbootVirtualMedia -Access "read-only" -Order "2"
		$mo_3 = $mo | Add-UcsLsbootStorage -Order 1
		$mo_3_1 = $mo_3 | Add-UcsLsbootLocalStorage
		$mo_3_1_1 = $mo_3_1 | Add-UcsLsbootDefaultLocalImage -Order 1
		Complete-UcsTransaction
	}

##Boot from SAN
If ($BootFromSAN -ieq "y")
	{
		foreach ($bm in $BootMatrix)
			{
				$APrimary = $bm["APrimary"]
				$ASecondary = $bm["ASecondary"]
				$BPrimary = $bm["BPrimary"]
				$BSecondary = $bm["BSecondary"]
				if ($ArrayPort[$APrimary -1]["Fabric"] -ieq "A")
					{
						if ($VHBAnameB -ne "")
							{
								$Description = "A_Primary:"+$ArrayPort[$APrimary - 1]["Name"]+"_Fabric:"+$ArrayPort[$APrimary - 1]["Fabric"]+"___A_Secondary:"+$ArrayPort[$ASecondary - 1]["Name"]+"_Fabric:"+$ArrayPort[$ASecondary - 1]["Fabric"]+"___B_Primary:"+$ArrayPort[$BPrimary - 1]["Name"]+"_Fabric:"+$ArrayPort[$BPrimary - 1]["Fabric"]+"___B_Secondary:"+$ArrayPort[$BSecondary - 1]["Name"]+"_Fabric:"+$ArrayPort[$BSecondary - 1]["Fabric"]
							}
						else
							{
								$Description = "A_Primary:"+$ArrayPort[$APrimary - 1]["Name"]+"_Fabric:"+$ArrayPort[$APrimary - 1]["Fabric"]+"___A_Secondary:"+$ArrayPort[$ASecondary - 1]["Name"]+"_Fabric:"+$ArrayPort[$ASecondary - 1]["Fabric"]
							}
						$FirstHBA = $VHBAnameA.Substring(0,($VHBAnameA.Length-1))+$ArrayPort[$APrimary - 1]["Fabric"]
						$FirstHBAPrimary = $ArrayPort[$APrimary - 1]["WWPN"]
						$FirstHBASecondary = $ArrayPort[$ASecondary - 1]["WWPN"]
						if ($BPrimary -ne "")
							{
								$SecondHBA = $VHBAnameA.Substring(0,($VHBAnameA.Length-1))+$ArrayPort[$BPrimary - 1]["Fabric"]
								$SecondHBAPrimary = $ArrayPort[$BPrimary - 1]["WWPN"]
								$SecondHBASecondary = $ArrayPort[$BSecondary - 1]["WWPN"]
							}
						$BootFromSANName = "Boot_from_"+$bm["Name"]
						Start-UcsTransaction
						$mo = Get-UcsOrg -Level root  | Add-UcsBootPolicy -Descr $Description -EnforceVnicName "yes" -Name $BootFromSANName -PolicyOwner "local" -RebootOnUpdate "no"
						if ($PXENIC -ne "")
							{
								$mo_1 = $mo | Add-UcsLsbootLan -ModifyPresent -Order "3" -Prot "pxe" 
								$mo_1_1 = $mo_1 | Add-UcsLsbootLanImagePath -BootIpPolicyName "" -ISCSIVnicName "" -ImgPolicyName "" -ImgSecPolicyName "" -ProvSrvPolicyName "" -Type "primary" -VnicName $PXENIC
							}
						$mo_2 = $mo | Add-UcsLsbootVirtualMedia -Access "read-only" -Order "2"
						$mo_3 = $mo | Add-UcsLsbootStorage -ModifyPresent -Order "1"
						$mo_3_1 = $mo_3 | Add-UcsLsbootSanImage -Type "primary" -VnicName $FirstHBA
						$mo_3_1_1 = $mo_3_1 | Add-UcsLsbootSanImagePath -Lun 0 -Type "primary" -Wwn $FirstHBAPrimary
						if ($ASecondary -ne "")
							{
								$mo_3_1_2 = $mo_3_1 | Add-UcsLsbootSanImagePath -Lun 0 -Type "secondary" -Wwn $FirstHBASecondary
							}
						if ($BPrimary -ne "")
							{
								$mo_3_2 = $mo_3 | Add-UcsLsbootSanImage -Type "secondary" -VnicName $SecondHBA
								$mo_3_2_1 = $mo_3_2 | Add-UcsLsbootSanImagePath -Lun 0 -Type "primary" -Wwn $SecondHBAPrimary
							}
						if ($BSecondary -ne "")
							{
								$mo_3_2_2 = $mo_3_2 | Add-UcsLsbootSanImagePath -Lun 0 -Type "secondary" -Wwn $SecondHBASecondary
							}
						Complete-UcsTransaction
					}
				if ($ArrayPort[$APrimary -1]["Fabric"] -ieq "B")
					{
						$Description = "B_Primary:"+$ArrayPort[$APrimary - 1]["Name"]+"_Fabric:"+$ArrayPort[$APrimary - 1]["Fabric"]+"___B_Secondary:"+$ArrayPort[$ASecondary - 1]["Name"]+"_Fabric:"+$ArrayPort[$ASecondary - 1]["Fabric"]+"___A_Primary:"+$ArrayPort[$BPrimary - 1]["Name"]+"_Fabric:"+$ArrayPort[$BPrimary - 1]["Fabric"]+"___A_Secondary:"+$ArrayPort[$BSecondary - 1]["Name"]+"_Fabric:"+$ArrayPort[$BSecondary - 1]["Fabric"]
						$FirstHBA = $VHBAnameA.Substring(0,($VHBAnameA.Length-1))+$ArrayPort[$APrimary - 1]["Fabric"]
						$FirstHBAPrimary = $ArrayPort[$APrimary - 1]["WWPN"]
						$FirstHBASecondary = $ArrayPort[$ASecondary - 1]["WWPN"]
						$SecondHBA = $VHBAnameA.Substring(0,($VHBAnameA.Length-1))+$ArrayPort[$BPrimary - 1]["Fabric"]
						$SecondHBAPrimary = $ArrayPort[$BPrimary - 1]["WWPN"]
						$SecondHBASecondary = $ArrayPort[$BSecondary - 1]["WWPN"]
						$BootFromSANName = "Boot_from_"+$bm["Name"]
						Start-UcsTransaction
						$mo = Get-UcsOrg -Level root  | Add-UcsBootPolicy -Descr $Description -EnforceVnicName "yes" -Name $BootFromSANName -PolicyOwner "local" -RebootOnUpdate "no"
						if ($PXENIC -ne "")
							{
								$mo_1 = $mo | Add-UcsLsbootLan -ModifyPresent -Order "3" -Prot "pxe" 
								$mo_1_1 = $mo_1 | Add-UcsLsbootLanImagePath -BootIpPolicyName "" -ISCSIVnicName "" -ImgPolicyName "" -ImgSecPolicyName "" -ProvSrvPolicyName "" -Type "primary" -VnicName $PXENIC
							}
						$mo_2 = $mo | Add-UcsLsbootVirtualMedia -Access "read-only" -Order "2"
						$mo_3 = $mo | Add-UcsLsbootStorage -ModifyPresent -Order "1"
						$mo_3_1 = $mo_3 | Add-UcsLsbootSanImage -Type "primary" -VnicName $FirstHBA
						$mo_3_1_1 = $mo_3_1 | Add-UcsLsbootSanImagePath -Lun 0 -Type "primary" -Wwn $FirstHBAPrimary
						if ($BSecondary -ne "")
							{
								$mo_3_1_2 = $mo_3_1 | Add-UcsLsbootSanImagePath -Lun 0 -Type "secondary" -Wwn $FirstHBASecondary
							}
						$mo_3_2 = $mo_3 | Add-UcsLsbootSanImage -Type "secondary" -VnicName $SecondHBA
						$mo_3_2_1 = $mo_3_2 | Add-UcsLsbootSanImagePath -Lun 0 -Type "primary" -Wwn $SecondHBAPrimary
						if ($ASecondary -ne "")
							{
								$mo_3_2_2 = $mo_3_2 | Add-UcsLsbootSanImagePath -Lun 0 -Type "secondary" -Wwn $SecondHBASecondary
							}
						Complete-UcsTransaction
					}
			}
	}

If ($BootFromiSCSI -ieq "y")
	{
		##Boot from iSCSI
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsBootPolicy -Descr "" -EnforceVnicName "yes" -Name "Boot_from_iSCSI" -PolicyOwner "local" -RebootOnUpdate "no"
		$mo_1 = $mo | Add-UcsLsbootIScsi -ModifyPresent -Order "1"
		$mo_1_1 = $mo_1 | Add-UcsLsbootIScsiImagePath -ISCSIVnicName $iSCSINicNameA -Type "primary" -VnicName ""
		if ($IsRedundantFI -ne $null)
			{
				$mo_1_2 = $mo_1 | Add-UcsLsbootIScsiImagePath -ISCSIVnicName $iSCSINicNameB -Type "secondary" -VnicName ""
			}
		if ($PXENIC -ne "")
			{
				$mo_2 = $mo | Add-UcsLsbootLan -ModifyPresent -Order "3" -Prot "pxe"
				$mo_2_1 = $mo_2 | Add-UcsLsbootLanImagePath -BootIpPolicyName "" -ISCSIVnicName "" -ImgPolicyName "" -ImgSecPolicyName "" -ProvSrvPolicyName "" -Type "primary" -VnicName $PXENIC
			}
		$mo_3 = $mo | Add-UcsLsbootVirtualMedia -Access "read-only" -Order "2"
		Complete-UcsTransaction

		#Boot Parameters
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "Boot_from_iSCSI" -LimitScope | Add-UcsVnicIScsiBootParams -Descr "" -PolicyOwner "local"
		$mo_1 = $mo | Add-UcsVnicIScsiBootVnic -AuthProfileName "" -Descr "" -InitiatorName "" -IqnIdentPoolName "IQN_Pool" -Name $iSCSINicNameA -PolicyOwner "local"
		$mo_1_1 = $mo_1 | Add-UcsVnicIScsiStaticTargetIf -AuthProfileName "" -IpAddress $iSCSItargetA -Name $iSCSIiqn -Port 3260 -Priority 1
		$mo_1_1_1 = $mo_1_1 | Add-UcsVnicLun -ModifyPresent -Bootable "no" -Id 0
		$mo_1_2 = $mo_1 | Add-UcsVnicIPv4If -Name ""
		if ($iSCSIPool -ieq "Pool")
			{
				$mo_1_2_1 = $mo_1_2 | Add-UcsManagedObject -ClassId VnicIPv4PooledIscsiAddr -PropertyMap @{IdentPoolName=$DefaultiSCSIInitiatorPoolA; }
			}
		else
			{
				$mo_1_2_1 = $mo_1_2 | Add-UcsVnicIPv4Dhcp
			}
		Complete-UcsTransaction
		if ($IsRedundantFI -ne $null)
			{
				Start-UcsTransaction
				$mo = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "Boot_from_iSCSI" -LimitScope | Get-UcsVnicIScsi -Name "boot-params" | Set-UcsManagedObject -PropertyMap @{Descr=""; PolicyOwner="local"; }
				$mo_1 = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "Boot_from_iSCSI" -LimitScope | Get-UcsVnicIScsi -Name "boot-params" | Get-UcsVnicIScsiBootVnic -Name $iSCSINicNameB | Set-UcsVnicIScsiBootVnic -AuthProfileName "" -Descr "" -InitiatorName "" -IqnIdentPoolName "IQN_Pool" -PolicyOwner "local"
				$mo_1_1 = $mo_1 | Add-UcsVnicIScsiStaticTargetIf -AuthProfileName "" -IpAddress $iSCSItargetB -Name $iSCSIiqn -Port 3260 -Priority 1
				$mo_1_1_1 = $mo_1_1 | Add-UcsVnicLun -ModifyPresent -Bootable "no" -Id 0
				$mo_1_2 = $mo_1 | Add-UcsVnicIPv4If -Name ""
				if ($iSCSIPool -ieq "Pool")
					{
						$mo_1_2_1 = $mo_1_2 | Add-UcsManagedObject -ClassId VnicIPv4PooledIscsiAddr -PropertyMap @{IdentPoolName=$DefaultiSCSIInitiatorPoolA; }
					}
				else
					{
						$mo_1_2_1 = $mo_1_2 | Add-UcsVnicIPv4Dhcp
					}
				Complete-UcsTransaction
			}
	}
		else
			{
			}
	
#Servers - Set IPMI Access Profiles
Write-Host -ForegroundColor DarkBlue "Setting IPMI Policies"
##Admin Rights
Start-UcsTransaction
$mo = Get-UcsOrg -Level root  | Add-UcsIpmiAccessProfile -Descr "full admin rights" -Name "admin" -PolicyOwner "local" 
$mo_1 = $mo | Add-UcsAaaEpUser -Descr "" -Name "admin" -Priv "admin" -Pwd $IPMIpassword 
Complete-UcsTransaction

##Read-Only Rights
Start-UcsTransaction
$mo = Get-UcsOrg -Level root  | Add-UcsIpmiAccessProfile -Descr "read-only rights" -Name "readonly" -PolicyOwner "local"
$mo_1 = $mo | Add-UcsAaaEpUser -Descr "" -Name "readonly" -Priv "readonly" -Pwd $IPMIpassword
Complete-UcsTransaction

#Servers - Set Local Disk Config Policies
Write-Host -ForegroundColor DarkBlue "Setting Local Disk Config Policies"
Get-UcsOrg -Level root  | Add-UcsLocalDiskConfigPolicy -Descr "" -Mode "any-configuration" -Name "Any_Config" -PolicyOwner "local" -ProtectConfig "yes"
Get-UcsOrg -Level root  | Add-UcsLocalDiskConfigPolicy -Descr "" -Mode "no-local-storage" -Name "No_Local_Storage" -PolicyOwner "local" -ProtectConfig "yes"
Get-UcsOrg -Level root  | Add-UcsLocalDiskConfigPolicy -Descr "" -Mode "raid-striped" -Name "RAID0" -PolicyOwner "local" -ProtectConfig "yes"
Get-UcsOrg -Level root  | Add-UcsLocalDiskConfigPolicy -Descr "" -Mode "raid-mirrored" -Name "RAID1" -PolicyOwner "local" -ProtectConfig "yes"
Get-UcsOrg -Level root  | Add-UcsLocalDiskConfigPolicy -Descr "" -Mode "no-raid" -Name "No_RAID" -PolicyOwner "local" -ProtectConfig "yes"
Get-UcsOrg -Level root  | Add-UcsLocalDiskConfigPolicy -Descr "" -Mode "raid-striped-parity" -Name "RAID5" -PolicyOwner "local" -ProtectConfig "yes"
Get-UcsOrg -Level root  | Add-UcsLocalDiskConfigPolicy -Descr "" -Mode "raid-striped-dual-parity" -Name "RAID6" -PolicyOwner "local" -ProtectConfig "yes"
Get-UcsOrg -Level root  | Add-UcsLocalDiskConfigPolicy -Descr "" -Mode "raid-mirrored-striped" -Name "RAID10" -PolicyOwner "local" -ProtectConfig "yes"
if (([int]$UCSMajorVersion -ge 2) -and ([int]$UCSMinorVersion -ge 2) -and ($PTUCSM -ge 2.2))
	{
		Get-UcsOrg -Level root  | Add-UcsLocalDiskConfigPolicy -Descr "" -Mode "raid-striped-parity-striped" -Name "RAID50" -PolicyOwner "local" -ProtectConfig "yes"
		Get-UcsOrg -Level root  | Add-UcsLocalDiskConfigPolicy -Descr "" -Mode "raid-striped-dual-parity-striped" -Name "RAID60" -PolicyOwner "local" -ProtectConfig "yes"
	}
else
	{
		Write-Host -ForegroundColor DarkBlue "Skipping configuration of RAID50 and RAID60 due to UCSM or PowerTool version"
		Write-Host ""
	}

#Set Maintenance Policies
Write-Host -ForegroundColor DarkBlue "Setting Maintenance Policies"
#Servers - Modify default Policy to User-Ack
Get-UcsOrg -Level root | Get-UcsMaintenancePolicy -Name "default" -LimitScope | Set-UcsMaintenancePolicy -Descr "" -PolicyOwner "local" -SchedName "" -UptimeDisr "user-ack" -Force
#Servers - Add User_Ack policy
Get-UcsOrg -Level root  | Add-UcsMaintenancePolicy -Descr "" -Name "User_Ack" -PolicyOwner "local" -SchedName "" -UptimeDisr "user-ack"

#Servers - Set Power Capping Policy
Write-Host -ForegroundColor DarkBlue "Setting Power Capping Policies"
Get-UcsOrg -Level root  | Add-UcsPowerPolicy -Descr "" -Name "No_Cap" -PolicyOwner "local" -Prio "no-cap"

#Servers - Set Scrub Policies
Write-Host -ForegroundColor DarkBlue "Setting Scrub Policies"
Get-UcsOrg -Level root  | Add-UcsScrubPolicy -BiosSettingsScrub "no" -Descr "" -DiskScrub "no" -Name "No_Scrub" -PolicyOwner "local"
Get-UcsOrg -Level root  | Add-UcsScrubPolicy -BiosSettingsScrub "yes" -Descr "" -DiskScrub "no" -Name "BIOS_Scrub" -PolicyOwner "local"
Get-UcsOrg -Level root  | Add-UcsScrubPolicy -BiosSettingsScrub "no" -Descr "" -DiskScrub "yes" -Name "Disk_Scrub" -PolicyOwner "local"
Get-UcsOrg -Level root  | Add-UcsScrubPolicy -BiosSettingsScrub "yes" -Descr "" -DiskScrub "yes" -Name "Full_Scrub" -PolicyOwner "local"

#Servers - Serial over LAN Policies
Write-Host -ForegroundColor DarkBlue "Setting Serial over LAN policies"
Get-UcsOrg -Level root  | Add-UcsSolPolicy -AdminState "disable" -Descr "" -Name "No_SoL" -PolicyOwner "local" -Speed "9600"
Get-UcsOrg -Level root  | Add-UcsSolPolicy -AdminState "enable"  -Descr "" -Name "SoL_9600" -PolicyOwner "local" -Speed "9600"
Get-UcsOrg -Level root  | Add-UcsSolPolicy -AdminState "enable"  -Descr "" -Name "SoL_19200" -PolicyOwner "local" -Speed "19200"
Get-UcsOrg -Level root  | Add-UcsSolPolicy -AdminState "enable"  -Descr "" -Name "SoL_38400" -PolicyOwner "local" -Speed "38400"
Get-UcsOrg -Level root  | Add-UcsSolPolicy -AdminState "enable"  -Descr "" -Name "SoL_57600" -PolicyOwner "local" -Speed "57600"
Get-UcsOrg -Level root  | Add-UcsSolPolicy -AdminState "enable"  -Descr "" -Name "SoL_115200" -PolicyOwner "local" -Speed "115200"

#Servers - Server Pool Policy Qualifications
Write-Host -ForegroundColor DarkBlue "Setting Server Pool Qualification Policies"
#Add all servers in all chassis'
Start-UcsTransaction
$mo = Get-UcsOrg -Level root  | Add-UcsServerPoolQualification -Descr "" -Name "All_Chassis" -PolicyOwner "local"
$mo_1 = $mo | Add-UcsChassisQualification -MaxId 255 -MinId 1
Complete-UcsTransaction
#Add all servers in all rack servers
Start-UcsTransaction
$mo = Get-UcsOrg -Level root  | Add-UcsServerPoolQualification -Descr "" -Name "All_Rack" -PolicyOwner "local"
$mo_1 = $mo | Add-UcsRackQualification -MaxId 255 -MinId 1
Complete-UcsTransaction

#LAN - Flow Control Polices
#Write-Host -ForegroundColor DarkBlue "Setting Flow Control Policies"
#Add-UcsFlowctrlItem -Name "No_Flow_Control" -Prio "auto" -Rcv "off" -Snd "off"

#LAN - Multicast Policies
Write-Host -ForegroundColor DarkBlue "Setting Multicast Policies"
Get-UcsOrg -Level root  | Add-UcsFabricMulticastPolicy -Descr "" -Name "Snooping_DEFAULT" -PolicyOwner "local" -QuerierIpAddr "0.0.0.0" -QuerierState "disabled" -SnoopingState "enabled"

Get-UcsOrg -Level root  | Add-UcsFabricMulticastPolicy -Descr "" -Name "Snooping_Querier" -PolicyOwner "local" -QuerierIpAddr $QuerierIpAddr -QuerierState "enabled" -SnoopingState "enabled"
Get-UcsOrg -Level root  | Add-UcsFabricMulticastPolicy -Descr "" -Name "Off" -PolicyOwner "local" -QuerierIpAddr "0.0.0.0" -QuerierState "disabled" -SnoopingState "disabled"

#LAN - QoS Policies
Write-Host -ForegroundColor DarkBlue "Setting LAN QoS Policies"
#Set Global QoS Settings
Write-Host -ForegroundColor DarkBlue "Setting Global QoS settings"
#FibreChannel
Start-UcsTransaction
$mo = Get-UcsOrg -Level root  | Add-UcsQosPolicy -Descr "" -Name "FibreChannel" -PolicyOwner "local"
$mo_1 = $mo | Add-UcsVnicEgressPolicy -ModifyPresent -Burst $FibreChannelQoSBurst -HostControl $FibreChannelQoSHostControl -Name "" -Prio "fc" -Rate $FibreChannelQoSRate
Complete-UcsTransaction
#Best-Effort
Start-UcsTransaction
$mo = Get-UcsOrg -Level root  | Add-UcsQosPolicy -Descr "" -Name "BestEffort" -PolicyOwner "local"
$mo_1 = $mo | Add-UcsVnicEgressPolicy -ModifyPresent -Burst $BestEffortQoSBurst -HostControl $BestEffortQoSHostControl -Name "" -Prio "best-effort" -Rate $BestEffortQoSRate
Complete-UcsTransaction
#Bronze
if ($BronzeQoSEnabled -ieq "y")
	{
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsQosPolicy -Descr "" -Name "Bronze" -PolicyOwner "local"
		$mo_1 = $mo | Add-UcsVnicEgressPolicy -ModifyPresent -Burst $BronzeQoSBurst -HostControl $BronzeQoSHostControl -Name "" -Prio "bronze" -Rate $BronzeQoSRate
		Complete-UcsTransaction
	}
#Silver
if ($SilverQoSEnabled -ieq "y")
	{
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsQosPolicy -Descr "" -Name "Silver" -PolicyOwner "local"
		$mo_1 = $mo | Add-UcsVnicEgressPolicy -ModifyPresent -Burst $SilverQoSBurst -HostControl $SilverQoSHostControl -Name "" -Prio "silver" -Rate $SilverQoSRate
		Complete-UcsTransaction
	}
#Gold
if ($GoldQoSEnabled -ieq "y")
	{
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsQosPolicy -Descr "" -Name "Gold" -PolicyOwner "local"
		$mo_1 = $mo | Add-UcsVnicEgressPolicy -ModifyPresent -Burst $GoldQoSBurst -HostControl $GoldQoSHostControl -Name "" -Prio "gold" -Rate $GoldQoSRate
		Complete-UcsTransaction
	}
#Platinum
if ($PlatinumQoSEnabled -ieq "y")
	{
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsQosPolicy -Descr "" -Name "Platinum" -PolicyOwner "local"
		$mo_1 = $mo | Add-UcsVnicEgressPolicy -ModifyPresent -Burst $PlatinumQoSBurst -HostControl $PlatinumQoSHostControl -Name "" -Prio "platinum" -Rate $PlatinumQoSRate
		Complete-UcsTransaction
	}

#Create Server Pool
Start-UcsTransaction
$mo = Get-UcsOrg -Level root  | Add-UcsServerPool -Descr "" -Name "All_Servers" -PolicyOwner "local"
Complete-UcsTransaction

# Automatic entry of blade and rack servers into server pool
Get-UcsOrg -Level root  | Add-UcsServerPoolPolicy -Descr "" -Name "Add_Blades" -PolicyOwner "local" -PoolDn "org-root/compute-pool-All_Servers" -Qualifier "All_Chassis"
Get-UcsOrg -Level root  | Add-UcsServerPoolPolicy -Descr "" -Name "Add_Racks" -PolicyOwner "local" -PoolDn "org-root/compute-pool-All_Servers" -Qualifier "All_Rack"

#Acknowledge all Chassis' so all links are active
$reack = $null
$chassis = $null
$chassiscount = $null
$chassisid = $null
$chassis = get-ucschassis
if ($chassis -ne $null)
	{
		Write-Host -ForegroundColor White -BackgroundColor DarkBlue "Re-Acking Chassis' to activate all IOM links....Please Wait"
		foreach ($chassiscount in $chassis)
			{
				$chassisid = [string]$chassiscount.Id
				Write-Host -ForegroundColor DarkBlue "   Re-Acking Chassis: "$chassisid
				#DO NOT REACK CHASSIS IN THE EMULATOR
				if ($UCSEmulator -ine "y")
					{
						$reack = Set-UcsChassis -chassis $chassisid -AdminState re-acknowledge -Force
					}
			}
	}

#Acknowledge all FEXs so all links are active
$reack = $null
$fex = $null
$fexcount = $null
$fexid = $null
$fex = Get-UcsFex
if ($fex -ne $null)
	{
		Write-Host -ForegroundColor White -BackgroundColor DarkBlue "Re-Acking Fexs to activate all FEX links....Please Wait"
		foreach ($fexcount in $fex)
			{
				$fexid = [string]$fexcount.Id
				Write-Host -ForegroundColor DarkBlue "   Re-Acking Fex: "$fexid
				##DO NOT REACK FEX IN THE EMULATOR
				if ($UCSEmulator -ine "y")
					{
						$reack = set-ucsfex -Fex $fexid -AdminState re-acknowledge -Force
					}
			}
	}

#Servers - UUID Suffix Pools
Write-Host -ForegroundColor DarkBlue "Setting UUID Suffix Pools"
Start-UcsTransaction
$mo = Get-UcsOrg -Level root  | Add-UcsUuidSuffixPool -AssignmentOrder "sequential" -Descr "" -Name "UUID_Pool" -PolicyOwner "local" -Prefix "derived"
$mo_1 = $mo | Add-UcsUuidSuffixBlock -From $UUIDfrom -To $UUIDto
Complete-UcsTransaction

#LAN - IP Pool - IP Pool ext-mgmt
Write-Host -ForegroundColor DarkBlue "Setting IP Management Pools"
$NetworkElement = Get-UcsNetworkElement -Id "A"
$Subnet = $NetworkElement.oobIfMask
Get-UcsOrg -Level root | Get-UcsIpPool -Name "ext-mgmt" -LimitScope | Set-UcsIpPool -AssignmentOrder "sequential" -Descr "" -PolicyOwner "local" -Force
if ($PriDNS -ne "")
	{
		Get-UcsOrg -Level root | Get-UcsIpPool -Name "ext-mgmt" -LimitScope | Add-UcsIpPoolBlock -DefGw $DefGw -From $MgmtIPstart -PrimDns $PriDNS -SecDns $SecDNS -To $MgmtIPend -Subnet $Subnet
	}
else
	{
		Get-UcsOrg -Level root | Get-UcsIpPool -Name "ext-mgmt" -LimitScope | Add-UcsIpPoolBlock -DefGw $DefGw -From $MgmtIPstart -To $MgmtIPend -Subnet $Subnet
	}
#LAN - IP Pool - IP Pool iscsi-initiator-pool
##Default Pool - Cannot be deleted so just putting bogus entries in it to prevent it from throwing a UCS error
Write-Host -ForegroundColor DarkBlue "Setting iSCSI Initiator Pools"
Get-UcsOrg -Level root | Get-UcsIpPool -Name "iscsi-initiator-pool" -LimitScope | Set-UcsIpPool -AssignmentOrder "sequential" -Descr "DO NOT USE THIS POOL" -PolicyOwner "local" -Force
Get-UcsOrg -Level root | Get-UcsIpPool -Name "iscsi-initiator-pool" -LimitScope | Add-UcsIpPoolBlock -DefGw $DefaultiSCSIPoolDefGW -From $DefaultiSCSIPoolFrom -PrimDns $DefaultiSCSIPrimDNS -SecDns $DefaultiSCSISecDNS -To $DefaultiSCSIPoolTo -Subnet $DefaultiSCSIPoolSubnet
##Fabric A Pool
if ($BootFromiSCSI -ieq "y")
	{
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsIpPool -AssignmentOrder "sequential" -Descr "" -Name $DefaultiSCSIInitiatorPoolA -PolicyOwner "local"
		$mo_1 = $mo | Add-UcsIpPoolBlock -DefGw $iSCSIDefGwA -From $iSCSIIPstartA -PrimDns $PriDNS -SecDns $SecDNS -To $iSCSIIPendA -Subnet $iSCSISubnetA
		Complete-UcsTransaction
		if ($IsRedundantFI -ne $null)
			{
				##Fabric B Pool
				Start-UcsTransaction
				$mo = Get-UcsOrg -Level root  | Add-UcsIpPool -AssignmentOrder "sequential" -Descr "" -Name $DefaultiSCSIInitiatorPoolB -PolicyOwner "local"
				$mo_1 = $mo | Add-UcsIpPoolBlock -DefGw $iSCSIDefGwB -From $iSCSIIPstartB -PrimDns $PriDNS -SecDns $SecDNS -To $iSCSIIPendB -Subnet $iSCSISubnetB
				Complete-UcsTransaction
			}
	}

#LAN - MAC Pools
Write-Host -ForegroundColor DarkBlue "Setting MAC Address Pools"
$networkcount = $null
foreach ($networkcount in $network)
	{
		if (($networkcount["macid"] -ne $null) -and ($networkcount["macid"] -ne ""))
			{
				$MACfromNEW = $MACfrom -replace "ID", $networkcount["macid"]
				$MACtoNEW = $MACto -replace "ID", $networkcount["macid"]
				Start-UcsTransaction
				$mo = Get-UcsOrg -Level root  | Add-UcsMacPool -AssignmentOrder "sequential" -Descr "" -Name $networkcount["vlanname"] -PolicyOwner "local"
				$mo_1 = $mo | Add-UcsMacMemberBlock -From $MACfromNEW -To $MACtoNEW
				Complete-UcsTransaction
			}
	}

#LAN - LAN Cloud - VLANs
##Add VLANs
Write-Host -ForegroundColor DarkBlue "Setting VLANs"
$networkcount = $null
foreach ($networkcount in $network)
	{
		if (($networkcount["vlannumber"] -ne $null) -and ($networkcount["vlannumber"] -ne ""))
			{
				Get-UcsLanCloud | Add-UcsVlan -CompressionType "included" -DefaultNet "no" -Id $networkcount["vlannumber"] -McastPolicyName "" -Name $networkcount["vlanname"] -PubNwName "" -Sharing "none"
			}
	}

#LAN - LAN Cloud - VLANS
##Native VLAN
Write-Host -ForegroundColor DarkBlue "Setting Native VLAN for Uplink LAN Ports"
$NativeCheck = Get-UcsLanCloud | Get-UcsVlan | where {$_.id -match $NativeVLANnumber}
if ($NativeCheck.Id -notcontains $NativeVLANnumber)
	{
		Get-UcsLanCloud | Add-UcsVlan -CompressionType "included" -DefaultNet "no" -Id $NativeVLANnumber -McastPolicyName "" -Name $NativeVLANname -PubNwName "" -Sharing "none"
		Get-UcsLanCloud | Get-UcsVlan -Name $NativeVLANname -LimitScope | Set-UcsVlan -CompressionType "included" -DefaultNet "yes" -Id $NativeVLANnumber -McastPolicyName "" -PubNwName "" -Sharing "none" -Force
	}
else
	{
		Get-UcsLanCloud | Get-UcsVlan -Name $NativeVLANname -LimitScope | Set-UcsVlan -CompressionType "included" -DefaultNet "yes" -Id $NativeVLANnumber -McastPolicyName "" -PubNwName "" -Sharing "none" -Force
	}

#LAN - vNIC Templates
Write-Host -ForegroundColor DarkBlue "Setting vNIC Templates"
$networkcount = $null
$AllVLANs = Get-UcsVlan
foreach ($networkcount in $network)
	{
		if (($networkcount["macid"] -ne $null) -and ($networkcount["macid"] -ne ""))
			{
				Start-UcsTransaction
				$mo = Get-UcsOrg -Level root  | Add-UcsVnicTemplate -Descr "" -IdentPoolName $networkcount["vlanname"] -Mtu $networkcount["mtu"] -Name $networkcount["vlanname"] -NwCtrlPolicyName "" -PinToGroupName "" -PolicyOwner "local" -QosPolicyName $networkcount["QoSPolicy"] -StatsPolicyName "default" -SwitchId $networkcount["fabric"] -TemplType "updating-template"
				if ($AllVLANs | where {$_.Name -ieq $networkcount["vlanname"]})
					{
						$mo_1 = $mo | Add-UcsVnicInterface -ModifyPresent -DefaultNet "yes" -Name $networkcount["vlanname"]
					}
				Complete-UcsTransaction
			}
	}
$networkcount = $null
foreach ($networkcount in $network)
	{
		if (($networkcount["trunknic"] -ne $null) -and ($networkcount["trunknic"] -ne ""))
			{
				Start-UcsTransaction
				$mo = Get-UcsOrg -Level root  | Add-UcsVnicTemplate -ModifyPresent -Name $networkcount["trunknic"]
				if ($networkcount["defaultvlan"] -ine "y")
					{
						$mo_1 = $mo | Add-UcsVnicInterface -ModifyPresent -DefaultNet "no" -Name $networkcount["vlanname"]
					}
				else
					{
						$mo_2 = $mo | Add-UcsVnicInterface -ModifyPresent -DefaultNet "yes" -Name $networkcount["vlanname"]
					}
				Complete-UcsTransaction
			}
	}
	
#LAN - LAN Connectivity Policies
#vNICs
Write-Host -ForegroundColor DarkBlue "Setting LAN Connectivity Policies"
Get-UcsOrg -Level root  | Add-UcsVnicLanConnPolicy -Descr "" -Name $DefaultLANConnectivity -PolicyOwner "local"
$count = 0
foreach ($nw in $network)
	{
		if ($nw.macid -ne "")
			{
				$count =+ 1
				$mo = Get-UcsOrg -Level root | Get-UcsVnicLanConnPolicy -Name $DefaultLANConnectivity -LimitScope -Descr "" -PolicyOwner "local"
				Start-UcsTransaction
				$mo_1 = $mo | Add-UcsVnic -AdaptorProfileName $DefaultLANAdapter -Addr "derived" -AdminVcon "any" -IdentPoolName "" -Mtu $nw["mtu"] -Name $nw["vlanname"] -NwCtrlPolicyName "" -NwTemplName $nw["vlanname"] -Order $count -PinToGroupName "" -QosPolicyName "" -StatsPolicyName "default" -SwitchId $nw["fabric"]
				Complete-UcsTransaction
			}
	}

#vNICs_iSCSI
if ($BootFromiSCSI -ieq "y")
	{
	Get-UcsOrg -Level root  | Add-UcsVnicLanConnPolicy -Descr "" -Name $DefaultLANwiSCSIConnectivity -PolicyOwner "local"
$count = 0
foreach ($nw in $network)
	{
		if ($nw.macid -ne "")
			{
				$count =+ 1
				$mo = Get-UcsOrg -Level root | Get-UcsVnicLanConnPolicy -Name $DefaultLANwiSCSIConnectivity -LimitScope -Descr "" -PolicyOwner "local"
				Start-UcsTransaction
				$mo_1 = $mo | Add-UcsVnic -AdaptorProfileName $DefaultLANAdapter -Addr "derived" -AdminVcon "any" -IdentPoolName "" -Mtu $nw["mtu"] -Name $nw["vlanname"] -NwCtrlPolicyName "" -NwTemplName $nw["vlanname"] -Order $count -PinToGroupName "" -QosPolicyName "" -StatsPolicyName "default" -SwitchId $nw["fabric"]
				Complete-UcsTransaction
			}
	}
		$mo = Get-UcsOrg -Level root | Get-UcsVnicLanConnPolicy -Name $DefaultLANwiSCSIConnectivity -LimitScope -Descr "" -PolicyOwner "local"
		Start-UcsTransaction
		$mo_1 = $mo | Add-UcsVnicIScsiLCP -AdaptorProfileName "default" -Addr "derived" -AdminVcon "any" -IdentPoolName "" -Name $iSCSINicNameA -NwTemplName "" -Order "unspecified" -PinToGroupName "" -QosPolicyName "" -StatsPolicyName "default" -SwitchId "A" -VnicName $iSCSINicNameA
		$mo_1_1 = $mo_1 | Add-UcsVnicVlan -ModifyPresent -Name "" -VlanName $iSCSINicNameA
		if ($IsRedundantFI -ne $null)
			{
				$mo_2 = $mo | Add-UcsVnicIScsiLCP -AdaptorProfileName "default" -Addr "derived" -AdminVcon "any" -IdentPoolName "" -Name $iSCSINicNameB -NwTemplName "" -Order "unspecified" -PinToGroupName "" -QosPolicyName "" -StatsPolicyName "default" -SwitchId "B" -VnicName $iSCSINicNameB
				$mo_2_1 = $mo_2 | Add-UcsVnicVlan -ModifyPresent -Name "" -VlanName $iSCSINicNameB
			}
		Complete-UcsTransaction
	}

#SAN - VSANs
Write-Host -ForegroundColor DarkBlue "Setting vSANs"
if ($FCPort[0].Port -ne "")
	{
		Write-Host -ForegroundColor DarkBlue "Setting vSANs"
		$vSANNameA = $DefaultvSANName+"_A"
		$vSANNameB = $DefaultvSANName+"_B"
		Get-UcsFiSanCloud -Id "A" | Add-UcsVsan -FcZoneSharingMode "coalesce" -FcoeVlan $FcoeVlanA -Id $VSANidA -Name $vSANNameA -ZoningState "disabled"
		if (($FcoeVlanB -ne "") -and ($VSANidB -ne ""))
			{
				Get-UcsFiSanCloud -Id "B" | Add-UcsVsan -FcZoneSharingMode "coalesce" -FcoeVlan $FcoeVlanB -Id $VSANidB -Name $vSANNameB -ZoningState "disabled"
			}
	}

#SAN - Pools - IQN Pools
if ($BootFromiSCSI -ieq "y")
	{
		Write-Host -ForegroundColor DarkBlue "Setting iSCSI IQN Pool"
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsIqnPoolPool -AssignmentOrder "sequential" -Descr "" -Name "IQN_Pool" -PolicyOwner "local" -Prefix "$IQNPrefix"
		$mo_1 = $mo | Add-UcsIqnPoolBlock -From $IQNFrom -Suffix "$IQNSuffix$UCSDomain" -To $IQNTo
		Complete-UcsTransaction
	}
	
#SAN - Pools - WWNN Pool
if ($FCPort[0].Port -ne "")
	{
		Write-Host -ForegroundColor DarkBlue "Set WWNN Pools"
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsWwnPool -AssignmentOrder "sequential" -Descr "" -Name "WWNN_Pool" -PolicyOwner "local" -Purpose "node-wwn-assignment"
		$mo_1 = $mo | Add-UcsWwnMemberBlock -From $WWNNfrom -To $WWNNto
		Complete-UcsTransaction
	}
	
#SAN - Pools - WWPN Pools
if ($FCPort[0].Port -ne "")
	{
		Write-Host -ForegroundColor DarkBlue "Set WWPN Pools"
		Start-UcsTransaction
		$WWPNPoolA = $DefaultWWPNPool+"_A"
		$WWPNPoolB = $DefaultWWPNPool+"_B"
		$mo = Get-UcsOrg -Level root  | Add-UcsWwnPool -AssignmentOrder "sequential" -Descr "" -Name $WWPNPoolA -PolicyOwner "local" -Purpose "port-wwn-assignment"
		$mo_1 = $mo | Add-UcsWwnMemberBlock -From $WWPNaFrom -To $WWPNaTo
		Complete-UcsTransaction
		if ($VHBAnameB -ne "")
			{
				Start-UcsTransaction
				$mo = Get-UcsOrg -Level root  | Add-UcsWwnPool -AssignmentOrder "sequential" -Descr "" -Name $WWPNPoolB -PolicyOwner "local" -Purpose "port-wwn-assignment"
				$mo_1 = $mo | Add-UcsWwnMemberBlock -From $WWPNbfrom -To $WWPNbto
				Complete-UcsTransaction
			}
	}
	
#SAN - vHBA Templates
if ($FCPort[0].Port -ne "")
	{
		Write-Host -ForegroundColor DarkBlue "Set vHBA Templates"
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsVhbaTemplate -Descr "" -IdentPoolName $WWPNPoolA -MaxDataFieldSize 2048 -Name $VHBAnameA -PinToGroupName "" -PolicyOwner "local" -QosPolicyName "FibreChannel" -StatsPolicyName "default" -SwitchId "A" -TemplType "updating-template"
		$mo_1 = $mo | Add-UcsVhbaInterface -ModifyPresent -Name $vSANNameA
		Complete-UcsTransaction

		if ($VHBAnameB -ne "")
			{
				Start-UcsTransaction
				$mo = Get-UcsOrg -Level root  | Add-UcsVhbaTemplate -Descr "" -IdentPoolName $WWPNPoolB -MaxDataFieldSize 2048 -Name $VHBAnameB -PinToGroupName "" -PolicyOwner "local" -QosPolicyName "FibreChannel" -StatsPolicyName "default" -SwitchId "B" -TemplType "updating-template"
				$mo_1 = $mo | Add-UcsVhbaInterface -ModifyPresent -Name $vSANNameB
				Complete-UcsTransaction
			}
	}

#SAN - SAN Connectivity Policy
if ($FCPort[0].Port -ne "")
	{
		Write-Host -ForegroundColor DarkBlue "Set SAN Connectivity Policies"
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsVnicSanConnPolicy -Descr "" -Name $DefaultHBAConnectivity -PolicyOwner "local"
		$mo_1 = $mo | Add-UcsVnicFcNode -ModifyPresent -Addr "pool-derived" -IdentPoolName "WWNN_Pool"
		$mo_2 = $mo | Add-UcsVhba -AdaptorProfileName "Windows" -Addr "derived" -AdminVcon "any" -IdentPoolName "" -MaxDataFieldSize 2048 -Name $VHBAnameA -NwTemplName $VHBAnameA -Order "1" -PersBind "disabled" -PersBindClear "no" -PinToGroupName "" -QosPolicyName "" -StatsPolicyName "default" -SwitchId "A"
		$mo_2_1 = $mo_2 | Add-UcsVhbaInterface -ModifyPresent -Name ""
		if ($VHBAnameB -ne "")
			{
				$mo_3 = $mo | Add-UcsVhba -AdaptorProfileName "Windows" -Addr "derived" -AdminVcon "any" -IdentPoolName "" -MaxDataFieldSize 2048 -Name $VHBAnameB -NwTemplName $VHBAnameB -Order "2" -PersBind "disabled" -PersBindClear "no" -PinToGroupName "" -QosPolicyName "" -StatsPolicyName "default" -SwitchId "B"
				$mo_3_1 = $mo_3 | Add-UcsVhbaInterface -ModifyPresent -Name ""
			}
		Complete-UcsTransaction
	}

#Admin - Management Interfaces
Write-Host -ForegroundColor DarkBlue "Set Administrative Management Interface Details"
Start-UcsTransaction
$mo = Get-UcsTopSystem | Set-UcsTopSystem -Descr $UCSDesc -Owner $UCSOwner -Site $UCSSite -Force
$mo_1 = Get-UcsSvcEp | Set-UcsManagedObject -PropertyMap @{Descr=""; PolicyOwner="local"; } -Force
$mo_1_1 = Get-UcsDns | Set-UcsDns -AdminState "enabled" -Descr "" -Domain $UCSDNSDomain -PolicyOwner "local" -Port 0 -Force
Complete-UcsTransaction
if ($SystemName -ne "")
	{
		Get-UcsTopSystem | Set-UcsTopSystem -Name $SystemName -force	
	}

#Admin - Time Zone Management
Write-Host -ForegroundColor DarkBlue "Set Timezone Management Information"
Get-UcsTimezone | Set-UcsTimezone -AdminState "enabled" -Descr "" -PolicyOwner "local" -Port 0 -Timezone $Timezone -Force
foreach ($ntpnumber in $NTPName)
	{
		if ($ntpnumber -ne "")
			{
				Add-UcsNtpServer -Descr "" -Name $ntpnumber
			}
	}

#Admin - DNS Management
Write-Host -ForegroundColor DarkBlue "Set DNS Management Information"
if ($PriDNS -ne "")
	{
		Add-UcsDnsServer -Descr "" -Name $PriDNS
	}
if ($SecDNS -ne "")
	{
		Add-UcsDnsServer -Descr "" -Name $SecDNS
	}

#LAN - Global Policies - VLAN Compression
$FabricInterconnectVersion = Get-UcsFiModule
if (($FabricInterconnectVersion.SyncRoot.model -contains "UCS-FI-6248UP") -or ($FabricInterconnectVersion.SyncRoot.model -contains "UCS-FI-6296UP"))
	{
		Write-Host -ForegroundColor DarkBlue "Seting VLAN Compression - ON for 6200 series Fabric Interconnects"
		Get-UcsLanCloud | Set-UcsLanCloud -MacAging "mode-default" -Mode "end-host" -VlanCompression "enabled" -Force
	}

# LAN - Build LAN Port Channels
if ($LANPortChannels -ieq "y")
	{
		Get-UcsFiLanCloud -Id "A" | Add-UcsUplinkPortChannel -AdminSpeed "10gbps" -AdminState "enabled" -FlowCtrlPolicy "default" -Name $LANPortChannelAName -OperSpeed "10gbps" -PortId $LANPortChannelANumber
		if ($IsRedundantFI -ne $null)
			{
				Get-UcsFiLanCloud -Id "B" | Add-UcsUplinkPortChannel -AdminSpeed "10gbps" -AdminState "enabled" -FlowCtrlPolicy "default" -Name $LANPortChannelBName -OperSpeed "10gbps" -PortId $LANPortChannelBNumber
			}
		foreach ($up in $UplinkPort)
			{
				Start-UcsTransaction
				$mo = Get-UcsFiLanCloud -Id "A" | Add-UcsUplinkPortChannel -ModifyPresent  -AdminSpeed "10gbps" -AdminState "enabled" -FlowCtrlPolicy "default" -Name $LANPortChannelAName -OperSpeed "10gbps" -PortId $LANPortChannelANumber
				$mo_1 = $mo | Add-UcsUplinkPortChannelMember -AdminState "enabled" -Name "" -PortId $up["Port"] -SlotId $up["Slot"]
				Complete-UcsTransaction
				if ($IsRedundantFI -ne $null)
					{
						Start-UcsTransaction
						$mo = Get-UcsFiLanCloud -Id "B" | Add-UcsUplinkPortChannel -ModifyPresent  -AdminSpeed "10gbps" -AdminState "enabled" -FlowCtrlPolicy "default" -Name $LANPortChannelBName -OperSpeed "10gbps" -PortId $LANPortChannelBNumber
						$mo_1 = $mo | Add-UcsUplinkPortChannelMember -AdminState "enabled" -Name "" -PortId $up["Port"] -SlotId $up["Slot"]
						Complete-UcsTransaction
					}
			}
	}

# Servers - Build Test Service Profile Templates
#Boot from HD
if ($BootFromHD -ieq "y")
	{
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsServiceProfile -AgentPolicyName "" -BiosProfileName "Performance" -BootPolicyName "Boot_From_HD" -Descr "Boot from Local Hard Drives" -DynamicConPolicyName "" -ExtIPPoolName "ext-mgmt" -ExtIPState "pooled" -HostFwPolicyName "default" -IdentPoolName "UUID_Pool" -LocalDiskPolicyName "RAID1" -MaintPolicyName "User_Ack" -MgmtAccessPolicyName "admin" -MgmtFwPolicyName "" -Name "Boot_from_HD" -PolicyOwner "local" -PowerPolicyName $DefaultPowerControl -ScrubPolicyName $DefaultScrub -SolPolicyName $DefaultSoL -SrcTemplName "" -StatsPolicyName "default" -Type "updating-template" -UsrLbl "" -Uuid "0" -VconProfileName ""
		if ($VHBAnameA -ne "")
			{
				$mo_26 = $mo | Add-UcsVnicConnDef -ModifyPresent -LanConnPolicyName $DefaultLANConnectivity -SanConnPolicyName $DefaultHBAConnectivity		
			}
		else
			{
				$mo_26 = $mo | Add-UcsVnicConnDef -ModifyPresent -LanConnPolicyName $DefaultLANConnectivity
			}
		$mo_27 = $mo | Add-UcsVnicDefBeh -ModifyPresent -Action "none" -Descr "" -Name "" -NwTemplName "" -PolicyOwner "local" -Type "vhba"
		$mo_51 = $mo | Add-UcsVnicFcNode -ModifyPresent -Addr "pool-derived" -IdentPoolName "node-default"
		if ($VHBAnameA -ne "")
			{
				$mo_52 = $mo | Add-UcsVhba -AdaptorProfileName "" -Addr "derived" -AdminVcon "any" -IdentPoolName "" -MaxDataFieldSize 2048 -Name $VHBAnameA -NwTemplName "" -Order "1" -PersBind "disabled" -PersBindClear "no" -PinToGroupName "" -QosPolicyName "" -StatsPolicyName "default" -SwitchId "A"
			}
		if ($VHBAnameB -ne "")
			{
				$mo_53 = $mo | Add-UcsVhba -AdaptorProfileName "" -Addr "derived" -AdminVcon "any" -IdentPoolName "" -MaxDataFieldSize 2048 -Name $VHBAnameB -NwTemplName "" -Order "2" -PersBind "disabled" -PersBindClear "no" -PinToGroupName "" -QosPolicyName "" -StatsPolicyName "default" -SwitchId "A"		
			}
		$mo_54 = $mo | Add-UcsServerPoolAssignment -ModifyPresent -Name "All_Servers" -Qualifier "" -RestrictMigration "no"
		$mo_55 = $mo | Set-UcsServerPower -State "admin-up" -Force
		$mo_56 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "1" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
		$mo_57 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "2" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
		$mo_58 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "3" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
		$mo_59 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "4" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
		Complete-UcsTransaction
	}

#Boot from SAN
if ($BootFromSAN -ieq "y")
	{
		foreach ($bm in $BootMatrix)
			{
				$APrimary = $bm["APrimary"]
				$ASecondary = $bm["ASecondary"]
				$BPrimary = $bm["BPrimary"]
				$BSecondary = $bm["BSecondary"]
						if ($VHBAnameB -ne "")
							{
								$Description = "A_Primary:"+$ArrayPort[$APrimary - 1]["Name"]+"_Fabric:"+$ArrayPort[$APrimary - 1]["Fabric"]+"___A_Secondary:"+$ArrayPort[$ASecondary - 1]["Name"]+"_Fabric:"+$ArrayPort[$ASecondary - 1]["Fabric"]+"___B_Primary:"+$ArrayPort[$BPrimary - 1]["Name"]+"_Fabric:"+$ArrayPort[$BPrimary - 1]["Fabric"]+"___B_Secondary:"+$ArrayPort[$BSecondary - 1]["Name"]+"_Fabric:"+$ArrayPort[$BSecondary - 1]["Fabric"]
							}
						else
							{
								$Description = "A_Primary:"+$ArrayPort[$APrimary - 1]["Name"]+"_Fabric:"+$ArrayPort[$APrimary - 1]["Fabric"]+"___A_Secondary:"+$ArrayPort[$ASecondary - 1]["Name"]+"_Fabric:"+$ArrayPort[$ASecondary - 1]["Fabric"]
							}
				Start-UcsTransaction
				$BootFromSANName = "Boot_from_"+$bm["Name"]
				$mo = Get-UcsOrg -Level root  | Add-UcsServiceProfile -AgentPolicyName "" -BiosProfileName "Performance" -BootPolicyName $BootFromSANName -Descr $Description -DynamicConPolicyName "" -ExtIPPoolName "ext-mgmt" -ExtIPState "pooled" -HostFwPolicyName "default" -IdentPoolName "UUID_Pool" -LocalDiskPolicyName "Any_Config" -MaintPolicyName "User_Ack" -MgmtAccessPolicyName "admin" -MgmtFwPolicyName "" -Name $BootFromSANName -PolicyOwner "local" -PowerPolicyName $DefaultPowerControl -ScrubPolicyName $DefaultScrub -SolPolicyName $DefaultSoL -SrcTemplName "" -StatsPolicyName "default" -Type "updating-template" -UsrLbl "" -Uuid "0" -VconProfileName ""
				$mo_28 = $mo | Add-UcsVnicConnDef -ModifyPresent -LanConnPolicyName $DefaultLANConnectivity -SanConnPolicyName $DefaultHBAConnectivity
				$mo_51 = $mo | Add-UcsVnicFcNode -ModifyPresent -Addr "pool-derived" -IdentPoolName "node-default"
				$mo_54 = $mo | Add-UcsServerPoolAssignment -ModifyPresent -Name "All_Servers" -Qualifier "" -RestrictMigration "no"
				$mo_55 = $mo | Set-UcsServerPower -State "admin-up" -Force
				$mo_56 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "1" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
				$mo_57 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "2" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
				$mo_58 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "3" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
				$mo_59 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "4" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
				Complete-UcsTransaction
			}
	}

#Boot from iSCSI
if ($BootFromiSCSI -ieq "y")
	{
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsServiceProfile -AgentPolicyName "" -BiosProfileName "Performance" -BootPolicyName "Boot_from_iSCSI" -Descr "Boot from iSCSI SAN. **NOT AN UPDATING TEMPLATE**" -DynamicConPolicyName "" -ExtIPPoolName "ext-mgmt" -ExtIPState "pooled" -HostFwPolicyName "default" -IdentPoolName "UUID_Pool" -LocalDiskPolicyName "Any_Config" -MaintPolicyName "User_Ack" -MgmtAccessPolicyName "admin" -MgmtFwPolicyName "" -Name "Boot_from_iSCSI" -PolicyOwner "local" -PowerPolicyName $DefaultPowerControl -ScrubPolicyName $DefaultScrub -SolPolicyName $DefaultSoL -SrcTemplName "" -StatsPolicyName "default" -Type "initial-template" -UsrLbl "" -Uuid "0" -VconProfileName ""
		if ($VHBAnameA -ne "")
			{
				$mo_26 = $mo | Add-UcsVnicConnDef -ModifyPresent -LanConnPolicyName $DefaultLANwiSCSIConnectivity -SanConnPolicyName ""
			}
		else
			{
				$mo_26 = $mo | Add-UcsVnicConnDef -ModifyPresent -LanConnPolicyName $DefaultLANwiSCSIConnectivity
			}
		$mo_27 = $mo | Add-UcsVnicDefBeh -ModifyPresent -Action "none" -Descr "" -Name "" -NwTemplName "" -PolicyOwner "local" -Type "vhba"
		$mo_51 = $mo | Add-UcsVnicFcNode -ModifyPresent -Addr "pool-derived" -IdentPoolName "node-default"
		$mo_52 = $mo | Add-UcsVnicIScsi -ModifyPresent -AdaptorProfileName "default" -Addr "derived" -AdminVcon "any" -AuthProfileName "" -ExtIPState "none" -IdentPoolName "" -InitiatorName "" -IqnIdentPoolName "" -Name $iSCSINicNameA -NwTemplName "" -Order "unspecified" -PinToGroupName "" -QosPolicyName "" -StatsPolicyName "default" -SwitchId "A" -VnicName $iSCSINicNameA
		$mo_52_1 = $mo_52 | Add-UcsVnicVlan -Name "" -VlanName "default"
		if ($IsRedundantFI -ne $null)
			{
				$mo_53 = $mo | Add-UcsVnicIScsi -ModifyPresent -AdaptorProfileName "default" -Addr "derived" -AdminVcon "any" -AuthProfileName "" -ExtIPState "none" -IdentPoolName "" -InitiatorName "" -IqnIdentPoolName "" -Name $iSCSINicNameB -NwTemplName "" -Order "unspecified" -PinToGroupName "" -QosPolicyName "" -StatsPolicyName "default" -SwitchId "A" -VnicName $iSCSINicNameB
				$mo_53_1 = $mo_53 | Add-UcsVnicVlan -Name "" -VlanName "default"
			}
		$mo_54 = $mo | Add-UcsServerPoolAssignment -ModifyPresent -Name "All_Servers" -Qualifier "" -RestrictMigration "no"
		$mo_55 = $mo | Set-UcsServerPower -State "admin-up" -Force
		$mo_56 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "1" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
		$mo_57 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "2" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
		$mo_58 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "3" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
		$mo_59 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "4" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
		Complete-UcsTransaction
		Write-Host -ForegroundColor DarkBlue "iSCSI Fabric A Boot Parameters being set"
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "Boot_from_iSCSI" -LimitScope | Add-UcsVnicIScsiBootParams -Descr "" -PolicyOwner "local"
		$mo_1 = $mo | Add-UcsVnicIScsiBootVnic -AuthProfileName "" -Descr "" -InitiatorName "" -IqnIdentPoolName "IQN_Pool" -Name $iSCSINicNameA -PolicyOwner "local"
		$mo_1_1 = $mo_1 | Add-UcsVnicIScsiStaticTargetIf -AuthProfileName "" -IpAddress $iSCSItargetA -Name $iSCSIiqn -Port 3260 -Priority 1
		$mo_1_1_1 = $mo_1_1 | Add-UcsVnicLun -ModifyPresent -Bootable "no" -Id 0
		if ($iSCSIPool -ieq "DHCP")
			{
				$mo_1_2 = $mo_1 | Add-UcsVnicIPv4If -Name ""
				$mo_1_2_1 = $mo_1_2 | Add-UcsVnicIPv4Dhcp
			}
		else
			{
				$mo_1_2 = $mo_1 | Add-UcsVnicIPv4If -Name ""
				$mo_1_2_1 = $mo_1_2 | Add-UcsVnicIPv4Dhcp
				$mo_1_2_2 = $mo_1_2 | Add-UcsManagedObject -ClassId VnicIPv4PooledIscsiAddr -PropertyMap @{IdentPoolName=$DefaultiSCSIInitiatorPoolA; }
			}
		Complete-UcsTransaction
		sleep 1
		if ($IsRedundantFI -ne $null)
			{
				Write-Host -ForegroundColor DarkBlue "iSCSI Fabric B Boot Parameters being set"
				Start-UcsTransaction
				$mo = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "Boot_from_iSCSI" -LimitScope | Get-UcsVnicIScsiBootParams | Set-UcsManagedObject -PropertyMap @{Descr=""; PolicyOwner="local"; } -Force
				$mo_1 = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "Boot_from_iSCSI" -LimitScope | Get-UcsVnicIScsiBootParams | Get-UcsVnicIScsiBootVnic -Name $iSCSINicNameB | Set-UcsVnicIScsiBootVnic -AuthProfileName "" -Descr "" -InitiatorName "" -IqnIdentPoolName "IQN_Pool" -PolicyOwner "local" -Force
				$mo_1_1 = $mo_1 | Add-UcsVnicIScsiStaticTargetIf -AuthProfileName "" -IpAddress $iSCSItargetB -Name $iSCSIiqn -Port 3260 -Priority 1
				$mo_1_1_1 = $mo_1_1 | Add-UcsVnicLun -ModifyPresent -Bootable "no" -Id 0
				if ($iSCSIPool -ieq "DHCP")
					{
						$mo_1_2 = $mo_1 | Add-UcsVnicIPv4If -Name ""
						$mo_1_2_1 = $mo_1_2 | Add-UcsVnicIPv4Dhcp
					}
				else
					{
						$mo_1_2 = $mo_1 | Add-UcsVnicIPv4If -Name ""
						$mo_1_2_1 = $mo_1_2 | Add-UcsVnicIPv4Dhcp
						$mo_1_2_2 = $mo_1_2 | Add-UcsManagedObject -ClassId VnicIPv4PooledIscsiAddr -PropertyMap @{IdentPoolName=$DefaultiSCSIInitiatorPoolB; }
					}
				Complete-UcsTransaction
			}
	}

#SAN - Configure Fibre Channel Ports
if ($FCPort[0].Port -ne "")
	{
		Write-Host -ForegroundColor DarkBlue "Configuring Fibre Channel Ports"
		Start-UcsTransaction
		foreach ($fc in $FCPort)
			{
				Get-UcsFiSanCloud -Id "A" | Add-UcsFcUplinkPort -ModifyPresent  -AdminState "enabled" -Name "" -PortId $fc["Port"] -SlotId $fc["Slot"] -UsrLbl $fc["LabelA"]
				if ($VHBAnameB -ne "")
					{
						Get-UcsFiSanCloud -Id "B" | Add-UcsFcUplinkPort -ModifyPresent  -AdminState "enabled" -Name "" -PortId $fc["Port"] -SlotId $fc["Slot"] -UsrLbl $fc["LabelB"]
					}
			}
		Complete-UcsTransaction

		if ($VHBAnameB -ne "")
					{
						Write-Host -ForegroundColor DarkBlue "Changing Unified Ports forces a reboot of the Fabric Interconnects or the expansion module."
						if ($fc["Slot"] -eq "1")
							{
								Write-Host -ForegroundColor White -BackgroundColor DarkBlue "	Fabric Interconnects are rebooting...Please wait...(This can take 10 or so minutes)"
								$FIReboot = 1
								do 
									{
										Disconnect-Ucs
										if ($UCSEmulator -ine "y")
											{
												Sleep 60
											}
										else
											{
												Sleep 5
											}
										do
											{
												Write-Host -ForegroundColor DarkBlue "Checking to see if UCSM VIP is active"
												$ping = new-object system.net.networkinformation.ping
												$results = $ping.send($myucs)
												if ($results.Status -ne "Success")
													{
														Write-Host -ForegroundColor DarkBlue "	Not Yet, Waiting..."
														if ($UCSEmulator -ine "y")
															{
																Sleep 60
															}
														else
															{
																Sleep 5
															}
														$FIReboot++
														if ($FIReboot -ge 15)
															{
																Write-Host ""
																Write-Host -ForegroundColor Red "The Fabric Interconnects have failed to reload in the 15 minute window"
																Write-Host -ForegroundColor Red "			Script did not complete all configuration items:"
																Write-Host -ForegroundColor Red "				VSANS to SAN Uplink Ports"
																Write-Host -ForegroundColor Red "				LAN Port Channels (If Used)"
																Write-Host -ForegroundColor Red "				SAN Port Channels (If Used)"
																Write-Host -ForegroundColor Red "				Service Profile Templates"
																Write-Host -ForegroundColor Red "				Customization Script (If Used)"
																Write-Host -ForegroundColor Red "					Exiting..."
																Disconnect-Ucs
																exit
															}
													}
												else
													{
														Write-Host -ForegroundColor DarkBlue "  Waiting for Fabric Interconnect to be fully up (~2 minutes)"
														if ($UCSEmulator -ine "y")
															{
																Sleep 120
															}
														else
															{
																Sleep 5
															}
														Write-Host -ForegroundColor DarkBlue "	Logging back into UCSM"
														$myCon = Connect-Ucs $myucs -Credential $cred
														if (($myucs | Measure-Object).count -ne ($myCon | Measure-Object).count) 
															{
																#Exit Script
																Write-Host -ForegroundColor Red "		Error Re-Logging into UCSM"
																Write-Host -ForegroundColor Red "			Script did not complete all configuration items:"
																Write-Host -ForegroundColor Red "				VSANS to SAN Uplink Ports"
																Write-Host -ForegroundColor Red "				LAN Port Channels (If Used)"
																Write-Host -ForegroundColor Red "				SAN Port Channels (If Used)"
																Write-Host -ForegroundColor Red "				Service Profile Templates"
																Write-Host -ForegroundColor Red "				Customization Script (If Used)"
																Write-Host -ForegroundColor Red "					Exiting..."
																Disconnect-Ucs
																exit
															}
														else
															{
																Write-Host -ForegroundColor DarkGreen "		Login Successful"
															}
													}
											}
										until ($results.Status -eq "Success")
									}
								until ((($FIStatus = Get-UcsManagedObject -Dn "sys/switch-A").Operability -eq "operable") -and ($FIStatus = Get-UcsManagedObject -Dn "sys/switch-B").Operability -eq "operable")
								Write-Host -ForegroundColor DarkGreen "			Reboot is Complete"
								Write-Host ""
							}
					}
				else
					{
						Write-Host -ForegroundColor DarkBlue "Changing Unified Ports forces a reboot of the Fabric Interconnects or the expansion module."
						if ($fc["Slot"] -eq "1")
							{
								Write-Host -ForegroundColor White -BackgroundColor DarkBlue "	Fabric Interconnects are rebooting...Please wait...(This can take 10 or so minutes)"
								$FIReboot = 1
								do 
									{
										Disconnect-Ucs
										if ($UCSEmulator -ine "y")
											{
												Sleep 60
											}
										else
											{
												Sleep 5
											}
										do
											{
												Write-Host -ForegroundColor DarkBlue "Checking to see if UCSM VIP is active"
												$ping = new-object system.net.networkinformation.ping
												$results = $ping.send($myucs)
												if ($results.Status -ne "Success")
													{
														Write-Host -ForegroundColor DarkBlue "	Not Yet, Waiting..."
														if ($UCSEmulator -ine "y")
															{
																Sleep 60
															}
														else
															{
																Sleep 5
															}
														$FIReboot++
														if ($FIReboot -ge 15)
															{
																Write-Host ""
																Write-Host -ForegroundColor Red "The Fabric Interconnects have failed to reload in the 15 minute window"
																Write-Host -ForegroundColor Red "			Script did not complete all configuration items:"
																Write-Host -ForegroundColor Red "				VSANS to SAN Uplink Ports"
																Write-Host -ForegroundColor Red "				LAN Port Channels (If Used)"
																Write-Host -ForegroundColor Red "				SAN Port Channels (If Used)"
																Write-Host -ForegroundColor Red "				Service Profile Templates"
																Write-Host -ForegroundColor Red "				Customization Script (If Used)"
																Write-Host -ForegroundColor Red "					Exiting..."
																Disconnect-Ucs
																exit
															}
													}
												else
													{
														Write-Host -ForegroundColor DarkBlue "  Waiting for Fabric Interconnect to be fully up (~2 minutes)"
														if ($UCSEmulator -ine "y")
															{
																Sleep 120
															}
														else
															{
																Sleep 5
															}
														Write-Host -ForegroundColor DarkBlue "	Logging back into UCSM"
														$myCon = Connect-Ucs $myucs -Credential $cred
														if (($myucs | Measure-Object).count -ne ($myCon | Measure-Object).count) 
															{
																#Exit Script
																Write-Host -ForegroundColor Red "		Error Re-Logging into UCSM"
																Write-Host -ForegroundColor Red "			Script did not complete all configuration items:"
																Write-Host -ForegroundColor Red "				VSANS to SAN Uplink Ports"
																Write-Host -ForegroundColor Red "				LAN Port Channels (If Used)"
																Write-Host -ForegroundColor Red "				SAN Port Channels (If Used)"
																Write-Host -ForegroundColor Red "				Service Profile Templates"
																Write-Host -ForegroundColor Red "				Customization Script (If Used)"
																Write-Host -ForegroundColor Red "					Exiting..."
																Disconnect-Ucs
																exit
															}
														else
															{
																Write-Host -ForegroundColor DarkGreen "		Login Successful"
															}
													}
											}
										until ($results.Status -eq "Success")
									}
								until (($FIStatus = Get-UcsManagedObject -Dn "sys/switch-A").Operability -eq "operable")
								Write-Host -ForegroundColor DarkGreen "			Reboot is Complete"
								Write-Host ""
							}
						elseif ($fc["Slot"] -gt "1")
							{
								if ($VHBAnameB -ne "")
									{
										$ModuleReboot = 1
										Write-Host -ForegroundColor White -BackgroundColor DarkBlue "	Expansion Module is rebooting...Please wait..."
										do 
											{
												Sleep 5
												Write-Host -ForegroundColor DarkBlue "	Waiting..."
												$ExpansionSlotA = "sys/switch-A/slot-"+$fc["slot"]
												$ExpansionSlotB = "sys/switch-B/slot-"+$fc["slot"]
												$ExpansionSlotDNa = Get-UcsManagedObject -Dn $ExpansionSlotA
												$ExpansionSlotDNb = Get-UcsManagedObject -Dn $ExpansionSlotB
												if ($ModuleReboot -ge 120)
													{
														Write-Host ""
														Write-Host -ForegroundColor Red "The Modules have failed to reload in the 5 minute window"
														Write-Host -ForegroundColor Red "			Script did not complete all configuration items:"
														Write-Host -ForegroundColor Red "				VSANS to SAN Uplink Ports"
														Write-Host -ForegroundColor Red "				LAN Port Channels (If Used)"
														Write-Host -ForegroundColor Red "				SAN Port Channels (If Used)"
														Write-Host -ForegroundColor Red "				Service Profile Templates"
														Write-Host -ForegroundColor Red "				Customization Script (If Used)"
														Write-Host -ForegroundColor Red "					Exiting..."
														Disconnect-Ucs
														exit
													}
												$ModuleReboot++
											} 
										until (($ExpansionSlotDNa.Operability -eq "operable") -and ($ExpansionSlotDNb.Operability -eq "operable"))
										Write-Host -ForegroundColor DarkGreen "	Complete"
										Write-Host ""
									}
							}
						else
							{
								$ModuleReboot = 1
								Write-Host -ForegroundColor White -BackgroundColor DarkBlue "	Expansion Module is rebooting...Please wait..."
								do 
									{
										Sleep 5
										Write-Host -ForegroundColor DarkBlue "	Waiting..."
										$ExpansionSlotA = "sys/switch-A/slot-"+$fc["slot"]
										$ExpansionSlotDNa = Get-UcsManagedObject -Dn $ExpansionSlotA
										if ($ModuleReboot -ge 120)
											{
												Write-Host ""
												Write-Host -ForegroundColor Red "The Modules have failed to reload in the 5 minute window"
												Write-Host -ForegroundColor Red "			Script did not complete all configuration items:"
												Write-Host -ForegroundColor Red "				VSANS to SAN Uplink Ports"
												Write-Host -ForegroundColor Red "				LAN Port Channels (If Used)"
												Write-Host -ForegroundColor Red "				SAN Port Channels (If Used)"
												Write-Host -ForegroundColor Red "				Service Profile Templates"
												Write-Host -ForegroundColor Red "				Customization Script (If Used)"
												Write-Host -ForegroundColor Red "					Exiting..."
												Disconnect-Ucs
												exit
											}
										$ModuleReboot++
									} 
								until ($ExpansionSlotDNa.Operability -eq "operable")
								Write-Host -ForegroundColor DarkGreen "	Complete"
								Write-Host ""
							}	
					}
					
		#Put FC ports into appropriate vSAN
		Write-Host -ForegroundColor DarkBlue "Attaching vSAN to each SAN Port"
		foreach ($fc in $FCPort)
			{
				Get-UcsFiSanCloud -Id "A" | Get-UcsVsan -Name $vSANNameA | Add-UcsVsanMemberFcPort -ModifyPresent  -AdminState "enabled" -Name "" -PortId $fc["Port"] -SlotId $fc["Slot"] -SwitchId "A"
				if ($VHBAnameB -ne "")
					{
						Get-UcsFiSanCloud -Id "B" | Get-UcsVsan -Name $vSANNameB | Add-UcsVsanMemberFcPort -ModifyPresent  -AdminState "enabled" -Name "" -PortId $fc["Port"] -SlotId $fc["Slot"] -SwitchId "B"
					}
			}
	}

# SAN - Build SAN Port Channels
if ($SANPortChannels -ieq "y")
	{
		Get-UcsFiSanCloud -Id "A" | Add-UcsFcUplinkPortChannel -AdminSpeed "auto" -AdminState "enabled" -Name $SANPortChannelAName -PortId $SANPortChannelANumber
		Get-UcsFiSanCloud -Id "A" | Get-UcsVsan -Name $vSANNameA | Add-UcsVsanMemberFcPortChannel -ModifyPresent  -AdminState "enabled" -Name "" -PortId $SANPortChannelANumber -SwitchId "A"
		if ($VHBAnameB -ne "")
			{
				Get-UcsFiSanCloud -Id "B" | Add-UcsFcUplinkPortChannel -AdminSpeed "auto" -AdminState "enabled" -Name $SANPortChannelBName -PortId $SANPortChannelBNumber
				Get-UcsFiSanCloud -Id "B" | Get-UcsVsan -Name $vSANNameB | Add-UcsVsanMemberFcPortChannel -ModifyPresent  -AdminState "enabled" -Name "" -PortId $SANPortChannelBNumber -SwitchId "B"
			}
		foreach ($fc in $FCPort)
			{
				Start-UcsTransaction
				$mo = Get-UcsFiSanCloud -Id "A" | Add-UcsFcUplinkPortChannel -ModifyPresent  -AdminSpeed "auto" -AdminState "enabled" -Name $SANPortChannelAName -PortId $SANPortChannelANumber
				$mo_1 = $mo | Add-UcsFabricFcSanPcEp -AdminSpeed "auto" -AdminState "enabled" -Name "" -PortId $fc["Port"] -SlotId $fc["Slot"]
				Complete-UcsTransaction
				if ($VHBAnameB -ne "")
					{
						Start-UcsTransaction
						$mo = Get-UcsFiSanCloud -Id "B" | Add-UcsFcUplinkPortChannel -ModifyPresent  -AdminSpeed "auto" -AdminState "enabled" -Name $SANPortChannelBName -PortId $SANPortChannelBNumber
						$mo_1 = $mo | Add-UcsFabricFcSanPcEp -AdminSpeed "auto" -AdminState "enabled" -Name "" -PortId $fc["Port"] -SlotId $fc["Slot"]
						Complete-UcsTransaction
					}
			}
	}
<#
# Create vNIC/vHBA Placement Policy
$PolicyName = "Boot_vHBA-First"
Start-UcsTransaction
$mo = Get-UcsOrg -Level root  | Add-UcsPlacementPolicy -Descr "Placing the boot vHBA first will allow for ISO builds of server with boot from SAN and multi-paths" -MezzMapping "linear-ordered" -Name $PolicyName -PolicyOwner "local"
$mo_1 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "1" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
$mo_2 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "2" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
$mo_3 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "3" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
$mo_4 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -Id "4" -InstType "auto" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc"
Complete-UcsTransaction

#Modify each Service Profile Template for proper vHBA/vNIC Placcement
$ServiceProfileTemplates = Get-UcsServiceProfile | where {$_.Type -ne "instance"}
foreach ($ServiceProfile in $ServiceProfileTemplates)
	{
		Start-UcsTransaction
		$mo = Get-UcsOrg -Level root  | Add-UcsServiceProfile -ModifyPresent  -VconProfileName $PolicyName -Name $ServiceProfile.Name
		$mo_2 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -InstType "policy" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc" -Id "1"
		$mo_3 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -InstType "policy" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc" -Id "2"
		$mo_4 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -InstType "policy" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc" -Id "3"
		$mo_5 = $mo | Add-UcsFabricVCon -ModifyPresent -Fabric "NONE" -InstType "policy" -Placement "physical" -Select "all" -Share "shared" -Transport "ethernet","fc" -Id "4"
		Complete-UcsTransaction
		$BootPolicyPrimary = Get-UcsBootPolicy -Name $ServiceProfile.BootPolicyName -Hierarchy | where {$_.Rn -eq "san-primary"}
		$BootPolicySecondary = Get-UcsBootPolicy -Name $ServiceProfile.BootPolicyName -Hierarchy | where {$_.Rn -eq "san-secondary"}
		$FirstvHBA = $BootPolicyPrimary.VnicName
		$SecondvHBA = $BootPolicySecondary.VnicName
		if ($FirstvHBA)
			{
				$ServiceProfile | Add-UcsManagedObject -ModifyPresent -ClassId LsVConAssign -PropertyMap @{VnicName=$FirstvHBA; AdminVcon="1"; Order="1"; Transport="fc"; }
			}
		if ($SecondvHBA)
			{
				$ServiceProfile | Add-UcsManagedObject -ModifyPresent -ClassId LsVConAssign -PropertyMap @{VnicName=$SecondvHBA; AdminVcon="2"; Order="1"; Transport="fc"; }
			}
		$vNICs = Get-UcsLsVConAssign -ServiceProfile $ServiceProfile.Name | where {$_.Transport -eq "ethernet"}
		$Order = 1
		foreach ($vNIC in $vNICs)
			{
				$ServiceProfile | Add-UcsManagedObject -ModifyPresent -ClassId LsVConAssign -PropertyMap @{VnicName=$vNIC.VnicName; AdminVcon="3"; Order=$Order; Transport="ethernet"; }
				$Order ++
			}
	}
#>

#Run any special customization Scripts
if ($CustomScript -ne $null)
	{
		Write-Host ""
		Write-Host -ForegroundColor DarkBlue "Executing customization Script"
		Disconnect-Ucs
		$FullCustomScript = ".\"+$CustomScript
		& $FullCustomScript
		Write-Host -ForegroundColor DarkGreen "		Custom Script Complete"
	}
Write-Host ""
Write-Host -ForegroundColor DarkGreen "Done with UCS base configuration Build"

#end script
##Disconnect from UCS
Write-Host -ForegroundColor DarkBlue "Disconnecting from UCS"
Disconnect-Ucs

##Notify user that script is complete
Write-Host -ForegroundColor White -BackgroundColor DarkBlue "Script Completing..."
Write-Host ""
Write-Host -ForegroundColor DarkBlue "There are still many things to do to complete your UCS setup that this tool has not completed:"
Write-Host -ForegroundColor DarkBlue "	1 - Admin Tab has many things that could be used.  Local users, LDAP, TACACS, RADIUS, Smart Call Home, etc."
Write-Host -ForegroundColor DarkBlue "	2 - Verify that the default Host Firmware Packages (default) is using the version of UCSM that you expect/want"
Write-Host -ForegroundColor DarkBlue "	4 - If using iSCSI boot, modify boot parameters to match the target address"
Write-Host -ForegroundColor DarkBlue "	4 - UCS has MANY customizable options.  Please go through the system and make sure everything is setup and working as expected"
Write-Host ""
Write-Host -ForegroundColor DarkBlue "IF SOMETHING DID NOT BUILD AS EXPECTED IT COULD BE DUE TO AN INVALID FIELD SUCH AS:"
Write-Host -ForegroundColor DarkBlue "	Invalid characters in the field"
Write-Host -ForegroundColor DarkBlue "	Field Length too long"
Write-Host -ForegroundColor DarkBlue "	Not providing all required parameters"
Write-Host ""
Write-Host -ForegroundColor DarkBlue "	I have tried my best to catch all these possible issues before proceeding so let me know what failed and I will correct this in my code"
Write-Host ""
Write-Host -ForegroundColor White -BackgroundColor DarkMagenta "*******************************************SPECIAL NOTES FOR THIS DEPLOYMENT*******************************************"
Write-Host $SpecialNotes -ForegroundColor DarkMagenta
Write-Host ""

}

Catch
{
Write-Host ""
Write-Host -ForegroundColor Red "WARNING WARNING WARNING-"
Write-Host -ForegroundColor Red "An error occured in the script and needs to be corrected before it will make further changes to UCS"
Write-Host ""
Write-Host -ForegroundColor Red "Error found is:"
Write-Host -ForegroundColor Red "" -NoNewline
$Error[0]
}

Finally
{
Write-Host ""
Write-Host -ForegroundColor White -BackgroundColor DarkBlue "Script Complete..."
Write-Host -ForegroundColor White -BackgroundColor DarkBlue "	Exiting..."

#Unload all modules
Remove-Module Data-*

#Disconnect from UCS
Disconnect-Ucs

#Exit Script
exit
}