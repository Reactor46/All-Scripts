<#

.SYNOPSIS
	This script logs into multiple UCS domains and lists the blades in each slot.  It also identifies empty slots.

.DESCRIPTION
	This script logs into multiple UCS domains and lists the blades in each slot.  It also identifies empty slots.  It benefits from labeling each chassis' User Label.

.EXAMPLE
	Find-UcsAvailableSlots.ps1
	This script can be run without any command line parameters.  User will be prompted for all parameters and options required

.EXAMPLE
	Find-UcsAvailableSlots.ps1 -ucs "1.2.3.4,5,6,7,8" -ucred
	-ucs -- UCS Manager IP or Host Name -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local" or "1.2.3.4,5.6.7.8" or "myucs1,myucs2" or "myucs1.domain.local,myucs2.domain.local"
	-ucred -- UCS Manager Credential Switch -- Adding this switch will immediately prompt you for your UCSM username and password
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords
	
.EXAMPLE
	Find-UcsAvailableSlots.ps1 -ucs "1.2.3.4,5,6,7,8" -saved "myucscred.csv" -skiperrors
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
	Version: v0.7.03
	Date: 4/21/2015
	Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
	UCSM IP Address(s) or Hostname(s)
	UCSM Username and Password
	UCSM Credentials Filename

.OUTPUTS
	Text file listing output report named: Available Slots for xxxxxxxx.txt
	
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

#Clear the Screen
Clear-Host

#Script kicking off
Write-Output "Script Running..."
Write-Output ""

#Do not show errors in script
$ErrorActionPreference = "SilentlyContinue"
#$ErrorActionPreference = "Stop"
#$ErrorActionPreference = "Continue"
#$ErrorActionPreference = "Inquire"

#Gather any credentials requested from command line
if ($UCREDENTIALS)
	{
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

#Tell the user what this script does
Write-Output "This script logs into multiple UCS domains and lists the blades in each slot."
Write-Output "It also identifies empty slots."
Write-Output ""
Write-Output "This script requires that the login for each UCS domain is the same and that"
Write-Output "the user has rights to pull inventory information"
Write-Output ""
Write-Output "This script requires to you enter a list of your UCS Domain IPs/Hostnames"
Write-Output "and provide the file output path.  These are found towards"
Write-Output "the top of the script"
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
		Write-Output ""
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


#Save the file
Write-Output ""
[string]$ConnectedUCS = $myucs
if ($ConnectedUCS.Length -ge 234)
	{
		$UCSList = $ConnectedUCS.substring(0,220)
	}
else
	{
		$UCSList = $ConnectedUCS
	}
$file = $PSScriptRoot + "\Available Slots for "+$UCSList+".txt"
del $file -Force
Write-Output "The file will be created as:"
Write-Output "	$file"
Write-Output ""

foreach ($myucslist in $myucs)
	{
		Write-Output "logging into: $myucslist"
		Out-File -FilePath $file -append -InputObject "logging into: $myucslist"
		$myCon = $null
		$myCon = Connect-Ucs $myucslist -Credential $cred
		if (($mycon).Name -ne ($myucslist)) 
			{
				#Exit Script
				Write-Output "     Error Logging into this UCS domain"
				Out-File -FilePath $file -Append -InputObject "     Error Logging into this UCS domain"
				if ($SKIPERROR)
					{
						$continue = "y"
					}
				else
					{
						$continue = Read-Host "Continue without this UCS domain (Y/N)"
					}
				if ($continue -ieq "n")
					{
						Write-Output "Exiting Script..."
						Out-File -FilePath $file -Append -InputObject "Exiting Script..."
						Disconnect-Ucs
						exit
					}
				else
					{
						Write-Output "		Continuing..."
						Out-File -FilePath $file -Append -InputObject "		Continuing..."
						Write-Output ""
						Out-File -FilePath $file -Append -InputObject ""
						Write-Output "----------------------------------------------------------"
						Write-Output ""
						$output = "----------------------------------------------------------"
						Out-File -FilePath $file -Append -InputObject $output
						Out-File -FilePath $file -Append -InputObject ""
					}
			}
		else
			{
				Write-Output "     Login Successful"
				Out-File -FilePath $file -Append -InputObject "     Login Successful"
				Write-Output ""
				Write-Output "Generating report...Please Wait..."
				Write-Output ""
				##Reset previously used values
				$chassis = $null
				$blade = $null
				$ch = $null
				$bladeID = $null
				$slot = $null
				$output = $null
				$ucs = $null
				$version = $null
				$description = $null
				$userName = $null
				$chassisNumber = $null
				$chassisUserLabel = $null
				##Get information about the UCS Domain
				$ucs = $myCon.Ucs
				$admin = Get-UcsTopSystem
				$ucsSite = $admin.Site
				$ucsOwner = $admin.Owner
				if ($ucsSite -eq "")
					{
						$ucsSite = "UNASSIGNED"
					}
				if ($ucsOwner -eq "")
					{
						$ucsOwner = "UNASSIGNED"
					}
				Out-File -FilePath $file -Append -InputObject ""
				$version = $myCon.Version
				Out-File -FilePath $file -Append -InputObject "UCS Domain: $ucs, UCS Version: $version, Site: $ucsSite, Owner: $ucsOwner"
				Out-File -FilePath $file -Append -InputObject ""
				##Get list of chassis in this domain				
				$chassis = Get-UcsChassis
				##Loop through each chassis
				foreach ($ch in $chassis)
					{
						##Provide Chassis Information
						$chassisNumber = $ch.Id
						$chassisUserLabel = $ch.UsrLbl
						if ($chassisUserLabel -ne "")
							{
								Out-File -FilePath $file -Append -InputObject "Chassis: $chassisNumber, User Label: $chassisUserLabel"
								Out-File -FilePath $file -Append -InputObject ""

							}
						else
							{
								Out-File -FilePath $file -Append -InputObject "Chassis: $chassisNumber, User Label: UNASSIGNED"
								Out-File -FilePath $file -Append -InputObject ""
							}
						##Check each slot in the chassis (1-8)
						$slot = 1
						do
							{
								##Get blade information for the slot
								$blade = Get-UcsBlade | where {($_.chassisid -eq $ch.Id) -and ($_.SlotId -eq $slot)}
								$bladeID = $blade.Model
								$chassisID =$ch.id
								$description = $blade.UsrLbl
								$userName = $blade.Name
								$output = "Chassis: $chassisID, Slot: $slot, Model: $bladeID"
								if (($description -ne "") -and ($bladeID -ne $null))
									{
										$output += ", User Label: $description"
									}
								if (($userName -ne "") -and ($bladeID -ne $null))
									{
										$output += ", Name: $userName"
									}
								##Look to see if the blade is single-wide
								if (("B230-BASE-M2", "N20-B6620-1", "N20-B6625-1", "N20-B6730-1", "N20-B6730-1", "UCSB-B200-M3", "UCSB-B22-M3", "UCSB-B200-M4") -icontains $bladeID)
									{
										$output += ", Single Wide Blade"
										Out-File -FilePath $file -Append -InputObject $output
										$slot += 1
										$blade = $null
										$bladeID =$null
									}
								##Look to see if the blade is double-wide.  If it is then don't look at the next slot as it's full
								elseif (("B440-BASE-M2", "N20-B6620-2", "N20-B6625-2", "N20-B6740-2", "UCSB-B420-M3", "UCSB-EX-M4-1") -icontains $bladeID)
									{
										$output += ", Double Wide Blade"
										Out-File -FilePath $file -Append -InputObject $output										
										$slot += 2
										$blade = $null
										$bladeID = $null
									}
								##If I don't get any blade information back then the slot is empty
								elseif ($bladeID -eq $null)
									{
										$output += ">>>---THIS SLOT IS EMPTY---<<<"
										Out-File -FilePath $file -Append -InputObject $output										
										$slot += 1
										$blade = $null
										$bladeID = $null
									}
								##If I get blade info back but it doesn't match any of the PIDs listed above then it's a new type of blade and this script needs to be upgraded to include it in the single-wide or double-wide section above
								else
									{
										$output += "!!!!!!ERROR: UNDEFINED BLADE TYPE!!!!!!"
										Out-File -FilePath $file -Append -InputObject $output
										$slot += 1
										$blade = $null
										$bladeID = $null
									}
							}
						while
							($slot -lt 9)
						$output = ""
						Out-File -FilePath $file -Append -InputObject $output
					}
				$output = "----------------------------------------------------------"
				Out-File -FilePath $file -Append -InputObject $output
				Out-File -FilePath $file -Append -InputObject ""
				Disconnect-Ucs
			}
	}
	
##Disconnect from all UCS domains upon completion
Write-Output "Disconnecting from all UCS domains"
Out-File -FilePath $file -Append -InputObject "Disconnecting from all UCS domains"
Disconnect-Ucs

##Exit the script
Write-Output ""
Out-File -FilePath $file -Append -InputObject ""
Write-Output "Script Complete"
Out-File -FilePath $file -Append -InputObject "Script Complete"
Write-Output "     Exiting..."
Out-File -FilePath $file -Append -InputObject "     Exiting..."
exit