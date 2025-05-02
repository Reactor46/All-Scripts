<#

.SYNOPSIS
	This script is designed to rename the NICs on a newly configured Windows server to match the names configured in UCSM.

.DESCRIPTION
	This script is designed to rename the NICs on a newly configured Windows server to match the names configured in UCSM.

.EXAMPLE
	Set-UcsWindowsNicNames.ps1
	This script can be run without any command line parameters.  User will be prompted for all parameters and options required

.EXAMPLE
	Set-UcsWindowsNicNames.ps1 -ucs "1.2.3.4" -ucred -server "5.6.7.8" -firewall -scred
	-ucs -- UCS Manager IP or Host Name -- Example: 1.2.3.4 or myucs or myucs.domain.local
	-ucred -- UCS Manager Credential Switch -- Adding this switch will immediately prompt you for your UCSM username and password
	-server -- Server IP or Host Name to rename NICs on -- Example: 5.6.7.8 or myserver or myserver.domain.local
	-firewall -- Acknowledge that the firewall is disabled or a rule is set to allow WMI calls to rename NICs -- Valid option is to include the switch "-firewall" which means the firewall is disabled or not
	-scred -- Server Credential Switch -- Adding this switch will immediately prompt you for your Windows Server administrative username and password
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords
	
.EXAMPLE
	Set-UcsWindowsNicNames.ps1 -ucs "1.2.3.4" -usaved "myucscred.csv" -server "5.6.7.8" -firewall -ssaved "myservercred.csv" -skiperror
	-ucs -- UCS Manager IP or Host Name -- Example: 1.2.3.4 or myucs or myucs.domain.local
	-savedcred -- UCSM credentials file -- Example: -savedcred "myucscred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\myucscred.csv
		Make sure the password file is located in the same folder as the script
	-server -- Server IP or Host Name to rename NICs on -- Example: 5.6.7.8 or myserver or myserver.domain.local
	-firewall -- Acknowledge that the firewall is disabled or a rule is set to allow WMI calls to rename NICs -- Valid option is to include the switch "-firewall" which means the firewall is disabled or not
	-ssavedcred -- Server credentials file -- Example: -savedcred "myservercred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\myservercred.csv
		Make sure the password file is located in the same folder as the script
	-skiperrors -- Tells the script to skip any prompts for errors and continues with 'y'
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords
	
.NOTES
	Author: Joe Martin
	Email: joemar@cisco.com
	Company: Cisco Systems, Inc.
	Version: v0.8.03
	Date: 7/22/2014
	Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
	UCSM IP Address or Hostname
	UCSM Username and Password
	UCSM Credentials File
	Server to rename the NICs
	Firewall switch to identify if the Windows firewall is disabled
	Server administrative Username and Password
	Server Credentials File

.OUTPUTS
	None
	
.LINK
	http://communities.cisco.com/people/joemar/content

#>

#Command Line Parameters
param(
	[string]$UCSM,				# UCSM IP Address or Hostname
	[switch]$UCREDENTIALS,		# UCSM Credentials (Username and Password)
	[string]$USAVEDCRED,		# Saved UCSM Credentials.  To create do: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} | Export-CSV -NoTypeInformation .\myucscred.csv
	[string]$SERVER,			# Server IP Address or Hostname
	[switch]$FIREWALL,			# Disable firewall switch
	[switch]$SCREDENTIALS,		# Server Credentials (Username and Password)
	[string]$SSAVEDCRED,		# Saved Server Credentials.  To create do: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} | Export-CSV -NoTypeInformation .\myservercred.csv
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
else
	{
		Write-Output ""
	}
if ($SCREDENTIALS)
	{
		Write-Output ""
		Write-Output "Enter Server Credentials"
		$creds = Get-Credential -Message "Enter Server Credentials"
		Write-Output ""
	}

#Change directory to the script root
cd $PSScriptRoot

#Check to see if UCSM credential files exists
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

#Check to see if Server credential files exists
if ($SSAVEDCRED)
	{
		if ((Test-Path $SSAVEDCRED) -eq $false)
			{
				Write-Output ""
				Write-Output "Your credentials file $SSAVEDCRED does not exist in the script directory"
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
Write-Output "Overview:"
Write-Output "	This script is designed to rename the NICs on a newly configured Windows"
Write-Output "	server to match the names configured in UCSM"
Write-Output ""
Write-Output "Prerequisites:"
Write-Output " Administrative access to UCSM"
Write-Output "	Administrative control of the server to have NICs renamed"
Write-Output "	Disabling the Windows Firewall on the server to have it's NICs renamed"
Write-Output "	or build a rule to allow remote WMI"
Write-Output ""
Write-Output "Details:"
Write-Output "	The script will rename all NICs controlled by UCSM"
Write-Output "	This script will rename Hyper-V vSwitches"
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

#Check and/or load Hyper-V Module for PowerShell
$Modules = Get-Module
$ModuleLoaded = $modules.name
Write-Output ""
Write-Output "Checking to see if Hyper-V module is loaded"
if ( -not ($Modules -like "Hyper-V"))
	{
		Write-Output "     Loading Module: Hyper-V"
		Import-Module Hyper-V
	}
else
	{
		Write-Output "     Hyper-V module is already loaded"
	}

#Select UCS Domain(s) for login
if (!$UCSM)
	{
		$myucs = Read-Host "Enter UCS system IP or Hostname"
	}
else
	{
		$myucs = $UCSM
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
$multilogin = Set-UcsPowerToolConfiguration -SupportMultipleDefaultUcs $false
Write-Output ""
Write-Output "Logging into UCS"
#Verify PowerShell Version to pick prompt type
$PSVersion = $psversiontable.psversion
$PSMinimum = $PSVersion.Major
if (!$UCREDENTIALS)
	{
		if (!$USAVEDCRED)
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
				$CredFile = import-csv $USAVEDCRED
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

#Enter the IP Address or Hostname of the server associated with the UCSM Service Profile
if (!$SERVER)
	{
		Write-Output ""
		Write-Output "IP Address or Hostname of the Server associated to the UCSM Service Profile above to update NIC names"
		$hostname = Read-Host "Enter the IP address of the server to modify"
	}
else
	{
		$hostname = $SERVER
	}

#Prompt User to disable the firewall on the remote server so WMI will work
if (!$FIREWALL)
	{
		Write-Output ""
		Write-Output "***Disable Windows Firewall on server $hostname before continuing or set a rule to allow WMI or the rename will fail***"
		Write-Output "						***Enter EXIT to cancel the script***"
		$proceed = Read-Host "Press OK once you have disabled the server $hostname`'s firewall or type EXIT to cancel the script"
		if ($proceed -ieq "exit")
			{
				Disconnect-Ucs
				Write-Output "Exit selected"
				Write-Output "	Exiting..."
				Disconnect-Ucs
				exit
			}
	}

#Test that the server is IP Reachable via Ping
$ping = new-object system.net.networkinformation.ping
$results = $ping.send($hostname)
if ($results.Status -ne "Success")
	{
		Write-Output ""
		Write-Output "	Can not access server $hostname by Ping"
		Disconnect-Ucs
		Write-Output "Exiting..."
		exit
	}
else
	{
		Write-Output ""
		Write-Output "Reachability to server $hostname is functioning via Ping"
	}
	
#Enter the administrator credentials for the server
if (!$SCREDENTIALS)
	{
		if (!$SSAVEDCRED)
			{
				if ($PSMinimum -ge "3")
					{
						Write-Output ""
						Write-Output "Gathering administrative credentials for the server"
						$PCCred = Get-Credential "administrator" -Message "Enter the login credentials to the server"
					}
				else
					{
						Write-Output ""
						Write-Output "Gathering administrative credentials for the server"
						$PCCred = Get-Credential
					}
			}
		else
			{
				$PCCredFile = import-csv $SSAVEDCRED
				$PCUsername = $PCCredfile.UserName
				$PCPassword = $PCCredfile.EncryptedPassword
				$PCCred = New-Object System.Management.Automation.PsCredential $PCUsername,(ConvertTo-SecureString $PCPassword)			
			}
	}
else
	{
		$PCCred = $creds
	}

#Verify Hyper-V service is installed and running
Write-Output ""
Write-Output "Checking to see if the Hyper-V Service is installed and running"
$Services = Get-WmiObject -ComputerName $hostname -Credential $PCCred -Class Win32_Process
$HyperVService = $Services.Name
if ($HyperVService -ccontains "vmms.exe")
	{
		Write-Output "	Hyper-V Service is installed and running"
	}
else
	{
		Write-Output "	Hyper-V Service is not installed or is not running"
	}

#Getting Service Profile from UUID of Server
Write-Output ""
Write-Output "Getting UCS Service Profile from Server UUID...Please Wait"
$uuid = (Get-WmiObject Win32_ComputerSystemProduct -ComputerName $hostname -Credential $PCCred).UUID 

##Find the service profile that has that UUID
$sp = Get-UcsServiceProfile
foreach ($splist in $sp)
	{
		##Reverse the order of the UUID Prefix
		$ucsuuidpart1 = ($splist.uuid).substring(0,2)		
		$ucsuuidpart2 = ($splist.Uuid).Substring(2,2)		
		$ucsuuidpart3 = ($splist.Uuid).Substring(4,2)		
		$ucsuuidpart4 = ($splist.Uuid).Substring(6,2)
		$ucsuuidpart5 = ($splist.Uuid).Substring(9,2)
		$ucsuuidpart6 = ($splist.Uuid).Substring(11,2)
		$ucsuuidpart7 = ($splist.Uuid).Substring(14,2)
		$ucsuuidpart8 = ($splist.Uuid).Substring(16,2)
		$uuidsuffix = $splist.UuidSuffix
		$ucsuuidfixed = "$ucsuuidpart4$ucsuuidpart3$ucsuuidpart2$ucsuuidpart1-$ucsuuidpart6$ucsuuidpart5-$ucsuuidpart8$ucsuuidpart7-$uuidsuffix"
		if ($ucsuuidfixed -ieq $uuid)
			{
				$serviceprofile = $splist.Name
				Write-Output "	Service Profile: $serviceprofile"
				$spfound = 1
			}
	}
if ($spfound -ne 1)
	{
		Write-Output "	Service Profile NOT found"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Begin processing the server
Write-Output ""
Write-Output "Processing server $hostname...Please Wait"
Write-Output ""
$spWorking = Get-UcsServiceProfile -Type instance | where {$_.name -eq $serviceprofile}
$winNicAdapters = $null
if (!($winNicAdapters = Get-WmiObject -ComputerName $hostname -Credential $PCCred -Class Win32_NetworkAdapter -ErrorAction SilentlyContinue))
	{
		Write-Output ""
		Write-Output "	Connection to Server $($hostname) failed:"
		Write-Output "    	- Check that the IP address provided matches the Service Profile name"
		Write-Output "    	- You must have administrative priviledges to the target server"
		Write-Output "		- The Windows Firewall must be disabled or WMI rule enabled on the remote server"
		Write-Output ""
		Disconnect-Ucs
		Write-Output "Exiting..."
		exit
	}
else
	{
		Write-Output "OK...Server:$($hostname): Configuring Physical Nic Names...(Please Wait)"
		#Get server NIC based on MAC address in the service profile...rename.
		$ucsNics = Get-UcsVnic -ServiceProfile $spWorking
		foreach ($ucsNicIn in $ucsNics)
			{	
				$winNicIn = $winNicAdapters | where {($_.MACAddress -eq $ucsNicIn.Addr) -and ($_.ServiceName -ne "VMSMP")}
				$newName = [string]::Format('{0}',$ucsNicIn.Name)
				if ($winNicIn)
					{
						if ($newName -ne $winNicIn.NetConnectionID) 
							{	
								$winNicIn.NetConnectionID = $NewName
								$winNicIn.Put() | Out-Null
							}
					}
				else
					{
						Write-Output ""
						Write-Output "		- Could not find a match for $($ucsNicIn.Name)."
					}
			}
				Write-Output "	Complete"
	}
#Rename Hyper-V vSwitch Names
Write-Output ""
Write-Output "OK...Server:$($hostname): Configuring Hyper-V vSwitch Names...(Please Wait)"
#Get UCS NICS
$ucsNics = Get-UcsVnic -ServiceProfile $spWorking
#Get Hyper-V NICs
$HyperVNics = GET-WMIOBJECT -ComputerName $hostname -Credential $PCCred -class Win32_NetworkAdapter -ErrorAction SilentlyContinue | where {$_.PhysicalAdapter -ieq "TRUE" -and $_.MacAddress -ne "" -and $_.ProductName -eq "Hyper-V Virtual Ethernet Adapter"}
foreach ($ucsNicIn in $ucsNics)
	{	
		$HyperVNicIn = $HyperVNics | where {($_.MACAddress -eq $ucsNicIn.Addr) -and ($_.PhysicalAdapter -ieq "TRUE") -and ($_.ProductName -eq "Hyper-V Virtual Ethernet Adapter")}
		$newName = [string]::Format('{0}','Hyper-V vSwitch('+$ucsNicIn.Name+')')
		if ($HypervNics)
			{
				if ($newName -ne $HyperVNicIn.NetConnectionID) 
					{	
						$HyperVNicIn.NetConnectionID = $NewName
						$HyperVNicIn.Put() | Out-Null
					}
			}
	}
Write-Output "	Complete"

#Rename the switch names in Hyper-V Manager
#######################    FUTURE    #######################

#Wrap up the script
Disconnect-Ucs
Write-Output ""
Write-Output "Processing is Complete"
Write-Output ""
Write-Output "Script Complete..."
Write-Output "		***Make sure to re-enable the firewall the on server: $hostname***"
Write-Output "Exiting"
exit