<#

.SYNOPSIS
	This script creates a NetApp Initiator Group and Clones a LUN based on UCS Service Profile information

.DESCRIPTION
	This script creates a NetApp Initiator Group and Clones a LUN based on UCS Service Profile information
	This script only functions on Clustered OnTAP

.EXAMPLE
	New-UCSNetAppLunClone.ps1
	This script can be run without any command line parameters.  User will be prompted for all parameters and options required

.EXAMPLE
	New-UCSNetAppLunClone.ps1 -ucs "1.2.3.4" -ucred -netapp "2.3.4.5" -ncred -serviceprofile "Test1" -vserver "Production" -volume "/vol/Boot_LUNs" -goldlun "Gold_LUN"
	-ucs -- UCS Manager IP address or Host Name -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local"
	-ucred -- UCS Manager Credential Switch -- Adding this switch will immediately prompt you for your UCSM username and password
	-netapp - NetApp IP address or Host Name -- Example: "1.2.3.4" or "mynetap" or "mynetapp.domain.local"
	-ncred -- NetApp Credential Switch -- Adding this switch will immediately prompt you for your NetApp username and password
	-serviceprofile -- Service Profile name to create LUN for
	-vserver -- NetApp vServer to add LUN and Initiator Group to
	-volume -- Name of the Volume or QTree that contains the Gold LUN and will contain the new LUN
	-goldlun -- Name of the Gold LUN to use as the source
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords

.EXAMPLE
	New-UcsNetAppLunClone.ps1 -ucs "1.2.3.4" -usaved "myucscred.csv" -netapp "2.3.4.5" -usaved "mynetappcred.csv" -serviceprofile "Test1" -vserver "Production" -volume "/vol/Boot_LUNs" -goldlun "Gold_LUN" -skiperrors
	-ucs -- UCS Manager IP address or Host Name -- Example: "1.2.3.4" or "myucs" or "myucs.domain.local"
	-usavedcred -- UCSM credentials file -- Example: -savedcred "myucscred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\myucscred.csv
		Make sure the password file is located in the same folder as the script
	-netapp - NetApp IP address or Host Name -- Example: "1.2.3.4" or "mynetap" or "mynetapp.domain.local"
	-nsavedcred -- NetApp credentials file -- Example: -savedcred "mynetappcred.csv"
		To create a credentials file: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} |Export-CSV -NoTypeInformation .\mynetappcred.csv
		Make sure the password file is located in the same folder as the script
	-serviceprofile -- Service Profile name to create LUN for
	-vserver -- NetApp vServer to add LUN and Initiator Group to
	-volume -- Name of the Volume or QTree that contains the Gold LUN and will contain the new LUN
	-goldlun -- Name of the Gold LUN to use as the source
	-skiperrors -- Tells the script to skip any prompts for errors and continues with 'y'
	All parameters are optional and any skipped will be prompted for during execution
	The only prompts that will always be presented to the user will be for User Names and Passwords


.NOTES
	Author: Joe Martin
	Email: joemar@cisco.com
	Company: Cisco Systems, Inc.
	Version: v0.1.00
	Date: 7/16/2014
	Disclaimer: Code provided as-is.  No warranty implied or included.  This code is for example use only and not for production

.INPUTS
	UCSM IP Address or Hostname
	UCSM Username and Password
	UCSM Credentials Filename
	NetApp IP Address or Hostname
	NetApp Username and Password
	NetApp Credentials Filename
	Service Profile
	Volume
	vServer
	Gold LUN

.OUTPUTS
	None
	
.LINK
	http://communities.cisco.com/people/joemar/content

#>

#Command Line Parameters
param(
	[string]$UCSM,				# IP Address or Hostname.
	[string]$NETAPP,			# IP Address or Hostname.
	[switch]$NCREDENTIALS,		# NetApp Credentials (Username and Password).
	[switch]$UCREDENTIALS,		# UCSM Credentials (Username and Password).
	[string]$SERVICEPROFILE,	# Service Profile Name
	[string]$VSERVER,			# vServer
	[string]$VOLUME,			# Volume
	[string]$GOLDLUN,			# Gold LUN
	[string]$USAVEDCRED,		# Saved UCSM Credentials.  To create do: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} | Export-CSV -NoTypeInformation .\myucscred.csv
	[string]$NSAVEDCRED,		# Saved NetApp Credentials.  To create do: $credential = Get-Credential ; $credential | select username,@{Name="EncryptedPassword";Expression={ConvertFrom-SecureString $_.password}} | Export-CSV -NoTypeInformation .\mynetappcred.csv
	[switch]$SKIPERROR			# Skip any prompts for errors and continues with 'y'
)

#Clear the screen
clear-host

#Script kicking off
Write-Output "Script Running..."
Write-Output ""

#Tell the user what the script does
Write-Output "This script creates a NetApp Initiator Group and Clones a LUN based on UCS Service Profile information"
Write-Output ""

if ($UCREDENTIALS)
	{
		Write-Output "Enter UCSM Credentials"
		Write-Output ""
		$cred = Get-Credential -Message "Enter UCSM Credentials"
	}

if ($NCREDENTIALS)
	{
		Write-Output "Enter NetApp Credentials"
		Write-Output ""
		$ncred = Get-Credential -Message "Enter NetApp Credentials"
	}

#Change directory to the script root
cd $PSScriptRoot

#Check to see if UCS credential files exists
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

#Check to see if NetApp credential files exists
if ($NSAVEDCRED)
	{
		if ((Test-Path $NSAVEDCRED) -eq $false)
			{
				Write-Output ""
				Write-Output "Your credentials file $NSAVEDCRED does not exist in the script directory"
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
				Write-Output "	Cisco UCS PowerTool Module did not load.  Please correct this issue and try again"
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

#Load the NetApp PowerShell Module
Write-Output ""
Write-Output "Checking NetApp Module"
$PowerToolLoaded = $null
$Modules = Get-Module
$PowerToolLoaded = $modules.name
if ( -not ($Modules -like "DataONTAP"))
	{
		Write-Output "	Loading Module: NetApp DataOnTAP"
		Import-Module DataONTAP
		$Modules = Get-Module
		if ( -not ($Modules -like "DataONTAP"))
			{
				Write-Output ""
				Write-Output "	NetApp DataOnTAP PowerShell Module did not load.  Please correct this issue and try again"
				Write-Output "		Exiting..."
				exit
			}
		else
			{
				Write-Output "		NetApp DataOnTAP Module now loaded"
			}
	}
else
	{
		Write-Output "	NetApp DataOnTAP Module is already Loaded"
	}

#Select UCS Domain(s) for login
if ($UCSM -ne "")
	{
		$myucs = $UCSM
	}
else
	{
		$myucs = Read-Host "Enter UCS system IP or Hostname"
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

#Select NetApp for login
if ($NETAPP -ne "")
	{
		$mynetapp = $NETAPP
	}
else
	{
		$mynetapp = Read-Host "Enter NetApp system IP or Hostname"
	}
[array]$mynetapp = ($mynetapp.split(",")).trim()
if ($mynetapp.count -eq 0)
	{
		Write-Output ""
		Write-Output "You didn't enter anything"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Make sure we are disconnected from UCS
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

#Test that NetApp is IP Reachable via Ping
Write-Output ""
Write-Output "Testing PING access to NetApp"
foreach ($netapplist in $mynetapp)
	{
		$ping = new-object system.net.networkinformation.ping
		$results = $ping.send($netapplist)
		if ($results.Status -ne "Success")
			{
				Write-Output "	Can not access NetApp $netapplist by Ping"
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
				Write-Output "	Successful access to $netapplist by Ping"
			}
	}
	
#Log into the UCS
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
foreach ($myucslist in $myucs)
	{
		Write-Output "		Logging into: $myucslist"
		$myCon = $null
		$myCon = Connect-Ucs $myucslist -Credential $credu
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

#Log into the NetApp
Write-Output ""
Write-Output "Logging into NetApp"
#Verify PowerShell Version to pick prompt type
$PSVersion = $psversiontable.psversion
$PSMinimum = $PSVersion.Major
if (!$NCREDENTIALS)
	{
		if (!$NSAVEDCRED)
			{
				if ($PSMinimum -ge "3")
					{
						Write-Output "	Enter your NetApp credentials"
						$credn = Get-Credential -Message "NetApp Login Credentials" -UserName "admin"
					}
				else
					{
						Write-Output "	Enter your NetApp credentials"
						$credn = Get-Credential
					}
			}
		else
			{
				$CredFile = import-csv $NSAVEDCRED
				$Username = $credfile.UserName
				$Password = $credfile.EncryptedPassword
				$credn = New-Object System.Management.Automation.PsCredential $Username,(ConvertTo-SecureString $Password)			
			}
	}
foreach ($mynetapplist in $mynetapp)
	{
		Write-Output "		Logging into: $mynetapplist"
		$myCon = $null
		$myConN = Connect-NcController -Name $mynetapplist -Credential $credn -Vserver $VSERVER
		if (($MyConN).Name -ne ($mynetapplist)) 
			{
				#Exit Script
				Write-Output "			Error Logging into this NetApp"
				if ($mynetapp.count -le 1)
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
								$continue = Read-Host "Continue without this NetApp (Y/N)"
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
$myConN = (Get-NcClusterNode | measure).Count
if ($myCon -eq 0)
	{
		Write-Output ""
		Write-Output "You are not logged into any NetApp systems"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
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

if (!$SERVICEPROFILE)
	{
		#Select Service Profile to add LUN to
		Write-Output ""
		Write-Output "Select the Service Profile to add a LUN to"
		Write-Output "	Press CANCEL or hit Esc to exit the script"
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
		$objLabel.Text = "Select from below."
		$objForm.Controls.Add($objLabel) 
		
		$objListbox = New-Object System.Windows.Forms.Listbox 
		$objListbox.Location = New-Object System.Drawing.Size(10,40) 
		$objListbox.Size = New-Object System.Drawing.Size(260,20)
		$objListBox.Sorted = $True

		#$objListbox.SelectionMode = "MultiExtended"

		foreach ($SP in $vHBAInfo.ServiceProfile)
			{
				$ServiceProfileFull = $SP -match "/ls-(?<content>.*)/fc-"
				$SERVICEPROFILE = $matches['content']
				[void] $objListbox.Items.Add($SERVICEPROFILE)
			}

		$objListbox.Height = 450
		$objForm.Controls.Add($objListbox) 
		$objForm.Topmost = $True
	
		$objForm.Add_Shown({$objForm.Activate()})
		[void] $objForm.ShowDialog()
	}
else
	{
		[array]$SPArray = ($SERVICEPROFILE.split(",")).trim()
		$Script:SelectedObjects = $SPArray
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

#Put selection into variable
[string]$SP = $Script:SelectedObjects

#What vServer will you add the LUN and IG to
if (!$VSERVER)
	{
		$VSERVER = Read-Host "Enter the vServer to use"
	}
$CheckVserver = Get-NcVserver -Name $VSERVER
if (!$CheckVserver)
	{
		Write-Output ""
		Write-Output "$VSERVER is not a valid vServer"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Where is the Gold LUN Located
if (!$VOLUME)
	{
		$VOLUME = Read-Host "Enter the Volume where the Gold LUN is located"
	}
$CheckVolume = Get-NcVol -Name $VOLUME
if (!$CheckVolume)
	{
		Write-Output ""
		Write-Output "$VOLUME is not a valid Volume"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#What LUN will you clone
if (!$GOLDLUN)
	{
		$GOLDLUN = Read-Host "Enter the Gold LUN to Clone"
	}
$Path = $VOLUME+"/"+$GOLDLUN
$CheckGoldLUN = Get-NcLun -Path $Path
if (!$CheckGoldLUN)
	{
		Write-Output ""
		Write-Output "$Path is not a valid LUN"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}

#Create new Initiator Group
$ExistingIgroup = Get-NcIgroup -Name $SP
if (!$ExistingIgroup)
	{
		$NewIgroup = New-NcIgroup -name $SP -protocol fcp -vserver $VSERVER
		if ($NewIgroup)
			{
				Write-Output ""
				Write-Output "Initiator Group: $SP successfully created"
			}
		else
			{
				Write-Output "Error creating Initiator Group: $SP"
				Disconnect-Ucs
				Write-Output "	Exiting..."
				exit
			}
	}
else
	{
		Write-Output "Initiator Group: $SP already exists"
		Write-Output "	Exiting..."
		Disconnect-Ucs
		exit
	}
	
#Add WWPNs to Initiator Group
$SPInfo = Get-UcsServiceProfile -Name $SP
$vHBAs = $SPInfo | Get-UcsVhba
if (!$vHBAs)
	{
		Write-Output "This service profile does not contain any vHBAs"
		$DontShow = Remove-NcIgroup -Name $SP -VserverContext $VSERVER -Force
		Disconnect-Ucs
		Write-Output "	Exiting..."
		exit
	}
if (!$ExistingIgroup)
	{
		foreach ($vHBAlist in $vHBAs)
			{
				$AddIgroupInitiator = Add-NcIgroupInitiator -Initiator $vHBAlist.Addr -Name $SP -vserver $VSERVER
				if (!$AddIgroupInitiator)
					{
						$Output = $SPInfo.Name
						Write-Output "Failed to add Initiator: $vHBAlist.Addr to Group: $Output"
						$DontShow = Remove-NcIgroup -Name $SP -VserverContext $VSERVER -Force
						Write-Output ""
						Write-Output "Removing Initiator Group: $SP"
						Write-Output "	Exiting..."
						Disconnect-Ucs
						exit
					}
				else
					{
						$Output = $vHBAlist.Addr
						Write-Output "	$Output added to Initiator Group: $SP"
					}
			}
	}

#Clone the gold LUN for the new Service Profile
$DestinationPath = $SP+"_C"
$ExistingLUN = Get-NcLun -Path $DestinationPath
if ($ExistingLUN)
	{
		Write-Output "LUN: $DestinationPath already exists"
		Write-Output "	Removing Initiator Group: $SP"
		$DontShow = Remove-NcIgroup -Name $SP -VserverContext $VSERVER -Force
		Write-Output "		Exiting..."
		Disconnect-Ucs
		exit
	}
else
	{
		$DontShow = New-NcClone -volume $VOLUME -SourcePath $GOLDLUN -DestinationPath $DestinationPath -VserverContext $VSERVER
		$Path = $VOLUME+"/"+$DestinationPath
		$NewClone = Get-NcLUN -path $Path
		if (!$NewClone)
			{
				Write-Output ""
				Write-Output "Cloning of LUN: $GOLDLUN to LUN: $DestinationPath failed"
				Write-Output "	Removing Initiator Group: $SP"
				$DontShow = Remove-NcIgroup -Name $SP -VserverContext $VSERVER -Force
				Write-Output "	Exiting..."
				Disconnect-Ucs
				exit
			}
		else
			{
				Write-Output ""
				Write-Output "Cloning of LUN: $GOLDLUN to LUN: $DestinationPath completed"
			}
	}

#Associate Initiator Group to new LUN
$NewLUN = $VOLUME+"/"+$DestinationPath
$AddIG = Add-NcLunMap -Path $NewLUN -InitiatorGroup $SP -vserver $VSERVER
if (!$AddIG)
	{
		Write-Output ""
		Write-Output "Initiator Group failed to associate to new LUN"
		Write-Output "	Removing LUN"
		Remove-NcLun -Path $NewLUN -VserverContext $VSERVER -Confirm:$false
		Write-Output "	Removing Initiator Group"
		$DontShow = Remove-NcIgroup -Name $SP -VserverContext $VSERVER -Force
		Write-Output "		Exiting..."
		Disconnect-Ucs
		exit
	}
else
	{
		Write-Output ""
		Write-Output "Initiator Group: $SP successfully associated to LUN: $NewLUN"
	}

#Disconnect from UCSM
Disconnect-Ucs

#Exit the Script
Write-Output ""
Write-Output "Script Complete"
exit