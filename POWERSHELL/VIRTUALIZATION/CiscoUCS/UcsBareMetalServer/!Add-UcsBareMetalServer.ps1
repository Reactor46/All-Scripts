<#
-serviceprofile SERVERNAME
#>
param (
$SERVICEPROFILE
)

#Note the time when the script started
$starttime = Get-Date

#Clear the screen
clear

#Change directory to the folder containing this script to find the other files used below
cd $PSScriptRoot

#Script Parameters
if (!$SERVICEPROFILE)
	{
		$DemoSP = "Test1"						#Name of server to build
	}
else
	{
		$DemoSP = $SERVICEPROFILE
	}

$SpList = @("HyperV-Host-1", "HyperV-Host-2!", "HyperV-Host-3", "HyperV-Host-SA1", "Infra-1!", "Infra-2", "Infra-3", "Infra-SA1", "StorageSpaces", "SQLTest1", "GOLD_LUN")
if ($SpList -icontains $SERVICEPROFILE)
	{
		Write-Output "That is a reserved SP Name"
		Disconnect-Ucs
		exit
	}
		

$starttime = Get-Date
$DemoUCS = "10.0.1.6" 							#IP Address of UCSM
$DemoNetApp = "10.0.0.13"						#IP Address of NetApp
$DemoSPT = "B22-M3"								#UCSM Service Profile Template
#$DemoSPT = "B420-M3"							#UCSM Service Profile Template
#$DemoSPT = "C240-M3"							#UCSM Service Profile Template
$DemoCredentials = "myucscred.csv"				#Encrypted credentials file for solution
$DemoVserver = "Joe"							#NetApp C-Mode vServer to use
$DemoVolume = "/vol/Boot_LUNs"					#NetApp Volume to put LUN into
$DemoLUN = "Gold_LUN"							#Gold_LUN with Server 2012R2 sysprep to clone
$DemoManufacture = "Cisco"						#Manufacture of SAN fabric
$DemoWWPNfile = "WWPN Targets - Joes Lab.csv"	#File with SAN zoneset name, controller names and target WWPNs
$DemoFabricA = "10.0.1.2"						#IP Address of SAN fabric A
$DemoFabricB = "10.0.1.3"						#IP Address of SAN fabric B

<#
Creating a bare metal server with boot from SAN in less than 60 seconds

Creating a bare metal server in UCS from a workload template
	Validate Firmware: BMC, CIMC, BIOS, Adapters
	Setting BIOS
	Setting Boot Options
	Creating 10 NICs, Connecting to LAN with associated VLANs. Setting MTU and QoS policies
	Creating 2 HBAs
	Creating UUID
	Local disk RAID settings
	Creating Operational Policies: IPMI, KVM, SoL, Maintenance, Power, Stats, Scrub
	Server is booted
#>
.\New-UcsSpFromSpt.ps1 -ucs $DemoUCS -saved $DemoCredentials -spt $DemoSPT -sp $DemoSP -skiperror

<#
Creating Boot LUN in NetApp Array
	Cloning Gold LUN with Sysprep'd Windows Server 2012R2
	Creating Initiator Group
		Adding WWPNs of Server with information gathered from UCSM
	Associating Initiator Group to LUN
#>
.\New-UcsNetAppLunClone.ps1 -ucs $DemoUCS -netapp $DemoNetApp -serviceprofile $DemoSP -vserver $DemoVserver -volume $DemoVolume -goldlun $DemoLUN -usaved $DemoCredentials -nsaved $DemoCredentials -skip

<#
Creating SAN fabric configurations in Nexus switches
	Creating device aliases in Fabric A using informataion from UCSM
	Creating zones in Fabric A
	Adding zones to zoneset in Fabric A
	Activating zoneset in Fabric A
	Saving configuration in Fabric A
	Creating device aliases in Fabric B using informataion from UCSM
	Creating zones in Fabric B
	Adding zones to zoneset in Fabric B
	Activating zoneset in Fabric B
	Saving configuration in Fabric B
#>
.\New-UcsFcZoning.ps1 -req y -ucs $DemoUCS -usaved $DemoCredentials -serviceprofile $DemoSP -manufacture $DemoManufacture -wwpn $DemoWWPNfile -output Equipment -fabrica $DemoFabricA -asaved $DemoCredentials -fabricb $DemoFabricB -bsaved $DemoCredentials -skiperrors

#Note the time when the script finished
$endtime = Get-Date

<#
Doing the same zoning as above but saving to a text file to show what commands are being sent to the fabrics
#>
.\New-UcsFcZoning.ps1 -req y -ucs $DemoUCS -usaved $DemoCredentials -serviceprofile $DemoSP -manufacture $DemoManufacture -wwpn $DemoWWPNfile -output file -skiperrors

<#
Launching UCSM KVM for server
#>
.\Invoke-UcsKvm.ps1 -ucs $DemoUCS -saved $DemoCredentials -service $DemoSP -skip

<#
Powering on Server
#>
clear
Write-Output "Powering on Server..."
$CredFile = import-csv $DemoCredentials
$Username = $CredFile.UserName
$Password = $CredFile.EncryptedPassword
$cred = New-Object System.Management.Automation.PsCredential $Username,(ConvertTo-SecureString $Password)
$MyCon = Connect-Ucs -Name $DemoUCS -Credential $cred
$Power = Set-UcsServerPower -ServiceProfile $DemoSP -state "admin-up" -Force
Disconnect-Ucs

#Show the user the start and end time
sleep -Seconds 5
clear
Write-Output "Script Start Time: $starttime"
Write-Output "Script End Time  : $endtime"
$runtime = $endtime - $starttime
$seconds = $runtime.TotalSeconds
#Show the user the total time to run the script
Write-Output "Time to build: $Seconds seconds"
Write-Output ""
#Wrap things up!
Write-Output "Server $DemoSP built and booting"