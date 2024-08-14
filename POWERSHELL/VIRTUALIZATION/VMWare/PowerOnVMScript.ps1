<#
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.152
	 Created on:   	15/06/2018 9:32 AM
	 Created by:   	wbuntin
	 Organization: 	Heat and Control Pty Ltd
	 Filename:     	PowerOnVMScript.ps1
	===========================================================================
	.DESCRIPTION

	This script will check the hosts and CVM's are alive before then powering on
	servers in order.
#>

# Import the VMware module so we can use the commands
Import-Module vmware.vimautomation.core

# Specify the hosts we will connect to
$vHostNames = @("BRISBANE-NU01", "BRISBANE-NU02", "BRISBANE-NU03", "BRISBANE-NU04")
$cvmIPAddress = @("192.168.7.150","192.168.7.151","192.168.7.152","192.168.7.153")

# Specify the username and password to connect to the hosts
$userName = "****"
$password = "****"

# Ping hosts until response
foreach ($vHost in $vHostNames)
{
	do
	{
		Start-Sleep -Seconds 1
	}
	while (!(Test-Connection -ComputerName $vHost -Count 1 -Quiet))
}

# Ping CVM Appliances until response before moving on
foreach ($vCVM in $cvmIPAddress)
{
	do
	{
		Start-Sleep -Seconds 1
	}
	while (!(Test-Connection -ComputerName $vCVM -Count 1 -Quiet))
}

# Connect to th CVM and check if the cluster has started.
start-process $ENV:SystemRoot\system32\cmd.exe -argumentlist ('/c "c:\Program Files\Putty\plink.cmd"') -windowstyle minimized

# Wait for the command to finish and log file to ne written
Start-Sleep -s 150

# Check the log file for the word DOWN, if it is still there wait 5 mins and re-run the putty command
$file = "c:\Apps\nutx.log"

do{
	Start-Sleep -s 300
	start-process $ENV:SystemRoot\system32\cmd.exe -argumentlist ('/c "c:\Program Files\Putty\plink.cmd"') -windowstyle minimized
} while ((Get-Content $file | %{ $_ -cmatch "DOWN" }) -Contains $true)

# Connect to the host servers
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore
Connect-VIServer -Server $vHostNames -Username $userName -Password $password

# Power On critical infrastructure servers
Start-VM -VM ROOT-INT02 -RunAsync

# Wait 5 mins
Start-Sleep -s 300

# Start NA domain DC
Start-VM -VM BRISBANE-NA02 -RunAsync

# Wait 5 mins
Start-Sleep -s 300

# Power on Brisbane DC's'
Start-VM -VM BRISBANE-DC01A -RunAsync
Start-VM -VM BRISBANE-DC02A -RunAsync
Start-VM -VM BRISBANE-DC03A -RunAsync

# Wait 5 mins
Start-Sleep -s 300

# Power on all other servers
Get-VM | Where {$_.PowerState -eq 'PoweredOff'} | Start-VM