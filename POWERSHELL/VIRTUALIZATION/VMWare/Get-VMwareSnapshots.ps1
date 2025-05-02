<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.150
	 Created on:   	24/04/2018 12:13 PM
	 Created by:   	wbuntin
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

#Import the VMware module so we can use the commands
Import-Module vmware.vimautomation.core

#Specify the vCenter servers we will connect to
$vCenterNames = @("AUCKLAND-VC01", "BRISBANE-VC01", "CHENNAI-VC02", "CAPETOWN-VC02", "LIVINGSTON-VC02", "NANJING-VC01")
Connect-VIServer -Server $vCenterNames

#Get a list of the snapshots on the vCenter server
get-vm | get-snapshot | select vm, name, description, created, sizegb | Format-List


