##################################################
# Joe Martin
# Cisco Systems, Inc.
# Custom UCS Settings for UCS Base Configuration Builder
# System: UCS-Laptop
# 5/1/14
# Code provided as-is.  No warranty implied or included.
# This code is for example use only and not for production
#
# This script will add custom features
#
##################################################

#Connect to UCS
Write-Host -ForegroundColor White -BackgroundColor DarkBlue "Reconnecting to the UCSM"
$FIReboot = 1
if ($IsRedundantFI -ne $null)
	{
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
										Write-Host -ForegroundColor Red "The Fabric Interconnects are not accessible"
										Write-Host -ForegroundColor Red "			Script did not complete all configuration items:"
										Write-Host -ForegroundColor Red "				Customization Script"
										Write-Host -ForegroundColor Red "					Exiting..."
										Disconnect-Ucs
										exit
									}
							}
						else
							{
								Write-Host -ForegroundColor DarkBlue "  Waiting for Fabric Interconnect to be fully up"
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
										Write-Host -ForegroundColor Red "				Customization Script"
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
	}
else
	{
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
										Write-Host -ForegroundColor Red "The Fabric Interconnects are not accessible"
										Write-Host -ForegroundColor Red "			Script did not complete all configuration items:"
										Write-Host -ForegroundColor Red "				Customization Script"
										Write-Host -ForegroundColor Red "					Exiting..."
										Disconnect-Ucs
										exit
									}
							}
						else
							{
								Write-Host -ForegroundColor DarkBlue "  Waiting for Fabric Interconnect to be fully up"
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
										Write-Host -ForegroundColor Red "				Customization Script"
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
	}
Write-Host ""
Write-Host -ForegroundColor White -BackgroundColor DarkBlue "This script will set the custom UCS settings"

####################################################CUSTOM UCS SETTINGS####################################################
#
#Build iSCSI Authentication Profiles
Get-UcsOrg -Level root  | Add-UcsManagedObject -ClassId IscsiAuthProfile -PropertyMap @{Descr=""; UserId="admin1"; PolicyOwner="local"; Name="admin1"; Password="letmein12345"; }
Get-UcsOrg -Level root  | Add-UcsManagedObject -ClassId IscsiAuthProfile -PropertyMap @{Descr=""; UserId="admin2"; PolicyOwner="local"; Name="admin2"; Password="letmein12345"; }

#Assign Authentication Policies to Boot Policies
Start-UcsTransaction
$mo = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "Boot_from_iSCSI" -LimitScope | Get-UcsVnicIScsiBootParams | Set-UcsManagedObject -PropertyMap @{Descr=""; PolicyOwner="local"; } -Force
$mo_1 = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "Boot_from_iSCSI" -LimitScope | Get-UcsVnicIScsiBootParams | Get-UcsVnicIScsiBootVnic -Name "iSCSI_A" | Set-UcsVnicIScsiBootVnic -AuthProfileName "admin1" -Force
$mo_1_1 = $mo_1 | Add-UcsVnicIScsiStaticTargetIf -ModifyPresent -AuthProfileName "admin2" -IpAddress "2.2.2.254" -Name "CHANGEME" -Port 3260 -Priority 1
Complete-UcsTransaction
Start-UcsTransaction
$mo = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "Boot_from_iSCSI" -LimitScope | Get-UcsVnicIScsiBootParams | Set-UcsManagedObject -PropertyMap @{Descr=""; PolicyOwner="local"; } -Force
$mo_1 = Get-UcsOrg -Level root | Get-UcsServiceProfile -Name "Boot_from_iSCSI" -LimitScope | Get-UcsVnicIScsiBootParams | Get-UcsVnicIScsiBootVnic -Name "iSCSI_B" | Set-UcsVnicIScsiBootVnic -AuthProfileName "admin1" -Force
$mo_1_1 = $mo_1 | Add-UcsVnicIScsiStaticTargetIf -ModifyPresent -AuthProfileName "admin2"-IpAddress "3.3.3.254" -Name "CHANGEME" -Port 3260 -Priority 1
Complete-UcsTransaction

#Disable LAN Port Channels
Get-UcsFiLanCloud -Id "A" | Get-UcsUplinkPortChannel -PortId 1 | Set-UcsUplinkPortChannel -AdminState "disabled" -Force
Get-UcsFiLanCloud -Id "B" | Get-UcsUplinkPortChannel -PortId 2 | Set-UcsUplinkPortChannel -AdminState "disabled" -Force

#Disable SAN Port Channels
Get-UcsFiSanCloud -Id "A" | Get-UcsFcUplinkPortChannel -PortId 100 | Set-UcsFcUplinkPortChannel -AdminState "disabled" -Force
Get-UcsFiSanCloud -Id "B" | Get-UcsFcUplinkPortChannel -PortId 101 | Set-UcsFcUplinkPortChannel -AdminState "disabled" -Force
#
####################################################END CUSTOM SETTINGS####################################################

#Disconnect from UCS
Disconnect-Ucs
