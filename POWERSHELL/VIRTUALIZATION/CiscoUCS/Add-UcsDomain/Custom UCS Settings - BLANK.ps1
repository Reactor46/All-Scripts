##################################################
# Joe Martin
# Cisco Systems, Inc.
# Custom UCS Settings for UCS Base Configuration Builder
# System: <Provide system description info here>
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

#
####################################################END CUSTOM SETTINGS####################################################

#Disconnect from UCS
Disconnect-Ucs
