# Copyright (c) Microsoft Corporation. All rights reserved.
#
# StartDagServerMaintenenance.ps1

# .SYNOPSIS
# Calls Suspend-MailboxDatabaseCopy on the database copies.
# Pauses the node in Failover Clustering so that it can not become the Primary Active Manager.
# Suspends database activation on each mailbox database.
# Sets the DatabaseCopyAutoActivationPolicy to Blocked on the server.
# Moves databases and cluster group off of the designated server.
#
# If there's a failure in any of the above, the operations are undone, with
# the exception of successful database moves.
#
# Can be run remotely, but it requires the cluster administrative tools to
# be installed (RSAT-Clustering).

# .PARAMETER serverName
# The name of the server on which to start maintenance. FQDNs are valid

# .PARAMETER whatif
# Does not actually perform any operations, but logs what would be executed
# to the verbose stream.

# .PARAMETER overrideMinimumTwoCopies
# Allows users to override the default minimum number of database copies to require
# to be up after shutdown has completed.  This is meant to allow upgrades
# in situations where users only have 2 copies of a database in their dag.

Param(
	[Parameter(Mandatory=$true)]
	[System.Management.Automation.ValidateNotNullOrEmptyAttribute()]
	[string] $serverName,

	[Parameter(Mandatory=$false)] [switch] $whatif = $false,
	[Parameter(Mandatory=$false)] [switch] $overrideMinimumTwoCopies = $false
)

# Global Values
$ServerCountinTwoServerDAG = 2

Import-LocalizedData -BindingVariable StartDagServerMaintenance_LocalizedStrings -FileName StartDagServerMaintenance.strings.psd1

# Define some useful functions.

# Load the Exchange snapin if it's no already present.
function LoadExchangeSnapin
{
    if (! (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:SilentlyContinue) )
    {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:Stop
    }
}


# The meat of the script!
&{
	# Get the current script name. The method is different if the script is
	# executed or if it is dot-sourced, so do both.
	$thisScriptName = $myinvocation.scriptname
	if ( ! $thisScriptName )
	{
		$thisScriptName = $myinvocation.MyCommand.Path
	}

	# Many of the script libraries already use $DagScriptTesting
	if ( $whatif )
	{
		$DagScriptTesting = $true;
	}

	# Load the Exchange cmdlets.
	& LoadExchangeSnapin

	# Load some of the common functions.
	. "$(split-path $thisScriptName)\DagCommonLibrary.ps1";
	
	Test-RsatClusteringInstalled

	# Allow an FQDN to be passed in, but strip it to the short name.
	$shortServerName = $serverName;
	if ( $shortServerName.Contains( "." ) )
	{
		$shortServerName = $shortServerName -replace "\..*$"
	}
	
	# Variables to keep track of what needs to be rolled back in the event of failure.
	$pausedNode = $false;
	$activationBlockedOnServer = $false;
	$databasesSuspended = $false;
	$scriptCompletedSuccessfully = $false;

	try {
        # Stage 1 - block auto activation on the server,
        # both at the server level and the database copy level
        
		log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0008 -f $shortServerName)
		if ($DagScriptTesting)
		{
			write-host ($StartDagServerMaintenance_LocalizedStrings.res_0009 -f $shortServerName,"Set-MailboxServer","-Identity","-DatabaseCopyAutoActivationPolicy")
		}
		else
		{
			Set-MailboxServer -Identity $shortServerName -DatabaseCopyAutoActivationPolicy:Blocked
			$activationBlockedOnServer = $true;
		}
			
		# Get all databases with multiple copies.
		$databases = Get-MailboxDatabase -Server $shortServerName | where { $_.ReplicationType -eq 'Remote' }
	
		if ( $databases )
		{
			# Suspend database copy. When suspended with ActivationOnly,
			# no alerts should be raised.
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0010 -f $shortServerName,"Start-DagServerMaintenance")
	
			if ( $DagScriptTesting )
			{
				$databases | foreach { write-host ($StartDagServerMaintenance_LocalizedStrings.res_0011 -f ($_.Name),(get-date -format s),$shortServerName,$false,"Suspend-MailboxDatabaseCopy","-ActivationOnly","-Confirm","-SuspendComment") }
			}
			else
			{
				$databases | foreach { Suspend-MailboxDatabaseCopy "$($_.Name)\$shortServerName" -ActivationOnly -Confirm:$false -SuspendComment "Suspended ActivationOnly by StartDagServerMaintenance.ps1 at $(get-date -format s)" }
				$databasesSuspended = $true;
			}
		}
		else
		{
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0012 -f $shortServerName,"get-mailboxdatabase")
		}
        
        # Stage 2 - pause the node in the cluster to stop it becoming the PAM
        
		# Explicitly connect to clussvc running on serverName. This script could
		# easily be run remotely.
		log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0000 -f $shortServerName);
		if ( $DagScriptTesting )
		{
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0001 )
		}
		else
		{
			# Try to fetch $dagName if we can.
			$dagName = $null
			$mbxServer = get-mailboxserver $serverName -erroraction:silentlycontinue
			if ( $mbxServer -and $mbxServer.DatabaseAvailabilityGroup )
			{
				$dagName = $mbxServer.DatabaseAvailabilityGroup.Name;
			}

			# Start with $serverName (which may or may not be a FQDN) before
			# falling back to the (short) names of the DAG.

			$outputStruct = Call-ClusterExe -dagName $dagName -serverName $serverName -clusterCommand "node $shortServerName /pause"
			$LastExitCode = $outputStruct[ 0 ];
			$output = $outputStruct[ 1 ];

			if ( $LastExitCode -eq 1753 )
			{
				# 1753 is EPT_S_NOT_REGISTERED
				log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0002 -f $serverName,"Start-DagServerMaintenance")
			}
			elseif ( $LastExitCode -eq 1722 )
			{
				# 1722 is RPC_S_SERVER_UNAVAILABLE
				log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0003 -f $serverName,"Start-DagServerMaintenance")
			}
			elseif ($LastExitCode -ne 0)
			{
				Log-Error ($StartDagServerMaintenance_LocalizedStrings.res_0004 -f $serverName,$LastExitCode,"Start-DagServerMaintenance") -stop
			}
			$pausedNode = $true;
		}
        
        # Stage 3 - move all the resources off the server
            
		log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0005 -f $shortServerName,"Start-DagServerMaintenance")
		$criticalMailboxResources = @(GetCriticalMailboxResources $shortServerName )
		$numCriticalResources = ($criticalMailboxResources | Measure-Object).Count
		log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0006 -f $numCriticalResources,"Start-DagServerMaintenance")

		if ( $criticalMailboxResources )
		{
			write-host ($StartDagServerMaintenance_LocalizedStrings.res_0007 -f ( PrintCriticalMailboxResourcesOutput($criticalMailboxResources)),$shortServerName)
		}

		# Only bother doing the operations if there's something to do.
		if( $numCriticalResources -gt 0 )
		{
	
			# Move the critical resources off the specified server. 
			# This includes Active Databases, and the Primary Active Manager.
			# If any error occurs in this stage, script execution will halt.
			# (If we don't assign the result to a variable then the script will
			# print out 'True')
			$ignoredResult = Move-CriticalMailboxResources -Server $shortServerName
			
		}
	
		# Check again to see if the moves were successful. (Unless -whatif was
		# specified, then it's pretty likely it will fail).
		if ( !$DagScriptTesting )
		{
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0013 -f $shortServerName,"Start-DagServerMaintenance")			
			$dagObject = Get-DatabaseAvailabilityGroup $dagName			
			$dagServers = $dagObject.Servers.Count			
			$stoppedDagServers = $dagObject.StoppedMailboxServers.Count			
			if (($dagServers - $stoppedDagServers) -eq $ServerCountinTwoServerDAG -or $overrideMinimumTwoCopies)
			{				
				$criticalMailboxResources = @(GetCriticalMailboxResources $shortServerName -AtleastNCriticalCopies ($ServerCountinTwoServerDAG - 1))
			}
			else
			{			
				$criticalMailboxResources = @(GetCriticalMailboxResources $shortServerName)	
			}			
			$numCriticalResources = ($criticalMailboxResources | Measure-Object).Count
		
			if( $numCriticalResources -gt 0 )
			{
				Log-CriticalResource $criticalMailboxResources
				write-error ($StartDagServerMaintenance_LocalizedStrings.res_0014 -f ( PrintCriticalMailboxResourcesOutput($criticalMailboxResources)),$shortServerName) -erroraction:stop 
			}
		}


		$scriptCompletedSuccessfully = $true;
	}
	finally
	{
		if ( ! $scriptCompletedSuccessfully )
		{
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0015 )

			# Create a new script block so that $ErrorActionPreference only
			# affects this scope.
			&{
				# Cleanup code is run with "Continue" ErrorActionPreference
				$ErrorActionPreference = "Continue"
                
				if ( $pausedNode )
				{
					# Explicitly connect to clussvc running on serverName. This script could
					# easily be run remotely.
					log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0018 -f $serverName,$shortServerName,$serverName);
					if ( $DagScriptTesting )
					{
						write-host ($StartDagServerMaintenance_LocalizedStrings.res_0019 )
					}
					else
					{
						$outputStruct = Call-ClusterExe -dagName $dagName -serverName $serverName -clusterCommand "node $shortServerName /resume"
						$LastExitCode = $outputStruct[ 0 ];
						$output = $outputStruct[ 1 ];

						# 0 is success, 5058 is ERROR_CLUSTER_NODE_NOT_PAUSED.
						if ( $LastExitCode -eq 5058 )
						{
							log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0020 -f $serverName)
						}
						elseif ( $LastExitCode -eq 1753 )
						{
							# 1753 is EPT_S_NOT_REGISTERED
							log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0021 -f $serverName,"Start-DagServerMaintenance")
						}
						elseif ( $LastExitCode -eq 1722 )
						{
							# 1722 is RPC_S_SERVER_UNAVAILABLE
							log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0022 -f $serverName,"Start-DagServerMaintenance")
						}
						elseif ( $LastExitCode -ne 0 )
						{
							Log-Error ($StartDagServerMaintenance_LocalizedStrings.res_0023 -f $serverName,$serverName,$shortServerName,$LastExitCode,"Start-DagServerMaintenance") -stop 
						}
				
					}
				}

				if ( $databasesSuspended )
				{
					if ( $databases )
					{
						# 1. Resume database copy. This clears the ActivationOnly suspension.
						foreach ( $database in $databases )
						{
							log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0024 -f ($database.Name),$shortServerName);
							if ( $DagScriptTesting )
							{
								write-host ($StartDagServerMaintenance_LocalizedStrings.res_0017 -f "resume-mailboxdatabasecopy")
							}
							else
							{
								Resume-MailboxDatabaseCopy "$($database.Name)\$shortServerName" -Confirm:$false
							}
						}
					}
				}
				
				if ( $activationBlockedOnServer )
				{
					log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0016 -f $shortServerName)
					
					if ( $DagScriptTesting )
					{
						write-host ($StartDagServerMaintenance_LocalizedStrings.res_0017 -f "set-mailboxserver")
					}
					else
					{
						Set-MailboxServer -Identity $shortServerName -DatabaseCopyAutoActivationPolicy:Unrestricted
					}
				}
			}
		}
	}
}

# SIG # Begin signature block
# MIIa0wYJKoZIhvcNAQcCoIIaxDCCGsACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUK7wVIc7kbgH4uBWJvG2UumkP
# jhSgghWCMIIEwzCCA6ugAwIBAgITMwAAAHGzLoprgqofTgAAAAAAcTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTUwMzIwMTczMjAz
# WhcNMTYwNjIwMTczMjAzWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkI4RUMtMzBBNC03MTQ0MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA6pG9soj9FG8h
# NigDZjM6Zgj7W0ukq6AoNEpDMgjAhuXJPdUlvHs+YofWfe8PdFOj8ZFjiHR/6CTN
# A1DF8coAFnulObAGHDxEfvnrxLKBvBcjuv1lOBmFf8qgKf32OsALL2j04DROfW8X
# wG6Zqvp/YSXRJnDSdH3fYXNczlQqOVEDMwn4UK14x4kIttSFKj/X2B9R6u/8aF61
# wecHaDKNL3JR/gMxR1HF0utyB68glfjaavh3Z+RgmnBMq0XLfgiv5YHUV886zBN1
# nSbNoKJpULw6iJTfsFQ43ok5zYYypZAPfr/tzJQlpkGGYSbH3Td+XA3oF8o3f+gk
# tk60+Bsj6wIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFPj9I4cFlIBWzTOlQcJszAg2
# yLKiMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAC0EtMopC1n8Luqgr0xOaAT4ku0pwmbMa3DJh+i+h/xd9N1P
# pRpveJetawU4UUFynTnkGhvDbXH8cLbTzLaQWAQoP9Ye74OzFBgMlQv3pRETmMaF
# Vl7uM7QMN7WA6vUSaNkue4YIcjsUe9TZ0BZPwC8LHy3K5RvQrumEsI8LXXO4FoFA
# I1gs6mGq/r1/041acPx5zWaWZWO1BRJ24io7K+2CrJrsJ0Gnlw4jFp9ByE5tUxFA
# BMxgmdqY7Cuul/vgffW6iwD0JRd/Ynq7UVfB8PDNnBthc62VjCt2IqircDi0ASh9
# ZkJT3p/0B3xaMA6CA1n2hIa5FSVisAvSz/HblkUwggTsMIID1KADAgECAhMzAAAA
# ymzVMhI1xOFVAAEAAADKMA0GCSqGSIb3DQEBBQUAMHkxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xIzAhBgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBMB4XDTE0MDQyMjE3MzkwMFoXDTE1MDcyMjE3MzkwMFowgYMxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIx
# HjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJZxXe0GRvqEy51bt0bHsOG0ETkDrbEVc2Cc66e2bho8
# P/9l4zTxpqUhXlaZbFjkkqEKXMLT3FIvDGWaIGFAUzGcbI8hfbr5/hNQUmCVOlu5
# WKV0YUGplOCtJk5MoZdwSSdefGfKTx5xhEa8HUu24g/FxifJB+Z6CqUXABlMcEU4
# LYG0UKrFZ9H6ebzFzKFym/QlNJj4VN8SOTgSL6RrpZp+x2LR3M/tPTT4ud81MLrs
# eTKp4amsVU1Mf0xWwxMLdvEH+cxHrPuI1VKlHij6PS3Pz4SYhnFlEc+FyQlEhuFv
# 57H8rEBEpamLIz+CSZ3VlllQE1kYc/9DDK0r1H8wQGcCAwEAAaOCAWAwggFcMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQfXuJdUI1Whr5KPM8E6KeHtcu/
# gzBRBgNVHREESjBIpEYwRDENMAsGA1UECxMETU9QUjEzMDEGA1UEBRMqMzE1OTUr
# YjQyMThmMTMtNmZjYS00OTBmLTljNDctM2ZjNTU3ZGZjNDQwMB8GA1UdIwQYMBaA
# FMsR6MrStBZYAck3LjMWFrlMmgofMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9j
# cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY0NvZFNpZ1BDQV8w
# OC0zMS0yMDEwLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljQ29kU2lnUENBXzA4LTMx
# LTIwMTAuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQB3XOvXkT3NvXuD2YWpsEOdc3wX
# yQ/tNtvHtSwbXvtUBTqDcUCBCaK3cSZe1n22bDvJql9dAxgqHSd+B+nFZR+1zw23
# VMcoOFqI53vBGbZWMrrizMuT269uD11E9dSw7xvVTsGvDu8gm/Lh/idd6MX/YfYZ
# 0igKIp3fzXCCnhhy2CPMeixD7v/qwODmHaqelzMAUm8HuNOIbN6kBjWnwlOGZRF3
# CY81WbnYhqgA/vgxfSz0jAWdwMHVd3Js6U1ZJoPxwrKIV5M1AHxQK7xZ/P4cKTiC
# 095Sl0UpGE6WW526Xxuj8SdQ6geV6G00DThX3DcoNZU6OJzU7WqFXQ4iEV57MIIF
# vDCCA6SgAwIBAgIKYTMmGgAAAAAAMTANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZIm
# iZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQD
# EyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMTAwODMx
# MjIxOTMyWhcNMjAwODMxMjIyOTMyWjB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALJyWVwZMGS/HZpgICBC
# mXZTbD4b1m/My/Hqa/6XFhDg3zp0gxq3L6Ay7P/ewkJOI9VyANs1VwqJyq4gSfTw
# aKxNS42lvXlLcZtHB9r9Jd+ddYjPqnNEf9eB2/O98jakyVxF3K+tPeAoaJcap6Vy
# c1bxF5Tk/TWUcqDWdl8ed0WDhTgW0HNbBbpnUo2lsmkv2hkL/pJ0KeJ2L1TdFDBZ
# +NKNYv3LyV9GMVC5JxPkQDDPcikQKCLHN049oDI9kM2hOAaFXE5WgigqBTK3S9dP
# Y+fSLWLxRT3nrAgA9kahntFbjCZT6HqqSvJGzzc8OJ60d1ylF56NyxGPVjzBrAlf
# A9MCAwEAAaOCAV4wggFaMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFMsR6MrS
# tBZYAck3LjMWFrlMmgofMAsGA1UdDwQEAwIBhjASBgkrBgEEAYI3FQEEBQIDAQAB
# MCMGCSsGAQQBgjcVAgQWBBT90TFO0yaKleGYYDuoMW+mPLzYLTAZBgkrBgEEAYI3
# FAIEDB4KAFMAdQBiAEMAQTAfBgNVHSMEGDAWgBQOrIJgQFYnl+UlE/wq4QpTlVnk
# pDBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtp
# L2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEE
# SDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2Nl
# cnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDANBgkqhkiG9w0BAQUFAAOCAgEAWTk+
# fyZGr+tvQLEytWrrDi9uqEn361917Uw7LddDrQv+y+ktMaMjzHxQmIAhXaw9L0y6
# oqhWnONwu7i0+Hm1SXL3PupBf8rhDBdpy6WcIC36C1DEVs0t40rSvHDnqA2iA6VW
# 4LiKS1fylUKc8fPv7uOGHzQ8uFaa8FMjhSqkghyT4pQHHfLiTviMocroE6WRTsgb
# 0o9ylSpxbZsa+BzwU9ZnzCL/XB3Nooy9J7J5Y1ZEolHN+emjWFbdmwJFRC9f9Nqu
# 1IIybvyklRPk62nnqaIsvsgrEA5ljpnb9aL6EiYJZTiU8XofSrvR4Vbo0HiWGFzJ
# NRZf3ZMdSY4tvq00RBzuEBUaAF3dNVshzpjHCe6FDoxPbQ4TTj18KUicctHzbMrB
# 7HCjV5JXfZSNoBtIA1r3z6NnCnSlNu0tLxfI5nI3EvRvsTxngvlSso0zFmUeDord
# EN5k9G/ORtTTF+l5xAS00/ss3x+KnqwK+xMnQK3k+eGpf0a7B2BHZWBATrBC7E7t
# s3Z52Ao0CW0cgDEf4g5U3eWh++VHEK1kmP9QFi58vwUheuKVQSdpw5OPlcmN2Jsh
# rg1cnPCiroZogwxqLbt2awAdlq3yFnv2FoMkuYjPaqhHMS+a3ONxPdcAfmJH0c6I
# ybgY+g5yjcGjPa8CQGr/aZuW4hCoELQ3UAjWwz0wggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBLswggS3
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAAAymzVMhI1xOFV
# AAEAAADKMAkGBSsOAwIaBQCggdQwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFDPZ
# A/fpGiRnMgE80YMOZVKGFquGMHQGCisGAQQBgjcCAQwxZjBkoDyAOgBTAHQAYQBy
# AHQARABhAGcAUwBlAHIAdgBlAHIATQBhAGkAbgB0AGUAbgBhAG4AYwBlAC4AcABz
# ADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG
# 9w0BAQEFAASCAQCQP0VTxxYdql7X53XZDqPrCtmsDKDiWH1DyVEvqE8uaOCCu03a
# G8PUj4zseB8s3emXI0E2HdwPzGB1ged5Kl3FjUOHHN8+dGS2muHeaKhTgcarDlf0
# TuOCHFfxzEjS9A3dbNS/2ubgN3IAMpVAThbeZsLqeR+heJED6ZKXaZC6A4rYiH1W
# u1geBn/57igmF8/ZtYTCYKztylLJ8mAgpuzAYfdI5pv/yVMClbi1SnxCilVTJHAw
# fP0mYUzMdBY0RrR9pmJQ9CyCL3QSQN2vfOq/SEw3+jWYNKzGhSzM6DwMXpT7xMjj
# Me54aBbobX2KYp7y5p/DgZopsNX3Ady/xuZToYICKDCCAiQGCSqGSIb3DQEJBjGC
# AhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAHGzLopr
# gqofTgAAAAAAcTAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEH
# ATAcBgkqhkiG9w0BCQUxDxcNMTUwNDEwMDI1NzQ5WjAjBgkqhkiG9w0BCQQxFgQU
# 0JgyUUiJZde43RUCwPdwuHu92okwDQYJKoZIhvcNAQEFBQAEggEAPmpJtoQtccx+
# QU4KC4GolDy7atsJRzSGqir91bWM7iFb5RDayhqn489UpIdJQPHXCooHAsuKXJZ2
# eSlfUygfYuAVnHjJB3LSElyguqvGad5L7njXYeEf7ljeJGwjY/LWErPPsAj4ECCB
# E51qTsdxofikftJZ/guD6Py6aF5DV8I9ErvtyfsfACg2X5zcE/ayEb5bPQ8oFF5K
# tdklu8sg/wG+unuEW07ckOxiFmbcOsw2RDpJhYHsVLkC2UxMCcAChH+C0NNbbS6N
# RGvjccwWP+iuWJAXTiW/t3rFed9944FRxPgz8dIFMyepWi/Jdk0n68x/gptIa5IG
# fag4Bgaobg==
# SIG # End signature block
