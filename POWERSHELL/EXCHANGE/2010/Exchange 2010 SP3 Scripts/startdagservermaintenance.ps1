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
# MIIabgYJKoZIhvcNAQcCoIIaXzCCGlsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUK7wVIc7kbgH4uBWJvG2UumkP
# jhSgghUmMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
# 9w0BAQUFADB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSMw
# IQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQTAeFw0xMjA5MDQyMTQy
# MDlaFw0xMzAzMDQyMTQyMDlaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC6pElsEPsi
# nGWiFpg7y2Fi+nQprY0GGdJxWBmKXlcNaWJuNqBO/SJ54B3HGmGO+vyjESUWyMBY
# LDGKiK4yHojbfz50V/eFpDZTykHvabhpnm1W627ksiZNc9FkcbQf1mGEiAAh72hY
# g1tJj7Tf0zXWy9kwn1P8emuahCu3IWd01PZ4tmGHmJR8Ks9n6Rm+2bpj7TxOPn0C
# 6/N/r88Pt4F+9Pvo95FIu489jMgHkxzzvXXk/GMgKZ8580FUOB5UZEC0hKo3rvMA
# jOIN+qGyDyK1p6mu1he5MPACIyAQ+mtZD+Ctn55ggZMDTA2bYhmzu5a8kVqmeIZ2
# m2zNTOwStThHAgMBAAGjggENMIIBCTATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQU3lHcG/IeSgU/EhzBvMOzZSyRBZgwHwYDVR0jBBgwFoAUyxHoytK0FlgB
# yTcuMxYWuUyaCh8wVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljQ29kU2lnUENBXzA4LTMxLTIwMTAu
# Y3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNDb2RTaWdQQ0FfMDgtMzEtMjAxMC5jcnQw
# DQYJKoZIhvcNAQEFBQADggEBACqk9+7AwyZ6g2IaeJxbxf3sFcSneBPRF1MoCwwA
# Qj84D4ncZBmENX9Iuc/reomhzU+p4LvtRxD+F9qHiRDRTBWg8BH/2pbPZM+B/TOn
# w3iT5HzVbYdx1hxh4sxOZLdzP/l7JzT2Uj9HQ8AOgXBTwZYBoku7vyoDd3tu+9BG
# ihcoMaUF4xaKuPFKaRVdM/nff5Q8R0UdrsqLx/eIHur+kQyfTwcJ7SaSbrOUGQH4
# X4HnrtqJj39aXoRftb58RuVHr/5YK5F/h9xGH1GVzMNiobXHX+vJaVxxkamNViAs
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggS6MIIDoqADAgEC
# AgphAo5CAAAAAAAfMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQTAeFw0xMjAxMDkyMjI1NThaFw0xMzA0MDkyMjI1NThaMIGzMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMScwJQYD
# VQQLEx5uQ2lwaGVyIERTRSBFU046RjUyOC0zNzc3LThBNzYxJTAjBgNVBAMTHE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQCW7I5HTVTCXJWA104LPb+XQ8NL42BnES8BTQzY0UYvEEDeC6RQ
# UhKIC0N6LT/uSG5mx5HmA8pu7HmpaiObzWKezWqkP+ejQ/9iR6G0ukT630DBhVR+
# 6KCnLEMjm1IfMjX0/ppWn41jd3swngozhXIbykrIzCXN210RLsewjPGPQ0hHBbV6
# IAvl8+/BuvSz2M04j/shqj0KbYUX0MrnhgPAM4O1JcTMWpzEw9piJU1TJRRhj/sb
# 4Oz3R8aAReY1UyM2d8qw3ZgrOcB1NQ/dgUwhPXYwxbKwZXMpSCfYwtKwhEe7eLrV
# dAPe10sZ91PeeNqG92GIJjO0R8agVIgVKyx1AgMBAAGjggEJMIIBBTAdBgNVHQ4E
# FgQUL+hGyGjTbk+yINDeiU7xR+5IwfIwHwYDVR0jBBgwFoAUIzT42VJGcArtQPt2
# +7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3NvZnQu
# Y29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNybDBY
# BggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNVHSUE
# DDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAc/99Lp3NjYgrfH3jXhVx
# 6Whi8Ai2Q1bEXEotaNj5SBGR8xGchewS1FSgdak4oVl/de7G9TTYVKTi0Mx8l6uT
# dTCXBx0EUyw2f3/xQB4Mm4DiEgogOjHAB3Vn4Po0nOyI+1cc5VhiIJBFL11FqciO
# s3xybRAnxUvYb6KoErNtNSNn+izbJS25XbEeBedDKD6cBXZ38SXeBUcZbd5JhaHa
# SksIRiE1qHU2TLezCKrftyvZvipq/d81F8w/DMfdBs9OlCRjIAsuJK5fQ0QSelzd
# N9ukRbOROhJXfeNHxmbTz5xGVvRMB7HgDKrV9tU8ouC11PgcfgRVEGsY9JHNUaeV
# ZTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEFBQAwXzETMBEG
# CgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsG
# A1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5MB4XDTEw
# MDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENvZGUgU2lnbmlu
# ZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCycllcGTBkvx2a
# YCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPVcgDbNVcKicqu
# IEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlcRdyvrT3gKGiX
# GqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZC/6SdCnidi9U
# 3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgGhVxOVoIoKgUy
# t0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdcpReejcsRj1Y8
# wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTL
# EejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYBBAGCNxUBBAUC
# AwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy82C0wGQYJKwYB
# BAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBWJ5flJRP8KuEK
# U5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNyb3NvZnQuY29t
# L3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3JsMFQGCCsGAQUF
# BwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcNAQEFBQADggIB
# AFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGjI8x8UJiAIV2s
# PS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbNLeNK0rxw56gN
# ogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y4k74jKHK6BOl
# kU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnpo1hW3ZsCRUQv
# X/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6H0q70eFW6NB4
# lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20OE049fClInHLR
# 82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8Z4L5UrKNMxZl
# Hg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9GuwdgR2VgQE6w
# QuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXrilUEnacOTj5XJ
# jdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEvmtzjcT3XAH5i
# R9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCCA++gAwIBAgIK
# YRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPyLGQBGRYDY29t
# MRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRNaWNyb3NvZnQg
# Um9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1MzA5WhcNMjEw
# NDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZUSNQrc7dGE4k
# D+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cOBJjwicwfyzMk
# h53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn1yjcRlOwhtDl
# KEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3U21StEWQn0gA
# SkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG7bfeI0a7xC1U
# n68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMBAAGjggGrMIIB
# pzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A+3b7syuwwzWz
# DzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1UdIwSBkDCBjYAU
# DqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/IsZAEZFgNjb20x
# GTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBS
# b290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0BxMuZTBQBgNV
# HR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9w
# cm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEESDBGMEQG
# CCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01p
# Y3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
# 9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwTq86+e4+4LtQS
# ooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+jwoFyI1I4vBT
# Fd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwYTp2Oawpylbih
# OZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPfwgphjvDXuBfr
# Tot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5ZlizLS/n+YWG
# zFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8csu89Ds+X57H21
# 46SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUwZuhCEl4ayJ4i
# IdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHudiG/m4LBJ1S2
# sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9La9Zj7jkIeW1
# sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g74TKIdbrHk/J
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0xggSy
# MIIErgIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQITMwAAAJ0ejSeu
# uPPYOAABAAAAnTAJBgUrDgMCGgUAoIHUMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBQz2QP36RokZzIBPNGDDmVShharhjB0BgorBgEEAYI3AgEMMWYwZKA8gDoAUwB0
# AGEAcgB0AEQAYQBnAFMAZQByAHYAZQByAE0AYQBpAG4AdABlAG4AYQBuAGMAZQAu
# AHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJ
# KoZIhvcNAQEBBQAEggEAcbubm/YDuzeamhJW1iuDPqLIQkAwupJv2RZbNQqxqEcg
# OkZEHlNqsnAWfIUFHuR5WmjB4DPxz0j7h9gvjMNrYFaynyZ/3Cfe4x07ikbbpQij
# Tro8yyoRyjchSw192g2VT/7uRalKTTutivDiO+JfEMkOInbfdFGH4T2BXbhzeep+
# z16N+Yp+g/atbSxnG0N4JndtmkaQGM++Kj5uhbJpqa80I1eDrxqB/MffaW+9A8zQ
# rwzdeDtZAzxeVw5XTEwsVBOhkTwGDPtjLZ4J1YbmYRQ7AIVk0Ryh7zi0GJE6fzcK
# cXJY3ZOGODyBGN3o6SrXMOfaLSDmnxXp1CwAIeh8zKGCAh8wggIbBgkqhkiG9w0B
# CQYxggIMMIICCAIBATCBhTB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECCmECjkIA
# AAAAAB8wCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJ
# KoZIhvcNAQkFMQ8XDTEzMDEwODA4NDcwMlowIwYJKoZIhvcNAQkEMRYEFODxo2z6
# bmSptRNrVmzT49pBVbNaMA0GCSqGSIb3DQEBBQUABIIBAAmfI5b9fpxYCRqyXY6L
# N2xBXAjzyihAeAep4L9ZyachOVmIu7oIkfOtmLyo+QChQ/HoZA7C6LfK6e+4/R+7
# 2r4vITQo1v55XOZcJSXow0mbt43ekpf+TDWd70ufXy1S9l6p+L9M9l020aYl4926
# PJ+DebYqMkvavFA9sMUsPT3HFOzyfM2xsqwaijY4M74EYDInco0Fj3Aob5y4aX3q
# zDgyPibhEBrbAoWO2XSSr65Ume8s47SBB2EMGb+yj1UKsX2OWl/yVs51cNe6yz7Z
# g1Xt4YbfbkG3s4Unr6Cuye/DXW5dkhMv3Fpasm0LgZf0qwjIiKIOJliA/8ABpljg
# j0Y=
# SIG # End signature block
