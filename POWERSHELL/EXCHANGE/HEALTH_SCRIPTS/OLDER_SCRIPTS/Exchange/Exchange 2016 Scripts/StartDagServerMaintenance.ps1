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

# .PARAMETER MoveComment
# The string which is passed to the MoveComment parameter of the
# Move-ActiveMailboxDatabase cmdlet.

Param(
	[Parameter(Mandatory=$true)]
	[System.Management.Automation.ValidateNotNullOrEmptyAttribute()]
	[string] $serverName,

	[string] $Force = 'false',
	[Parameter(Mandatory=$false)] [switch] $whatif = $false,
	[Parameter(Mandatory=$false)] [switch] $overrideMinimumTwoCopies = $false,
    [Parameter(Mandatory=$false)] [string] $MoveComment = "BeginMaintenance"
)

# Global Values
$ServerCountinTwoServerDAG = 2
$RetryCount = 2
$HAComponent = 'HighAvailability'

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

# Handle Cluster Error Codes during Start-DagServerMaintenance
function HandleClusterErrorCode ([string]$Server = $servername, [int]$ClusterErrorCode, [string]$Action)
{
	switch ($ClusterErrorCode)
	{
		# 0 is success
		0		
		{   
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0026 -f $Server,$Action,"Start-DagServerMaintenance")
			# clear $LastExitCode
			cmd /c "exit 0"
		}
        
        # 5 is returned when the Server is powered down 
		5
		{
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0025 -f $Server,$Action,"Server powered down","Start-DagServerMaintenance")
			# clear $LastExitCode
			cmd /c "exit 0"
		}
        
        # 70 is ERROR_SHARING_PAUSED - The remote server has been paused or is in the process of being started
		70
		{
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0025 -f $Server,$Action,"ERROR_SHARING_PAUSED","Start-DagServerMaintenance")
			# clear $LastExitCode
			cmd /c "exit 0"
		}
		
		# 1753 is EPT_S_NOT_REGISTERED
		1753
		{
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0025 -f $Server,$Action,"EPT_S_NOT_REGISTERED","Start-DagServerMaintenance")
			# clear $LastExitCode
			cmd /c "exit 0"
		}
		
		# 1722 is RPC_S_SERVER_UNAVAILABLE
		1722
		{
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0025 -f $Server,$Action,"RPC_S_SERVER_UNAVAILABLE","Start-DagServerMaintenance")
			# clear $LastExitCode
			cmd /c "exit 0"
		}
		
		# 5042 is ERROR_CLUSTER_NODE_NOT_FOUND
		5042
		{
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0025 -f $Server,$Action,"ERROR_CLUSTER_NODE_NOT_FOUND","Start-DagServerMaintenance")
			# clear $LastExitCode
			cmd /c "exit 0"
		}
		
		# 5043 is ERROR_CLUSTER_LOCAL_NODE_NOT_FOUND
		5043
		{
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0025 -f $Server,$Action,"ERROR_CLUSTER_LOCAL_NODE_NOT_FOUND","Start-DagServerMaintenance")
			# clear $LastExitCode
			cmd /c "exit 0"
		}
		
		# 5050 is ERROR_CLUSTER_NODE_DOWN
		5050
		{
			log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0025 -f $Server,$Action,"ERROR_CLUSTER_NODE_DOWN","Start-DagServerMaintenance")
			# clear $LastExitCode
			cmd /c "exit 0"
		}
		
		# Best effort to pause the node. Not a known code, so warning at this point.
		default {Log-Warning ($StartDagServerMaintenance_LocalizedStrings.res_0004 -f $Server,$Action,$ClusterErrorCode,"Start-DagServerMaintenance") -stop}
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
	$serverComponentStateSet = $false;
    $serverComponentStateOrg = $null;
	$scriptCompletedSuccessfully = $false;

	try {
        # Stage 1 - block auto activation on the server.
        # Also set HA server component state to Inactive (in AD and registry). This will block actives from moving back to the server.
        
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
		
		log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0029 -f $shortServerName)
		if ($DagScriptTesting)
		{
			write-host ($StartDagServerMaintenance_LocalizedStrings.res_0030 -f $shortServerName,"Set-ServerComponentState","-Component HighAvailability","-State Inactive")
		}
		else
		{
            # Get-ServerComponentState test to see if HighAvailability component exists on the server
            # If it doesn't exist skip calling Set-ServerComponentState
            # For all other cases suppress the error and let Set-ServerComponentState run
            $componentExists = $true
            try
            {
                $Error.Clear()
			    $serverComponentStateOrg = Get-ServerComponentState $serverName -Component $HAComponent -ErrorAction:Stop;
            }
            catch
            {
                if ($Error.Exception.Gettype().Name -ilike 'ArgumentException')
                {
                    $componentExists = $false
                }
                $Error.Clear()
            }
            
			
			if ($componentExists)
			{
                if($serverComponentStateOrg -and $serverComponentStateOrg.State -eq 'Active')
                {
				    Set-ServerComponentState $serverName -Component $HAComponent -Requester "Maintenance" -State Inactive
				    $serverComponentStateSet = $true;
                }
			}
			else
			{
				Log-Warning ($StartDagServerMaintenance_LocalizedStrings.res_0034 -f $HAComponent, $shortServerName)
			}
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
			HandleClusterErrorCode -ClusterErrorCode $LastExitCode -Action "Pause"
            $Error.Clear()
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
		
		# Move the critical resources off the specified server. 
		# This includes Active Databases, and the Primary Active Manager.
		# If any error occurs in this stage, script execution will halt.
		# (If we don't assign the result to a variable then the script will
		# print out 'True')
		$try = 0
		$dagObject = Get-DatabaseAvailabilityGroup $dagName			
		$dagServers = $dagObject.Servers.Count			
		$stoppedDagServers = $dagObject.StoppedMailboxServers.Count
        $ignoredResult = ""
		while (($numCriticalResources -gt 0) -and ($try -lt $RetryCount))
		{
			# Sleep for 60 seconds if this is not the first move attempt
			if ($try -gt 0)
			{
				Sleep-ForSeconds 60
			}
			
			$ignoredResult = Move-CriticalMailboxResources -Server $shortServerName -MoveComment $MoveComment -Force $Force
		
			# Check again to see if the moves were successful. (Unless -whatif was
			# specified, then it's pretty likely it will fail).
			if ( !$DagScriptTesting )
			{
				log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0013 -f $shortServerName,"Start-DagServerMaintenance")			
						
				if (($dagServers - $stoppedDagServers) -eq $ServerCountinTwoServerDAG -or $overrideMinimumTwoCopies)
				{				
					$criticalMailboxResources = @(GetCriticalMailboxResources $shortServerName -AtleastNCriticalCopies ($ServerCountinTwoServerDAG - 1))
				}
				else
				{			
					$criticalMailboxResources = @(GetCriticalMailboxResources $shortServerName)	
				}			
				$numCriticalResources = ($criticalMailboxResources | Measure-Object).Count
			}
			$try++
		}
		if( $numCriticalResources -gt 0 )
		{
			log-warning "MoveResults follow:`n$ignoredResult`nEndOfMoveResults"
			Log-CriticalResource $criticalMailboxResources
			
            if($ignoredResult.Contains("AmDbMoveMoveSuppressedBlackoutException"))
            {
                Write-Error ($StartDagServerMaintenance_LocalizedStrings.res_0035) -ErrorAction:Stop
            }
            else
            {
			    write-error ($StartDagServerMaintenance_LocalizedStrings.res_0014 -f ( PrintCriticalMailboxResourcesOutput($criticalMailboxResources)),$shortServerName, $ignoredResult) -erroraction:stop 
            }
		}
		$scriptCompletedSuccessfully = $true;
	}
	finally
	{
		# Rollback only if something failed and Force flag was not used
		if ( !$scriptCompletedSuccessfully)
		{
			if ($Force -ne 'true')
			{
				log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0015 -f "Start-DagServerMaintenance")

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
							HandleClusterErrorCode -ClusterErrorCode $LastExitCode -Action "Resume"
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
					
					if ( $serverComponentStateSet -and $serverComponentStateOrg -and $serverComponentStateOrg.State -eq 'Active')
					{
						log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0031 -f $shortServerName)
						
						if ( $DagScriptTesting )
						{
							write-host ($StartDagServerMaintenance_LocalizedStrings.res_0032 -f "Set-ServerComponentState")
						}
						else
						{
							Set-ServerComponentState $serverName -Component 'HighAvailability' -Requester "Maintenance" -State Active							
						}
					}					
				}
			}
			else
			{
				log-verbose ($StartDagServerMaintenance_LocalizedStrings.res_0027 -f "Start-DagServerMaintenance")
			}
		}		
	}
}

# SIG # Begin signature block
# MIIdugYJKoZIhvcNAQcCoIIdqzCCHacCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUyu98dain3spwHCE3slmFj7l/
# 1tygghhkMIIEwzCCA6ugAwIBAgITMwAAAJvgdDfLPU2NLgAAAAAAmzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI5
# WhcNMTcwNjMwMTkyMTI5WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjcyOEQtQzQ1Ri1GOUVCMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAjaPiz4GL18u/
# A6Jg9jtt4tQYsDcF1Y02nA5zzk1/ohCyfEN7LBhXvKynpoZ9eaG13jJm+Y78IM2r
# c3fPd51vYJxrePPFram9W0wrVapSgEFDQWaZpfAwaIa6DyFyH8N1P5J2wQDXmSyo
# WT/BYpFtCfbO0yK6LQCfZstT0cpWOlhMIbKFo5hljMeJSkVYe6tTQJ+MarIFxf4e
# 4v8Koaii28shjXyVMN4xF4oN6V/MQnDKpBUUboQPwsL9bAJMk7FMts627OK1zZoa
# EPVI5VcQd+qB3V+EQjJwRMnKvLD790g52GB1Sa2zv2h0LpQOHL7BcHJ0EA7M22tQ
# HzHqNPpsPQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFJaVsZ4TU7pYIUY04nzHOUps
# IPB3MB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBACEds1PpO0aBofoqE+NaICS6dqU7tnfIkXIE1ur+0psiL5MI
# orBu7wKluVZe/WX2jRJ96ifeP6C4LjMy15ZaP8N0OckPqba62v4QaM+I/Y8g3rKx
# 1l0okye3wgekRyVlu1LVcU0paegLUMeMlZagXqw3OQLVXvNUKHlx2xfDQ/zNaiv5
# DzlARHwsaMjSgeiZIqsgVubk7ySGm2ZWTjvi7rhk9+WfynUK7nyWn1nhrKC31mm9
# QibS9aWHUgHsKX77BbTm2Jd8E4BxNV+TJufkX3SVcXwDjbUfdfWitmE97sRsiV5k
# BH8pS2zUSOpKSkzngm61Or9XJhHIeIDVgM0Ou2QwggYHMIID76ADAgECAgphFmg0
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
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCBhAwggP4
# oAMCAQICEzMAAABkR4SUhttBGTgAAAAAAGQwDQYJKoZIhvcNAQELBQAwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xNTEwMjgyMDMxNDZaFw0xNzAx
# MjgyMDMxNDZaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCTLtrY5j6Y2RsPZF9NqFhN
# FDv3eoT8PBExOu+JwkotQaVIXd0Snu+rZig01X0qVXtMTYrywPGy01IVi7azCLiL
# UAvdf/tqCaDcZwTE8d+8dRggQL54LJlW3e71Lt0+QvlaHzCuARSKsIK1UaDibWX+
# 9xgKjTBtTTqnxfM2Le5fLKCSALEcTOLL9/8kJX/Xj8Ddl27Oshe2xxxEpyTKfoHm
# 5jG5FtldPtFo7r7NSNCGLK7cDiHBwIrD7huTWRP2xjuAchiIU/urvzA+oHe9Uoi/
# etjosJOtoRuM1H6mEFAQvuHIHGT6hy77xEdmFsCEezavX7qFRGwCDy3gsA4boj4l
# AgMBAAGjggF/MIIBezAfBgNVHSUEGDAWBggrBgEFBQcDAwYKKwYBBAGCN0wIATAd
# BgNVHQ4EFgQUWFZxBPC9uzP1g2jM54BG91ev0iIwUQYDVR0RBEowSKRGMEQxDTAL
# BgNVBAsTBE1PUFIxMzAxBgNVBAUTKjMxNjQyKzQ5ZThjM2YzLTIzNTktNDdmNi1h
# M2JlLTZjOGM0NzUxYzRiNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzcitW2oynUC
# lTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEGCCsGAQUF
# BwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0MAwGA1Ud
# EwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjiDGRDHd1crow7hSS1nUDWvWas
# W1c12fToOsBFmRBN27SQ5Mt2UYEJ8LOTTfT1EuS9SCcUqm8t12uD1ManefzTJRtG
# ynYCiDKuUFT6A/mCAcWLs2MYSmPlsf4UOwzD0/KAuDwl6WCy8FW53DVKBS3rbmdj
# vDW+vCT5wN3nxO8DIlAUBbXMn7TJKAH2W7a/CDQ0p607Ivt3F7cqhEtrO1Rypehh
# bkKQj4y/ebwc56qWHJ8VNjE8HlhfJAk8pAliHzML1v3QlctPutozuZD3jKAO4WaV
# qJn5BJRHddW6l0SeCuZmBQHmNfXcz4+XZW/s88VTfGWjdSGPXC26k0LzV6mjEaEn
# S1G4t0RqMP90JnTEieJ6xFcIpILgcIvcEydLBVe0iiP9AXKYVjAPn6wBm69FKCQr
# IPWsMDsw9wQjaL8GHk4wCj0CmnixHQanTj2hKRc2G9GL9q7tAbo0kFNIFs0EYkbx
# Cn7lBOEqhBSTyaPS6CvjJZGwD0lNuapXDu72y4Hk4pgExQ3iEv/Ij5oVWwT8okie
# +fFLNcnVgeRrjkANgwoAyX58t0iqbefHqsg3RGSgMBu9MABcZ6FQKwih3Tj0DVPc
# gnJQle3c6xN3dZpuEgFcgJh/EyDXSdppZzJR4+Bbf5XA/Rcsq7g7X7xl4bJoNKLf
# cafOabJhpxfcFOowMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkqhkiG9w0B
# AQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAG
# A1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTEw
# HhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
# aWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# q/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03a8YS2Avw
# OMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akrrnoJr9eW
# WcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0RrrgOGSsbmQ1
# eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy4BI6t0le
# 2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9sbKvkjh+
# 0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAhdCVfGCi2
# zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8kA/DRelsv
# 1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTBw3J64HLn
# JN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmnEyimp31n
# gOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90lfdu+Hgg
# WCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0wggHpMBAG
# CSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2oynUClTAZ
# BgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/
# BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBaBgNVHR8E
# UzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9k
# dWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsGAQUFBwEB
# BFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNVHSAEgZcw
# gZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsGAQUFBwIC
# MDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABlAG0AZQBu
# AHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKbC5YR4WOS
# mUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11lhJB9i0ZQ
# VdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6I/MTfaaQ
# dION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0wI/zRive
# /DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560STkKxgrC
# xq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQamASooPoI/
# E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGaJ+HNpZfQ
# 7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ahXJbYANah
# Rr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA9Z74v2u3
# S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33VtY5E90Z1W
# Tk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr/Xmfwb1t
# bWrJUnMTDXpQzTGCBMAwggS8AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB1DAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQU2Ikxd/CxQH6HF3/8SH1ol5TAZZowdAYKKwYB
# BAGCNwIBDDFmMGSgPIA6AFMAdABhAHIAdABEAGEAZwBTAGUAcgB2AGUAcgBNAGEA
# aQBuAHQAZQBuAGEAbgBjAGUALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBAI/xZvO9Ze1d7sBWjTHr
# Aep0u41EEcbOmMmE6cnc+9LGmHxA/IPrsXinPaOnc22XxAu/3K/cCtzypfNWHKf8
# iqPtMicWyliHYq73jKMERlM99JQIVLgO2ocq1fdCrdstFo5zQdXn0rttH8Ij+HSc
# Kd7mIV2aiqV2xLovSLwscDpPmsJ3SnKOmoI0v+IFk7mWycmJVzCPPlbIUGTMG2gp
# R+0EyJMmKWBMP5U1fEWgaVmJLZcvHGr8OLNvYwyfcNzzfFHch8U03/Y9xEuHYTDX
# VrHw1E750MZXBohKm33rWGdZds5YkfJcuJKxN3oBy9yxTqSryD8jrP3OWz9dCz40
# m2mhggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBAhMzAAAAm+B0N8s9TY0uAAAAAACbMAkGBSsOAwIaBQCgXTAY
# BgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMx
# ODQ0MDZaMCMGCSqGSIb3DQEJBDEWBBQGKVdqvmVgSUCrKvV3TjA5oWyGPDANBgkq
# hkiG9w0BAQUFAASCAQBqmifOwpcXRVjSjxIrm0rqHYwLqSJjCa5JjOxxj9l5zjBJ
# Ua7UAfIHCINHuHCXTBZe3gfq3V+gIn/xFRuyFz9RuO5QvJZ8mm26RedmeCszarHi
# +l4s5tHacstnwXsz4ouUIdEIgMV4UoNaSvuX2O6DsNANjl4+AYvF8L20FAD7SySJ
# u3uvtUzkWdrf/Py3otEaO00Jhd2+MaVHTsxgdgqdEND3Oik++EHTL8NZ1QIMuIQs
# t9ibe67GZqPfRWiUaN0VkIanjreYZhAT6G4PyeklipiQMk3neMZSDFph7iyebVV6
# eimx40oNyTx2Er6h3KYo/mksXEv95DinYKA3cBH0
# SIG # End signature block
