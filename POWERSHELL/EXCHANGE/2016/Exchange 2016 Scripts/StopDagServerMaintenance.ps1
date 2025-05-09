# Copyright (c) Microsoft Corporation. All rights reserved.
#
# StopDagServerMaintenance


<#
 .SYNOPSIS
 Resumes the things that were suspended by StartDagServerMaintenance, but does 
 NOT move resources back to the server.

 Can be run remotely, but it requires the cluster administrative tools to
 be installed (RSAT-Clustering).

.PARAMETER serverName
The name of the server on which to end maintenance. FQDNs are valid.

.PARAMETER SwitchedOver
Whether the script should mark machine as switched over or not.

.PARAMETER TakeTraffic
Whether the script should configure the component to take traffic or not.

.PARAMETER whatif
Does not actually perform any operations, but logs what would be executed
to the verbose stream.

.EXAMPLE
c:\ps> .\StopDagServerMaintenance.ps1 -server foo

.EXAMPLE
c:\ps> .\StopDagServerMaintenance.ps1 -server foo.contoso.com
#>


Param(
	[Parameter(Mandatory=$true)]
	[System.Management.Automation.ValidateNotNullOrEmptyAttribute()]
	[string] $serverName,

	[Parameter(Mandatory=$false)] [switch] $whatif = $false,
	
	[Parameter(Mandatory=$false)] [switch] $setDatabaseCopyActivationDisabledAndMoveNow = $false,

	[Parameter(Mandatory=$false)] [System.Nullable``1[[System.Boolean]]] $takeTraffic = $null
)

$HAComponent = 'HighAvailability'

Import-LocalizedData -BindingVariable StopDagServerMaintenance_LocalizedStrings -FileName StopDagServerMaintenance.strings.psd1

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
	LoadExchangeSnapin

	# Load some of the common functions.	
	. "$(split-path $thisScriptName)\DagCommonLibrary.ps1";

	Test-RsatClusteringInstalled

	# Allow an FQDN to be passed in, but strip it to the short name.
	$shortServerName = $serverName;
	if ( $shortServerName.Contains( "." ) )
	{
		$shortServerName = $shortServerName -replace "\..*$"
	}

	# if $takeTraffic is $null I want both conditions to be executed. 
	#
	# condition 1 - executes cluster resume
	# condition 2 - executes AD operation to activation unblock and for movenow flags.
	#
	# if $takeTraffic is $null both condition 1 and condition 2 are true
	# if $takeTraffic is $false only condition 1
	# if $takeTraffic is $true only condition 2
	# 
	# $takeTraffic will be $null if newer script is called by older version of workflow and thus will have exactly same behaviour as before split
	# 
	# in new workflow $takeTraffic -eq $false will be before patches are applied and before monitoring is started and $takeTraffic -eq $true is after patches are applied and monitoring is started.
	if ($takeTraffic -eq $null -or $takeTraffic -eq $false)
	{
		# Resume the cluster node before doing any of the database operations. If
		# the script fails in the middle, people are more likely to notice
		# suspended database copies than a Paused cluster node.

		# Explicitly connect to clussvc running on serverName. This script could
		# easily be run remotely.
		log-verbose ($StopDagServerMaintenance_LocalizedStrings.res_0000 -f $serverName,$shortServerName,$serverName);
		if ( $DagScriptTesting )
		{
			write-host ($StopDagServerMaintenance_LocalizedStrings.res_0001 )
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

			$outputStruct = Call-ClusterExe -dagName $dagName -serverName $serverName -clusterCommand "node $serverName /resume"
			$LastExitCode = $outputStruct[ 0 ];
			$output = $outputStruct[ 1 ];

			# 0 is success, 5058 is ERROR_CLUSTER_NODE_NOT_PAUSED.
			if ( $LastExitCode -eq 5058 )
			{
				log-verbose ($StopDagServerMaintenance_LocalizedStrings.res_0002 -f $serverName)
				# clear $LastExitCode
				cmd /c "exit 0"
			}
			elseif ( $LastExitCode -eq 1753 )
			{
				# 1753 is EPT_S_NOT_REGISTERED
				log-verbose ($StopDagServerMaintenance_LocalizedStrings.res_0003 -f $severName,"Stop-DagServerMaintenance")
				# clear $LastExitCode
				cmd /c "exit 0"
			}
			elseif ( $LastExitCode -eq 1722 )
			{
				# 1722 is RPC_S_SERVER_UNAVAILABLE
				log-verbose ($StopDagServerMaintenance_LocalizedStrings.res_0004 -f $severName,"Stop-DagServerMaintenance")
				# clear $LastExitCode
				cmd /c "exit 0"
			}
			elseif ( $LastExitCode -ne 0 )
			{
				Log-Error ($StopDagServerMaintenance_LocalizedStrings.res_0008 -f $serverName,$serverName,$shortServerName,$LastExitCode,"Start-DagServerMaintenance") -stop 
			}
		}
	}
	
	if ($takeTraffic -eq $null -or $takeTraffic -eq $true)
	{
		log-verbose ($StopDagServerMaintenance_LocalizedStrings.res_0007 -f $shortServerName)
	
		if ( $setDatabaseCopyActivationDisabledAndMoveNow )
		{
			write-host ($StopDagServerMaintenance_LocalizedStrings.res_0009 -f $shortServerName)
		
			if ( $DagScriptTesting )
			{
				write-host ($StopDagServerMaintenance_LocalizedStrings.res_0006 -f "Set-MailboxServer -Identity '$shortServerName' -DatabaseCopyActivationDisabledAndMoveNow:`$true")
			}
			else
			{
				Set-MailboxServer -Identity $shortServerName -DatabaseCopyActivationDisabledAndMoveNow:$true
			}
		}
		else
		{
			write-host ($StopDagServerMaintenance_LocalizedStrings.res_0010 -f $shortServerName)
		
			if ( $DagScriptTesting )
			{
				write-host ($StopDagServerMaintenance_LocalizedStrings.res_0006 -f "Set-MailboxServer -Identity '$shortServerName' -DatabaseCopyActivationDisabledAndMoveNow:`$false")
			}
			else
			{
				Set-MailboxServer -Identity $shortServerName -DatabaseCopyActivationDisabledAndMoveNow:$false
			}
		}

		if ( $DagScriptTesting )
		{
			write-host ($StopDagServerMaintenance_LocalizedStrings.res_0006 -f "set-mailboxserver")
		}
		else
		{
			Set-MailboxServer -Identity $shortServerName -DatabaseCopyAutoActivationPolicy:Unrestricted
		}
	
		if ( $DagScriptTesting )
		{
			write-host ($StopDagServerMaintenance_LocalizedStrings.res_0011 -f "Set-ServerComponentState")
		}
		else
		{
			$components = @(Get-ServerComponentState $serverName | where {$_.Component -ilike $HAComponent})
			
			if ($components.Count -gt 0)
			{
				Set-ServerComponentState $serverName -Component $HAComponent -Requester "Maintenance" -State Active
			}
			else
			{
				Log-Warning ($StopDagServerMaintenance_LocalizedStrings.res_0013 -f $HAComponent, $shortServerName)
			}			
		}
	
		# Best effort resume copies for backward compatibility of machines who were activation suspended with the previous build.
		try
		{
			$databases = Get-MailboxDatabase -Server $shortServerName | where { $_.ReplicationType -eq 'Remote' }

			if ( $databases )
			{
				# 1. Resume database copy. This clears the ActivationOnly suspension.
				foreach ( $database in $databases )
				{
					log-verbose ($StopDagServerMaintenance_LocalizedStrings.res_0005 -f $database.Name,$shortServerName);
					if ( $DagScriptTesting )
					{
						write-host ($StopDagServerMaintenance_LocalizedStrings.res_0006 -f "resume-mailboxdatabasecopy")
					}
					else
					{
						Resume-MailboxDatabaseCopy "$($database.Name)\$shortServerName" -Confirm:$false
					}
				}
			}
		}
		catch
		{
			Log-Warning ($StopDagServerMaintenance_LocalizedStrings.res_0012 -f $shortServerName, $_);
        	$Error.Clear()
		}
	}
}

# SIG # Begin signature block
# MIIduAYJKoZIhvcNAQcCoIIdqTCCHaUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUJDuhFywMjx0C9s1eqkw7ZAnV
# s3GgghhkMIIEwzCCA6ugAwIBAgITMwAAAJ1CaO4xHNdWvQAAAAAAnTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTMw
# WhcNMTcwNjMwMTkyMTMwWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjE0OEMtQzRCOS0yMDY2MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAy8PvNqh/8yl1
# MrZGvO1190vNqP7QS1rpo+Hg9+f2VOf/LWTsQoG0FDOwsQKDBCyrNu5TVc4+A4Zu
# vqN+7up2ZIr3FtVQsAf1K6TJSBp2JWunjswVBu47UAfP49PDIBLoDt1Y4aXzI+9N
# JbiaTwXjos6zYDKQ+v63NO6YEyfHfOpebr79gqbNghPv1hi9thBtvHMbXwkUZRmk
# ravqvD8DKiFGmBMOg/IuN8G/MPEhdImnlkYFBdnW4P0K9RFzvrABWmH3w2GEunax
# cOAmob9xbZZR8VftrfYCNkfHTFYGnaNNgRqV1rEFt866re8uexyNjOVfmR9+JBKU
# FbA0ELMPlQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFGTqT/M8KvKECWB0BhVGDK52
# +fM6MB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAD9dHEh+Ry/aDJ1YARzBsTGeptnRBO73F/P7wF8dC7nTPNFU
# qtZhOyakS8NA/Zww74n4gvm1AWfHGjN1Ao8NiL3J6wFmmON/PEUdXA2zWFYhgeRe
# CPmATbwNN043ecHiGjWO+SeMYpvl1G4ma0NIUJau9DmTkfaMvNMK+/rNljr3MR8b
# xsSOZxx2iUiatN0ceMmIP5gS9vUpDxTZkxVsMfA5n63j18TOd4MJz+G0I62yqIvt
# Yy7GTx38SF56454wqMngiYcqM2Bjv6xu1GyHTUH7v/l21JBceIt03gmsIhlLNo8z
# Ii26X6D1sGCBEZV1YUyQC9IV2H625rVUyFZk8f4wggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBL4wggS6AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB0jAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUMo93aXLp/59U7jBYntsiXT6UABcwcgYKKwYB
# BAGCNwIBDDFkMGKgOoA4AFMAdABvAHAARABhAGcAUwBlAHIAdgBlAHIATQBhAGkA
# bgB0AGUAbgBhAG4AYwBlAC4AcABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQAuMUAf7xRooQOKtlyJMkEU
# 6rwU5vg9a76bSj+QwF+qbebLH8sBNVlYn4paM71gp/jb2NoMNqZ78HurRSETpXr9
# 4QATlBcwKq/rQEVjih5kCjqi8j1gagAWw77fIGZQQ2lmTHZrEVHrt1CkwprBI+9c
# rJs6wLC0xtmkj/WFMBG2wziEoukEG8NK+tCZrOHO8TAEmJoikCwuvPoQTmtMaIGn
# S0uaJ6oU1DQD/BnWTsYcZo0V4STP02Gs2RKIV+NN0PhjT3n5sMUeEedfHhZ3ONFW
# 5/EXMRvskMpYKczVsAEV2CX/bj2iQ+VBhg9VV4vWiWbrZWWi8cUXIyDXPJ7oqgOe
# oYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVT
# MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQK
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1l
# LVN0YW1wIFBDQQITMwAAAJ1CaO4xHNdWvQAAAAAAnTAJBgUrDgMCGgUAoF0wGAYJ
# KoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTYwOTAzMTg0
# NDA2WjAjBgkqhkiG9w0BCQQxFgQUMZcU+89Kfk5t//PS1KPQ+4cVhv8wDQYJKoZI
# hvcNAQEFBQAEggEAmBxw/Qw+rNLoX+J4rjm31q7cdFAwsMW+pqerVcIOpF1eGB0G
# OOWisMBGFFHrkqNCWFmWwgFrsGBl9OtxD6ygycY/s98WH9ceLELcBvCa8dc4hdv1
# BDdzoldlYd7az3Q2xhRGfyP50AZPdh615nhml/CsiMr3G8T5V7q4I2W2XwP0Kyax
# hEAgfXzK5GDJDiFUm4+YkuL6lhBdli+vFCIWthd0JAhsRljcJsvCqG/SBrAYv//I
# Gtzz3avSnPhbUzGk6Pf4RlVM/bHLmCL567V2cF3c/Y58HnMYjkRl1Vd7/ePGiFD+
# IJJICQ6aQ97RKTQL8FFks/MbY+t27nuqvMJnfw==
# SIG # End signature block
