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

	[Parameter(Mandatory=$false)] [switch] $whatif = $false

)

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
		}
		elseif ( $LastExitCode -eq 1753 )
		{
			# 1753 is EPT_S_NOT_REGISTERED
			log-verbose ($StopDagServerMaintenance_LocalizedStrings.res_0003 -f $severName,"Stop-DagServerMaintenance")
		}
		elseif ( $LastExitCode -eq 1722 )
		{
			# 1722 is RPC_S_SERVER_UNAVAILABLE
			log-verbose ($StopDagServerMaintenance_LocalizedStrings.res_0004 -f $severName,"Stop-DagServerMaintenance")
		}
		elseif ( $LastExitCode -ne 0 )
		{
			Log-Error ($StopDagServerMaintenance_LocalizedStrings.res_0008 -f $serverName,$serverName,$shortServerName,$LastExitCode,"Start-DagServerMaintenance") -stop 
		}

	}

	# Get all databases with multiple copies.
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


	log-verbose ($StopDagServerMaintenance_LocalizedStrings.res_0007 -f $shortServerName)
	
	if ( $DagScriptTesting )
	{
		write-host ($StopDagServerMaintenance_LocalizedStrings.res_0006 -f "set-mailboxserver")
	}
	else
	{
		Set-MailboxServer -Identity $shortServerName -DatabaseCopyAutoActivationPolicy:Unrestricted
	}
}

# SIG # Begin signature block
# MIIabAYJKoZIhvcNAQcCoIIaXTCCGlkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUHXswbid3e+ESUFMFqqGTQs+f
# ubugghUmMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# AgphApJKAAAAAAAgMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQTAeFw0xMjAxMDkyMjI1NTlaFw0xMzA0MDkyMjI1NTlaMIGzMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMScwJQYD
# VQQLEx5uQ2lwaGVyIERTRSBFU046QjhFQy0zMEE0LTcxNDQxJTAjBgNVBAMTHE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQDNY8P3orVH2fk5lFOsa4+meTgh9NFuPRU4FsFzLJmP++A8W+Gi
# TIuyq4u08krHuZNQ0Sb1h1OwenXzRw8XK2csZOtg50qM4ItexFZpNzoqBF0XAci/
# qbfTyaujkaiB4z8tt9jkxHgAeTnP/cdk4iMSeJ4MdidJj5qsDnTjBlVVScjtoXxw
# RWF+snEWHuHBQbnS0jLpUiTzNlAE191vPEnVC9R2QPxlZRN+lVE61f+4mqdXa0sV
# 5TdYH9wADcWu6t+BR/nwUq7mz1qs4BRKpnGUgPr9sCMwYOmDYKwOJB7/EfLnpVVi
# le6ff54p1s1LCnaD8EME3wZJIPJTWkSwUXo5AgMBAAGjggEJMIIBBTAdBgNVHQ4E
# FgQUzRmsYU2UZyv2r5R1WdvwWACDoTowHwYDVR0jBBgwFoAUIzT42VJGcArtQPt2
# +7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3NvZnQu
# Y29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNybDBY
# BggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNVHSUE
# DDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAURzW3zYd3IC8AqfGmd8W
# yJAGaRWaMrnzZBDSivNiKRt5j8pgqbNdxOQWLTWMRJvx9FxhpEGumou6UNFqeWqz
# YoydTVFQC6mO1L+iEMBH1U51UokeP5Zqjy5AFAZy7j9wWZJdhe2DmfqChJ7kAwEh
# E6sn1sxxSeXaHf9vPAlO1Y1m6AzJf+4xFAI3X3tp7Ik+RX8lROcfGtbFGsNK5OHx
# hJjnT/mpmKcYRuyEbOypAwr9fHpSHZxrpKgPmJKkknhcK3jjfbLH2bZwfd9bc1O/
# qtRmUEvwyTuVheXBSmWdJVhQuyBUkXk6GwdcalcorzHHn+fDHe5H/SfXf8903GXF
# PzCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEFBQAwXzETMBEG
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
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0xggSw
# MIIErAIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQITMwAAAJ0ejSeu
# uPPYOAABAAAAnTAJBgUrDgMCGgUAoIHSMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBTtbn3+sDPy4rD7rn1LZOAep3hQ8zByBgorBgEEAYI3AgEMMWQwYqA6gDgAUwB0
# AG8AcABEAGEAZwBTAGUAcgB2AGUAcgBNAGEAaQBuAHQAZQBuAGEAbgBjAGUALgBw
# AHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqG
# SIb3DQEBAQUABIIBAAMbry9gbbB3fBdnDdXDpHU1ow8W8SnjR8nvBhVNLbmFC3ks
# tWFkxxLLfeDLr4ZXxzuCfhocw5A9MPYOOY6LoD3QcJwGZ6gwt4Qhk5xspGM96rTI
# +oPI6EMJqqdRnDQmd1EtAmvaYnfo44lutbmFZd8pJLNglrCksY2iLfQgH7l61kBT
# Wj1d9vxc7iE5z4ZIzM5L7CHUtKU0t4irTW7TWDZowTaHL0qFIz/d+z2fTEwyfyCo
# BmYuG7FewcfudJzA95GR3fdvqrFXbW83e7m+RHGCIyk93apuyObY+x9xBnDiAlRM
# q8Y/jEa0IOhoA7o2KAN2Z2qGojiZjD7Zr5slWLihggIfMIICGwYJKoZIhvcNAQkG
# MYICDDCCAggCAQEwgYUwdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAgphApJKAAAA
# AAAgMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqG
# SIb3DQEJBTEPFw0xMzAxMDgwODQ3MDFaMCMGCSqGSIb3DQEJBDEWBBTwmh3HR1xc
# zbh8oSEyBdoybetBJDANBgkqhkiG9w0BAQUFAASCAQCGLeK6biaWCgGJovLaF098
# tFqILp94FKx7G58JD8wqiHtrcKfALCji7ZFqWLmo96hyjIkdzGNT0tipZoPc3STV
# 4M89VHWc1TQ5nOFuETa4ubVgqW+tgbZwwzRt4icXbX39IPRMVArkA6JQ/470OfK8
# l8xKK6NWbMxS29LlKKXA3IGuz880gQH7OsYv+xRHLJ5Cu2T84wxrp7dLH3pXkPzz
# JDpl/dSfs5gwJ8nK08WS7Ryp4Fi2gOnJoh74cnqahYyzHoB9tS1bbWOjQNmSTgTN
# Vb16UXRo9udQIp6/Zf7v1yPRKgmnINiVVvueaswQ3mZKLa1PC3lkoqmIyZAnBAK/
# SIG # End signature block
