<#
.EXTERNALHELP CheckInvalidRecipients-help.xml
#>

# Copyright (c) Microsoft Corporation. All rights reserved.  
# 
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

# Synopsis: This script is designed to return information on invalid recipient objects and possible attemtpt to fix them.
#
# The script will attempt to fix two classes of errors:
# 1. Primary SMTP Address Problems:	If a recipient has multiple SMTP addresses listed as primary or the primary SMTP is invalid, the script
#					will try to set the WindowsEmailAddress as the primary SMTP address, since that is the address Exchange
#					2003 would have recognized as the primary (although E12 does not).
# 2. Distribution Group Hidden Membership:	If a distribution group has HideDLMembershipEnabled set to true, but ReportToManagerEnabled, 
#					ReportToOriginatorEnabled and/or SendOofMessageToOriginatorEnabled are set to true, then the membership
#					is not actually securely hidden. The script will set ReportToManagerEnabled, ReportToOriginatorEnabled and
#					SendOofMessageToOriginatorEnabled to false to fix the distribution group.
#
# Usage:
#
#    .\CheckInvalidRecipients -help
#
#       Gets help for the script 
#
#    .\CheckInvalidRecipients
#
#       Returns all the invalid Recipients in the Org.
#
#    .\CheckInvalidRecipients -OrganizationalUnit 'Users' -FixErrors
#
#       Fixes all the invalid recipients in the Users container of the local domain.
#

Param(
[string] $OrganizationalUnit,
[string] $ResultSize = "Unlimited",
[string] $Filter,
[string] $DomainController,
[switch] $FixErrors,
[switch] $RemoveInvalidProxies,
[switch] $ShowInvalidProxies,
[switch] $OutputObjects
)

#load hashtable of localized string
Import-LocalizedData -BindingVariable CheckInvalidRecipients_LocalizedStrings -FileName CheckInvalidRecipients.strings.psd1

# Catch any random input and output the help text
if ($args) {
exit
}

############################################################ Function Declarations ################################################################

function HasValidWindowsEmailAddress($obj)
{
	return $obj.WindowsEmailAddress.IsValidAddress
}

function HasInvalidPrimarySmtp($obj)
{
	return !$obj.PrimarySmtpAddress.IsValidAddress
}

function IsValid($obj)
{
	if (!$obj.IsValid)
	{ return $false }
	
	foreach ($address in $obj.EmailAddresses)
	{
		if ($address -is [Microsoft.Exchange.Data.InvalidProxyAddress])
		{ return $false }
	}

	return $true
}

function WriteErrorMessage($str)
{
	Write-host $str -ForegroundColor Red
}

function WriteInformation($str)
{
	Write-host $str -ForegroundColor Yellow
}

function WriteSuccess($str)
{
	Write-host $str -ForegroundColor Green
}

function WriteWarning($str)
{
	$WarningPreference = $Global:WarningPreference
	write-warning $str
}

function PrintValidationError($obj)
{
	foreach($err in $obj.Validate())
	{
		WriteErrorMessage('{0},{1},{2}' -f $obj.Id,$err.PropertyDefinition.Name,$err.Description)
	}
}

function EvaluateErrors($Recipient)
{
	PrintValidationError($Recipient)
	
	$tasknoun = $null

	# We're comparing the RecipientType to the enum value instead of strings, because the strings may be localized and then this comparison would fail
	switch ($Recipient.RecipientType)
	{
	{$_ -eq [Microsoft.Exchange.Data.Directory.Recipient.RecipientType]::UserMailbox}				{ $tasknoun = "Mailbox" }
	{$_ -eq [Microsoft.Exchange.Data.Directory.Recipient.RecipientType]::MailUser}				{ $tasknoun = "Mailuser" }
	{$_ -eq [Microsoft.Exchange.Data.Directory.Recipient.RecipientType]::MailContact}			{ $tasknoun = "Mailcontact" }
	{$_ -eq [Microsoft.Exchange.Data.Directory.Recipient.RecipientType]::MailUniversalDistributionGroup}	{ $tasknoun = "DistributionGroup" }
	{$_ -eq [Microsoft.Exchange.Data.Directory.Recipient.RecipientType]::MailUniversalSecurityGroup}		{ $tasknoun = "DistributionGroup" }
	{$_ -eq [Microsoft.Exchange.Data.Directory.Recipient.RecipientType]::MailNonUniversalGroup}		{ $tasknoun = "DistributionGroup" }
	{$_ -eq [Microsoft.Exchange.Data.Directory.Recipient.RecipientType]::DynamicDistributionGroup}					{ $tasknoun = "DynamicDistributionGroup" }
	}

	if (($tasknoun -ne $null) -AND ($FixErrors -OR $ShowInvalidProxies))
	{
		# Prepare the appropriate get/set tasks that need to run
		$GetRecipientCommand = "get-$tasknoun"
		if (![String]::IsNullOrEmpty($DomainController))
		{ $GetRecipientCommand += " -DomainController $DomainController" }

		$SetRecipientCommand = "set-$tasknoun"
		if (![String]::IsNullOrEmpty($DomainController))
		{ $SetRecipientCommand += " -DomainController $DomainController" }

		# Read the object using the correct Get-Task, so we get all the email properties
		$Recipient = &$GetRecipientCommand $Recipient.Identity

		# Nothing to do if the recipient is completely valid except output it to the pipeline
		if (IsValid($Recipient))
		{
			# Output the object to the pipeline
			if ($OutputObjects)
			{ Write-Output $Recipient }

			return;
		}

		# Collect all the invalid proxy addresses in case we need them later
		$InvalidProxies = @()
		foreach ($Address in $Recipient.EmailAddresses)
		{
			if ($Address -is [Microsoft.Exchange.Data.InvalidProxyAddress])
			{
				$InvalidProxies += $Address
			}
		}

		if ($ShowInvalidProxies -AND ($InvalidProxies.Length -gt 0))
		{
			foreach ($Address in $InvalidProxies)
			{
				WriteErrorMessage('{0},{1},{2}' -f $Recipient.Id,"EmailAddresses",$Address.ParseException.ToString())
			}
		}

		if ($FixErrors)
		{
			$RecipientModified = $false
		
			# Fix the major PrimarySmtpAddress problems
			# If the WindowsEmailAddress is valid, we'll set that as the Primary since Exchange 2003 used that as the Primary
			if ((HasValidWindowsEmailAddress($Recipient)) -AND 
			    (HasInvalidPrimarySmtp($Recipient)))
			{
				$Recipient.PrimarySmtpAddress = $Recipient.WindowsEmailAddress
				WriteInformation($CheckInvalidRecipients_LocalizedStrings.res_0001 -f $Recipient.Identity, $Recipient.WindowsEmailAddress)
				$RecipientModified = $true
			}

			# If the ExternalEmailAddress is missing from the EmailAddresses collection, we should add it back
			if (($null -ne $Recipient.ExternalEmailAddress) -AND
			    !($Recipient.EmailAddresses.Contains($Recipient.ExternalEmailAddress)))
			{
				$Recipient.EmailAddresses.Add($Recipient.ExternalEmailAddress)
				$RecipientModified = $true
			}

			# Remove all the invalid proxy addresses if the user specified the RemoveInvalidProxies flag
			if ($RemoveInvalidProxies -AND ($InvalidProxies.Length -gt 0))
			{
				foreach ($Address in $InvalidProxies)
				{
					# Using this DummyVariable so the script doesn't output the result of the Remove operation
					$DummyVariable = $Recipient.EmailAddresses.Remove($Address)
					WriteInformation($CheckInvalidRecipients_LocalizedStrings.res_0002 -f $Recipient.Identity, $Address)
				}
				$RecipientModified = $true
			}

			# Let's try to save the object back to AD
			if ($RecipientModified)
			{
				$numErrors = $error.Count
				&$SetRecipientCommand -Instance $Recipient
				if ($error.Count -eq $numErrors)
				{
					WriteSuccess($CheckInvalidRecipients_LocalizedStrings.res_0003 -f $Recipient.Identity)
				}
				else
				{
					WriteErrorMessage($CheckInvalidRecipients_LocalizedStrings.res_0004 -f $Recipient.Identity)
				}
			}

			# Re-read the recipient if we modified it in any way and we want to output it to the pipeline
			if ($OutputObjects)
			{ $Recipient = &$GetRecipientCommand $Recipient.Identity }

		} # if ($FixErrors)
	} # if (($tasknoun -ne $null) -AND ($FixErrors -OR $ShowInvalidProxies))

	# Output the object to the pipeline
	if ($OutputObjects)
	{ Write-Output $Recipient }

} # EvaluateErrors

############################################################ Function Declarations End ############################################################

############################################################ Main Script Block ####################################################################

#Ignore Warnings output by the task
$WarningPreference = 'SilentlyContinue'

if ($RemoveInvalidProxies -AND !$FixErrors)
{ WriteWarning($CheckInvalidRecipients_LocalizedStrings.res_0005) }

# Check if we have any pipeline input
# If yes, MoveNext will return true and we won't run our get tasks
if ($input.MoveNext())
{
	# Reset the enumerator so we can look at the first object again
	$input.Reset()

	if ($ResultSize -NE "Unlimited")
	{ WriteWarning($CheckInvalidRecipients_LocalizedStrings.res_0006) }
	if (![String]::IsNullOrEmpty($OrganizationalUnit))
	{ WriteWarning($CheckInvalidRecipients_LocalizedStrings.res_0007) }
	if (![String]::IsNullOrEmpty($Filter))
	{ WriteWarning($CheckInvalidRecipients_LocalizedStrings.res_0008) }

	foreach ($Recipient in $input)
	{
		# skip over inputs that we can't handle
		if ($Recipient -eq $null -OR 
		    $Recipient.RecipientType -eq $null -OR
		    $Recipient -isnot [Microsoft.Exchange.Data.Directory.ADObject])
		{ continue; }
	
		EvaluateErrors($Recipient)
	}
}
else
{
	$cmdlets =
		@("get-User",
		"get-Contact",
		"get-Group",
		"get-DynamicDistributionGroup")

	foreach ($task in $cmdlets)
	{
		$command = "$task -ResultSize $ResultSize"
		if (![String]::IsNullOrEmpty($OrganizationalUnit))
		{ $command += " -OrganizationalUnit $OrganizationalUnit" }
		if (![String]::IsNullOrEmpty($DomainController))
		{ $command += " -DomainController $DomainController" }
		if (![String]::IsNullOrEmpty($Filter))
		{ $command += " -Filter $Filter" }

		invoke-expression $command | foreach { EvaluateErrors($_) }
	}
}

############################################################ Main Script Block END ################################################################

# SIG # Begin signature block
# MIIaegYJKoZIhvcNAQcCoIIaazCCGmcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUdLu/uLnYGkP9ikFUeohJAUv5
# m7GgghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggTDMIIDq6ADAgEC
# AhMzAAAAKzkySMGyyUjzAAAAAAArMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQTAeFw0xMjA5MDQyMTEyMzRaFw0xMzEyMDQyMTEyMzRaMIGz
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRN
# T1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046QzBGNC0zMDg2LURFRjgxJTAj
# BgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCmtjAOA2WuUFqGa4WfSKEeycDuXkkHheBwlny+
# uV9iXwYm04s5uxgipS6SrdhLiDoar5uDrsheOYzCMnsWeO03ODrxYvtoggJo7Ou7
# QIqx/qEsNmJgcDlgYg77xhg4b7CS1kANgKYNeIs2a4aKJhcY/7DrTbq7KRPmXEiO
# cEY2Jv40Nas04ffa2FzqmX0xt00fV+t81pUNZgweDjIXPizVgKHO6/eYkQLcwV/9
# OID4OX9dZMo3XDtRW12FX84eHPs0vl/lKFVwVJy47HwAVUZbKJgoVkzh8boJGZaB
# SCowyPczIGznacOz1MNOzzAeN9SYUtSpI0WyrlxBSU+0YmiTAgMBAAGjggEJMIIB
# BTAdBgNVHQ4EFgQUpRgzUz+VYKFDFu+Oxq/SK7qeWNAwHwYDVR0jBBgwFoAUIzT4
# 2VJGcArtQPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1w
# UENBLmNybDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# dDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAfsywe+Uv
# vudWtc9z26pS0RY5xrTN+tf+HmW150jzm0aIBWZqJoZe/odY3MZjjjiA9AhGfCtz
# sQ6/QarLx6qUpDfwZDnhxdX5zgfOq+Ql8Gmu1Ebi/mYyPNeXxTIh+u4aJaBeDEIs
# ETM6goP97R2zvs6RpJElcbmrcrCer+TPAGKJcKm4SlCM7i8iZKWo5k1rlSwceeyn
# ozHakGCQpG7+kwINPywkDcZqJoFRg0oQu3VjRKppCMYD6+LPC+1WOuzvcqcKDPQA
# 0yK4ryJys+fEnAsooIDK4+HXOWYw50YXGOf6gvpZC3q8qA3+HP8Di2OyTRICI08t
# s4WEO+KhR+jPFTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEF
# BQAwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jv
# c29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9y
# aXR5MB4XDTEwMDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENv
# ZGUgU2lnbmluZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCy
# cllcGTBkvx2aYCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPV
# cgDbNVcKicquIEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlc
# RdyvrT3gKGiXGqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZ
# C/6SdCnidi9U3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgG
# hVxOVoIoKgUyt0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdc
# pReejcsRj1Y8wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBTLEejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYB
# BAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy8
# 2C0wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBW
# J5flJRP8KuEKU5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNy
# b3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3Js
# MFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcN
# AQEFBQADggIBAFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGj
# I8x8UJiAIV2sPS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbN
# LeNK0rxw56gNogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y
# 4k74jKHK6BOlkU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnp
# o1hW3ZsCRUQvX/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6
# H0q70eFW6NB4lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20O
# E049fClInHLR82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8
# Z4L5UrKNMxZlHg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9G
# uwdgR2VgQE6wQuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXri
# lUEnacOTj5XJjdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEv
# mtzjcT3XAH5iR9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCC
# A++gAwIBAgIKYRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPy
# LGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRN
# aWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1
# MzA5WhcNMjEwNDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZ
# USNQrc7dGE4kD+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cO
# BJjwicwfyzMkh53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn
# 1yjcRlOwhtDlKEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3
# U21StEWQn0gASkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG
# 7bfeI0a7xC1Un68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMB
# AAGjggGrMIIBpzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A
# +3b7syuwwzWzDzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1Ud
# IwSBkDCBjYAUDqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/Is
# ZAEZFgNjb20xGTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1p
# Y3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0
# BxMuZTBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUH
# AQEESDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# L2NlcnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcD
# CDANBgkqhkiG9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwT
# q86+e4+4LtQSooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+
# jwoFyI1I4vBTFd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwY
# Tp2OawpylbihOZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPf
# wgphjvDXuBfrTot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5
# ZlizLS/n+YWGzFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8cs
# u89Ds+X57H2146SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUw
# ZuhCEl4ayJ4iIdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHu
# diG/m4LBJ1S2sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9
# La9Zj7jkIeW1sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g
# 74TKIdbrHk/Jmu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUW
# a8kTo/0xggS1MIIEsQIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIHOMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBR5zaYvZtDJHvkh2EOpzsnA2roTVTBuBgorBgEEAYI3AgEMMWAw
# XqA2gDQAQwBoAGUAYwBrAEkAbgB2AGEAbABpAGQAUgBlAGMAaQBwAGkAZQBuAHQA
# cwAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAw
# DQYJKoZIhvcNAQEBBQAEggEAFXpbg89tn1AZfEUrDutJV98gQUg3g0Xf8nSSyn5Z
# GWmHBDBPPqqDI20h0ZIrFklMvexAc+V9TseR/+Fu+NWBgk52eoPDVcVPgPLwP2oD
# CWsFvYn3zNOqZfOdvJwt1wDKtKKnPPaQWhP3Yc8q/tJTIsXmcoqwbYzXXpNeOHPW
# 968P2GJMM2yg5Qtacj1FOLl54qOH+0OsYOb4CWyw3Zu7aykNttnNhzcxS59vl6hK
# bLDlGgK1/wCq2SJLjp5ZyOdNGN6FgDU4S7rjkj8dkUqojIjOdaqSdan8hPzrl3ex
# /+JQa/739sPaCbdMyOTCzc0zk8ULVwL8lCxz+70pqCFjlqGCAigwggIkBgkqhkiG
# 9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMA
# AAArOTJIwbLJSPMAAAAAACswCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkq
# hkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTEzMDIwNTA2MzcyMVowIwYJKoZIhvcN
# AQkEMRYEFLh8QK6h1x5sFzXYLHvG34t/HtD8MA0GCSqGSIb3DQEBBQUABIIBAHQS
# Z+tVXhurxvDKKqsFz3dieXos3+vVBhbnfSPL2uVKIwgIFUVVE3bM257JtYMaYA/+
# kRGgya7QMJGHnwBUewL9GR/dM/Jm0qyQY3/LJ1FQp5RxyJxfAyvZglRSb3o3C7MJ
# uAp+zqh5RAZvGCf/SS/UyJTg2akt1C27PZE7wMYVgcsRj6r6A4cxg+NNCp6lyRFU
# LbAH6bjmSgdJhmvgtMFJjxFR1NwmHp/9LZyII1RcVhWuTRu9F17FYZhvTx0Orgyg
# b8iBprCmdWNgmzPjck5nRK5y2Fa3F9wM/0NkVx1atzrb1M5L97HjhAEIKBXklO7Y
# iBIP8tKt050by5nz4aQ=
# SIG # End signature block
