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
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUdLu/uLnYGkP9ikFUeohJAUv5
# m7GgghhqMIIE2jCCA8KgAwIBAgITMwAAAR+XYwozuYPXKwAAAAABHzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzM3
# WhcNMjAwMTEwMjEwNzM3WjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046NDlCQy1FMzdBLTIzM0MxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQCppklVnT29zi13dODY0ejMsdoe7n2iCvC6QdH5FJkRYfy+
# cXoHBmpDgDF/65Kt9GMmu/K8HKAzjKHeG18rgRXQagLwIIH5yCRbXGwOfuHIu1dC
# 26o/CT22+YlRvBJwH36WVjML8BLNDT3Fr+yhU4ZM7Hbegql4r5kSgsrrjyx5bJY5
# r2N0G7RDnbhRd79iqXbvDnvkatjB5xgluzfQEAPbJjXjmRb5685DEEZg1qFsQJer
# XuBA+ZVevuCX0DuDj8UmhHGC5Y32sulFTn283R6LU+8+AALtbHOOIHV7QHNYV8mN
# jxHuKLvE9tNEGIpbG2WF2yQkSGe3sRbGQmaILWeHAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUuPNVyPmK8/JJioMtQFlTUeF3IOgwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAmAYfr1fEosYv9VTf
# 0Msya6aFm0Id6Zq1O5jNy74ByTh7EEac/l/4e3DOyrczHS6zwvMKYzLtmifeGZvD
# 70qbbUfF+yjpzpyu00uuzZ1HNOpktp5/dJXkzz0NyVnEeFGOXLpNyZNIA9dKGDwN
# XbsEUukTX9lJFx5RcBhE8AOl22IHSgJ6NYf4DpATCjSJbC9IrKYGBchHobCLZHEt
# cLBjxXiWJRG2YY+LBAVW95gwNdPmLCKrob7SdNLK1VnM35Q2VgNF7YfDc5nw4E7C
# 4VaZvlyuDET6fYycIVPx5GsLhx3it4a+WKcBwarK7inH9skUArxMZrpWmjuQ/o4b
# GprEnjCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
# AQELBQAwfjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYG
# A1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xOTA1MDIy
# MTM3NDZaFw0yMDA1MDIyMTM3NDZaMHQxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xHjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJVaxoZpRx00HvFVw2Z19mJUGFgU
# ZyfwoyrGA0i85lY0f0lhAu6EeGYnlFYhLLWh7LfNO7GotuQcB2Zt5Tw0Uyjj0+/v
# UyAhL0gb8S2rA4fu6lqf6Uiro05zDl87o6z7XZHRDbwzMaf7fLsXaYoOeilW7SwS
# 5/LjneDHPXozxsDDj5Be6/v59H1bNEnYKlTrbBApiIVAx97DpWHl+4+heWg3eTr5
# CXPvOBxPhhGbHPHuMxWk/+68rqxlwHFDdaAH9aTJceDFpjX0gDMurZCI+JfZivKJ
# HkSxgGrfkE/tTXkOVm2lKzbAhhOSQMHGE8kgMmCjBm7kbKEd2quy3c6ORJECAwEA
# AaOCAX4wggF6MB8GA1UdJQQYMBYGCisGAQQBgjdMCAEGCCsGAQUFBwMDMB0GA1Ud
# DgQWBBRXghquSrnt6xqC7oVQFvbvRmKNzzBQBgNVHREESTBHpEUwQzEpMCcGA1UE
# CxMgTWljcm9zb2Z0IE9wZXJhdGlvbnMgUHVlcnRvIFJpY28xFjAUBgNVBAUTDTIz
# MDAxMis0NTQxMzUwHwYDVR0jBBgwFoAUSG5k5VAF04KqFzc3IrVtqMp1ApUwVAYD
# VR0fBE0wSzBJoEegRYZDaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# cmwvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNybDBhBggrBgEFBQcBAQRV
# MFMwUQYIKwYBBQUHMAKGRWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y2VydHMvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNydDAMBgNVHRMBAf8E
# AjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBaD4CtLgCersquiCyUhCegwdJdQ+v9Go4i
# Elf7fY5u5jcwW92VESVtKxInGtHL84IJl1Kx75/YCpD4X/ZpjAEOZRBt4wHyfSlg
# tmc4+J+p7vxEEfZ9Vmy9fHJ+LNse5tZahR81b8UmVmUtfAmYXcGgvwTanT0reFqD
# DP+i1wq1DX5Dj4No5hdaV6omslSycez1SItytUXSV4v9DVXluyGhvY5OVmrSrNJ2
# swMtZ2HKtQ7Gdn6iNntR1NjhWcK6iBtn1mz2zIluDtlRL1JWBiSjBGxa/mNXiVup
# MP60bgXOE7BxFDB1voDzOnY2d36ztV0K5gWwaAjjW5wPyjFV9wAyMX1hfk3aziaW
# 2SqdR7f+G1WufEooMDBJiWJq7HYvuArD5sPWQRn/mjMtGcneOMOSiZOs9y2iRj8p
# pnWq5vQ1SeY4of7fFQr+mVYkrwE5Bi5TuApgftjL1ZIo2U/ukqPqLjXv7c1r9+si
# eOcGQpEIn95hO8Ef6zmC57Ol9Ba1Ths2j+PxDDa+lND3Dt+WEfvxGbB3fX35hOaG
# /tNzENtaXK15qPhErbCTeljWhLPYk8Tk8242Z30aZ/qh49mDLsiL0ksurxKdQtXt
# v4g/RRdFj2r4Z1GMzYARfqaxm+88IigbRpgdC73BmwoQraOq9aLz/F1555Ij0U3o
# rXDihVAzgzCCBgcwggPvoAMCAQICCmEWaDQAAAAAABwwDQYJKoZIhvcNAQEFBQAw
# XzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29m
# dDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# MB4XDTA3MDQwMzEyNTMwOVoXDTIxMDQwMzEzMDMwOVowdzELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAn6Fssd/b
# SJIqfGsuGeG94uPFmVEjUK3O3RhOJA/u0afRTK10MCAR6wfVVJUVSZQbQpKumFww
# JtoAa+h7veyJBw/3DgSY8InMH8szJIed8vRnHCz8e+eIHernTqOhwSNTyo36Rc8J
# 0F6v0LBCBKL5pmyTZ9co3EZTsIbQ5ShGLieshk9VUgzkAyz7apCQMG6H81kwnfp+
# 1pez6CGXfvjSE/MIt1NtUrRFkJ9IAEpHZhEnKWaol+TTBoFKovmEpxFHFAmCn4Tt
# VXj+AZodUAiFABAwRu233iNGu8QtVJ+vHnhBMXfMm987g5OhYQK1HQ2x/PebsgHO
# IktU//kFw8IgCwIDAQABo4IBqzCCAacwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4E
# FgQUIzT42VJGcArtQPt2+7MrsMM1sw8wCwYDVR0PBAQDAgGGMBAGCSsGAQQBgjcV
# AQQDAgEAMIGYBgNVHSMEgZAwgY2AFA6sgmBAVieX5SUT/CrhClOVWeSkoWOkYTBf
# MRMwEQYKCZImiZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0
# MS0wKwYDVQQDEyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHmC
# EHmtFqFKoKWtTHNY9AcTLmUwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQu
# Y3JsMFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwEwYDVR0l
# BAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcNAQEFBQADggIBABCXisNcA0Q23em0rXfb
# znlRTQGxLnRxW20ME6vOvnuPuC7UEqKMbWK4VwLLTiATUJndekDiV7uvWJoc4R0B
# hqy7ePKL0Ow7Ae7ivo8KBciNSOLwUxXdT6uS5OeNatWAweaU8gYvhQPpkSokInD7
# 9vzkeJkuDfcH4nC8GE6djmsKcpW4oTmcZy3FUQ7qYlw/FpiLID/iBxoy+cwxSnYx
# PStyC8jqcD3/hQoT38IKYY7w17gX606Lf8U1K16jv+u8fQtCe9RTciHuMMq7eGVc
# WwEXChQO0toUmPU8uWZYsy0v5/mFhsxRVuidcJRsrDlM1PZ5v6oYemIp76KbKTQG
# dxpiyT0ebR+C8AvHLLvPQ7Pl+ex9teOkqHQ1uE7FcSMSJnYLPFKMcVpGQxS8s7Ow
# TWfIn0L/gHkhgJ4VMGboQhJeGsieIiHQQ+kr6bv0SMws1NgygEwmKkgkX1rqVu+m
# 3pmdyjpvvYEndAYR7nYhv5uCwSdUtrFqPYmhdmG0bqETpr+qR/ASb/2KMmyy/t9R
# yIwjyWa9nR2HEmQCPS2vWY+45CHltbDKY7R4VAXUQS5QrJSwpXirs6CWdRrZkocT
# dSIvMqgIbqBbjCW/oO+EyiHW6x5PyZruSeD3AWVviQt9yGnI5m7qp5fOMSn/DsVb
# XNhNG6HY+i+ePy5VFmvJE6P9MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCBKgwggSkAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAFRno2PQHGjDkEAAAAAAVEwCQYFKw4DAhoFAKCB
# vDAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUec2mL2bQyR75IdhDqc7JwNq6E1Uw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAB4T6EpkFN7CzssOhDs/MIc1xEhEmcislp4ZYKWzfb3b
# yh+WguxOG9OBYjzbQd5/UdEAPWrqVyYx91hLwYG9qV+UyrpVUkvy8tgZRrkgkWc6
# QGUOtJfZ1Uh4rXpLAsouqTxvFR0yVQLCzXMR9levTFek5Oo06IwrXGbOZ0g4ayE5
# HPesovX6axC3y/+JjhWqX8OeYRycvpL7GJFxU7pBQP3rhdkU6BGF96uUwct+KRkl
# dcSnU3mbexTm4ZKfP3hrkGyWngntrU8p4YItdkW7sXV3+o8l+uv6Wfx+zK3DWu9N
# y7UxW3ZkWK53KbgXUXopG5172L2cCK3rk1J2F1RrEkuhggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# H5djCjO5g9crAAAAAAEfMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI1NTdaMCMGCSqGSIb3DQEJ
# BDEWBBTBK1/aMowVGG97AI2XW0XCwY94ijANBgkqhkiG9w0BAQUFAASCAQAPDZ9+
# I4ZhxrN3hDd90tqleam8rftrK4tkm862ZHL0MIf7h85C/mxQJM3yKqklzjmPtJsW
# PupQR7tZJLoZc1VCx3S3HczJsiFnc1DGrPkQ1m6Fam14cNqPd/jMb/Lm1o+lWezI
# H+7BRnneRdmoXsz6B8qUx5vQ5F+nj8cvtU+SKi+y/ry0Hb/q/qeWVWNoIns2Ermc
# WtzL4dcBVAyyh8EpUKk7+1ig6rpRI42F8zlrEJAS2ZI2dNgHwrA7AE2vxldRAC6a
# 34ur2xQLajimLpR2yjO3Me4oAeH8kYJC/v3YvqU+Y9oaMTriiHpiBlnmukXhHuB9
# J2nJ7j5yCZHBUmQK
# SIG # End signature block
