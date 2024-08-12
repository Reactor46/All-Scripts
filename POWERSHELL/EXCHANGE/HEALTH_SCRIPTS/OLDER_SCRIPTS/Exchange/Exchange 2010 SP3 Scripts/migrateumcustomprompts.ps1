# Copyright (c) Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
# MigrateUMCustomPrompts.ps1
#
# This script migrates a copy of all Exchange 2007 Unified Messaging custom prompts 
# to Exchange 2010. This includes prompts in all Exchange 2007 UM Dial Plans and
# UM Automated Attendants.
#
# Usage:
#     .\MigrateUMCustomPrompts.ps1
#

#
# Helper functions
#

function IsPromptFileSet($promptFileName)
{
	if ([string]::IsNullOrEmpty($promptFileName))
	{
		return $false
	}
	else
	{
		return $true
	}
}

function IsPromptMigrationRequired($dp)
{
	# Prompt migration is required only if a legacy prompt publishing 
	# point is set on the dial plan
	if ([string]::IsNullOrEmpty($dp.LegacyPromptPublishingPoint))
	{
		return $false
	}
	else
	{
		return $true
	}
}

function ImportDialPlanPrompt($dp, $promptFileName)
{
	$filePath = $dp.LegacyPromptPublishingPoint.ToString() + "\" + $dp.Guid.ToString() + "\" + $promptFileName

	if ([System.IO.File]::Exists($filePath))
	{
		"..Fetch custom prompt from `'$filePath`'"
		[byte[]]$prompt = Get-Content -Path $filePath -Encoding byte -ReadCount 0 -ErrorAction Stop

		"..Import custom prompt to Exchange 2010"
		Import-UMPrompt -UMDialPlan $dp.Name -PromptFileName $promptFileName -PromptFileData $prompt -ErrorAction Stop
	}
	else
	{
		# Write message as a warning and as regular output so that it shows up in the log.
		$msg = "`'$filePath`' was not found. The file may not exist or you may not have the required permissions."
		$msg
		Write-Warning $msg
		$script:anyWarnings = $true
	}
}

function ImportAutoAttendantPrompt($dp, $aa, $promptFileName)
{
	$filePath = $dp.LegacyPromptPublishingPoint + "\" + $dp.Guid.ToString() + "\" + $aa.Guid.toString() + "\" + $promptFileName

	if ([System.IO.File]::Exists($filePath))
	{
		"..Fetch custom prompt from `'$filePath`'"
		[byte[]]$prompt = Get-Content -Path $filePath -Encoding byte -ReadCount 0 -ErrorAction Stop

		"..Import custom prompt to Exchange 2010"
		Import-UMPrompt -UMAutoAttendant $aa.Name -PromptFileName $promptFileName -PromptFileData $prompt -ErrorAction Stop
	}
	else
	{
		# Write message as a warning and as regular output so that it shows up in the log.
		$msg = "`'$filePath`' was not found. The file may not exist or you may not have the required permissions."
		$msg
		Write-Warning $msg
		$script:anyWarnings = $true
	}
}

#
# Main script starts here
#

# Stop script execution on error
$ErrorActionPreference = "Stop"

# Script variable to determine whether any warnings were raised
$script:anyWarnings = $false

# Retrieve all dial plans and migrate prompts if required
$dialPlans = @(Get-UMDialPlan -ErrorAction Stop)
foreach ($dp in $dialPlans)
{
	$migratePrompts = IsPromptMigrationRequired($dp)
	if ($migratePrompts -eq $false)
	{
		"`n`n UM Dial Plan `'$dp`' does not have a legacy prompt publishing point..."
		continue
	}

	"`n`n------ Processing UM Dial Plan `'$dp`'... ------"

	if (IsPromptFileSet($dp.WelcomeGreetingFilename))
	{
		".Custom welcome greeting"
		ImportDialPlanPrompt $dp $dp.WelcomeGreetingFilename
	}

	if (IsPromptFileSet($dp.InfoAnnouncementFilename))
	{
		".Informational announcement"
		ImportDialPlanPrompt $dp $dp.InfoAnnouncementFilename
	}
}

# Retrieve all auto attendants and migrate prompts if required
$autoAttendants = @(Get-UMAutoAttendant -ErrorAction Stop)
foreach ($aa in $autoAttendants)
{
	$dp = Get-UMDialPlan $aa.UMDialPlan -ErrorAction Stop
	$migratePrompts = IsPromptMigrationRequired($dp)
	if ($migratePrompts -eq $false)
	{
		"`n`n UM Auto Attendant `'$aa`' belongs to UM Dial Plan `'$dp`' which does not have a legacy prompt publishing point..."
		continue
	}
	
	"`n`n------ Processing UM Auto Attendant `'$aa`'... ------"
	
	if (IsPromptFileSet($aa.BusinessHoursWelcomeGreetingFilename))
	{
		".Custom business hours welcome greeting prompt"
		ImportAutoAttendantPrompt $dp $aa  $aa.BusinessHoursWelcomeGreetingFilename
	}

	if (IsPromptFileSet($aa.BusinessHoursMainMenuCustomPromptFilename))
	{
		".Custom business hours main menu prompt"
		ImportAutoAttendantPrompt $dp $aa $aa.BusinessHoursMainMenuCustomPromptFilename
	}

	if (IsPromptFileSet($aa.InfoAnnouncementFilename))
	{
		".Informational Announcement"
		ImportAutoAttendantPrompt $dp $aa $aa.InfoAnnouncementFilename
	}

	if (IsPromptFileSet($aa.AfterHoursWelcomeGreetingFilename))
	{
		".Custom after hours welcome greeting prompt"
		ImportAutoAttendantPrompt $dp $aa $aa.AfterHoursWelcomeGreetingFilename
	}

	if (IsPromptFileSet($aa.AfterHoursMainMenuCustomPromptFilename))
	{
		".Custom after hours main menu prompt"
		ImportAutoAttendantPrompt $dp $aa $aa.AfterHoursMainMenuCustomPromptFilename
	}

	$holidaySchedule = @($aa.HolidaySchedule)
	foreach ($holiday in $holidaySchedule)
	{
		".Holiday: `'" + $holiday.greeting + "`'"
		ImportAutoAttendantPrompt $dp  $aa $holiday.greeting
	}
	
	$businessHoursKeyMapping = @($aa.BusinessHoursKeyMapping)
	foreach ($keyMapping in $businessHoursKeyMapping)
	{
		if (IsPromptFileSet($keyMapping.PromptFileName))
		{
			$promptFileName = $keyMapping.PromptFileName
			".BusinessHoursKeyMapping: Key " + $keyMapping.Key.ToString() +": `'" + $promptFileName + "`'"
			ImportAutoAttendantPrompt $dp $aa $promptFileName
		}
	}
	
	$afterHoursKeyMapping = @($aa.AfterHoursKeyMapping)
	foreach ($keyMapping in $afterHoursKeyMapping)
	{
		if (IsPromptFileSet($keyMapping.PromptFileName))
		{
			$promptFileName = $keyMapping.PromptFileName
			".AfterHoursKeyMapping: Key " + $keyMapping.Key.ToString() +": `'" + $promptFileName + "`'"
			ImportAutoAttendantPrompt $dp $aa $promptFileName
		}
	}
}

if ($script:anyWarnings -eq $true)
{
	Write-Warning "Some prompt files were not found in the prompt publishing point and could not be imported. Please review script output for details."
}
else
{
	"`n`nFinished!"
}
# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUa0YL6Fb9/pFfVmOVvcy9Wmk0
# HcygghhqMIIE2jCCA8KgAwIBAgITMwAAARzbbpm3tnP6bwAAAAABHDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzM1
# WhcNMjAwMTEwMjEwNzM1WjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046RDJDRC1FMzEwLTRBRjExJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQCxqRuPkgvAvJMVHxyEsWMAs/pxAn3vnvfWrFqQj2NkG9kP
# E3XXn9Xn7n7WsHbuuVdpi4nSyPfLTriA2kzbF+eco/ZTVRbanYk8BXwZGgUzRgF4
# LxQq4INdpNmH2zBti8HK7xURC8HoBB82c5VnZp1AZvgnWRs+6wbzXnauqbwoGuTJ
# XPzaPXivUjL2W+W9G9NMJ5nrmkcNcmq/ncqA88qrofMBqly6y+SL1EdCR0oVYl1A
# ZOgf+ALrh/TMeA1Bld+EFzJa/rEo1QB3IPcwm3xQfW26SYOyQFPIfLjXkBs+VYrc
# S27bByATdjsOJ06krz5tc2fKLv+ao5r1sOIvFDcFAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUb8nAx97t5y1LdYL20QwUPKqBH8UwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAWVKU4uhqdIGVX+vj
# MkduTPqjk59ZxNeOrJX/O7MP5OkObcq6T+vqTyjmeTsiNoO0btyofj9bUJUAic8z
# 10V/rwlvvsYUyzlnTos7+76NU86PoQuMGTLuPfmEAQD4rpUs1kyJchz2m0q7/AbI
# usbsTTLzJ8TW7vyEluJG9LhLAxvAz7dvWdcWQBmh52egoL84XvUq4g0lFNqkiSIV
# 7z7IFsXbvXzhS2NnOLIdpHjGfxhIvRCTFNKCxflV+O8/AqERd6txTeBFpWPRvN0U
# S+GOJvA77FxAvGH2vaH3zQ3WeQxVBAJ6LrUCiKkKm+gJFwE/2ftF5zEMuZS9Zg/F
# EnmzLDCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUqo5/fp2nM/TrL0Dg81D/44izeDsw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBACr6TykRxi8v1VRyVQqQAH5lMty3aoCvtehxdj1YOWjJ
# sL0I0dUVZ5TPc7POy1noDhYORRseG2ZEMlwUsox3rp25X4Yt0wOwNjiKmdePKePj
# 5sJ/7Ax0+LqVFSPeV8II3CYRa+1mW35f3pIZ/WfiocBi9bm7d3RqWrNe24eODZy9
# 59p6tcSFktlPiwfJM5xtEBBxdal+2D2/BGu0Q5jD4GdljhXHJBS5MWSu2LQ3n1oJ
# 5RCUWgBmuiV8MPCcdLZSoHybDN4gVGwaWU5HBETF/a/IjIE+hoOE3PPm0ZtEsOgC
# rlHIMFNE02LcUWHSP1CgibRKBO3NV0X0vV4oF8+axwqhggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# HNtumbe2c/pvAAAAAAEcMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI2MDFaMCMGCSqGSIb3DQEJ
# BDEWBBTl0kRDjmkmt2WaT+JUS1vOGbWLaDANBgkqhkiG9w0BAQUFAASCAQACJYxw
# 1dULnA/DEU4nz20267eFNUUVj7sX3Nr2QpprUWgpy8D91oqKnVHfJkz7dky5H2zt
# +wwfCjGYhwgbbm6SbLMJhPEaFlYjwCEB/VtSsQq1si8uWnhVRAcx7zcKp52Tx0pt
# QoJLrdDbbvLU60AwERYs5pQAcksgHjYOlGSH7kCx7ZOgmXEkyXsel9NlkYff/eMJ
# 1zeLOWj9WQmrPAR+NWN/IS+eLjK9iM9xXClXDTJDcw6ZBmmSpdAEFz/RYxUITsD2
# n7zvHYVTtmpJm2IA2rp8rYEN9KaqgnehBpvylBuCkLgynWx780Ffru6pxLNwYLzT
# g9/ZhSvzSh2Tr5mW
# SIG # End signature block
