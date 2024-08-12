<#
.EXTERNALHELP AggregatePFData-help.xml
#>

# Copyright (c) 2009 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

Param(
	[string] $PublicFolder,
	[string] $Server
)

#load hashtable of localized string
Import-LocalizedData -BindingVariable AggregatePFData_LocalizedStrings -FileName AggregatePFData.strings.psd1

$script:SkipCount = 0
$script:SkippedString = ""

# Function that dumps aggregated information from various pf cmdlets
function dump-pf($pf, $TargetServer)
{
	# We don't care about the IPM root
	if ($pf.Identity.ToString() -ne "\")
	{
		$HasReplica = $False

		if ($TargetServer)
		{
			# Caller specified a particular server.
			# Make sure that server's database has a replica of this public folder.
			$Server = $TargetServer
			$dbName = (get-publicfolderdatabase -Server $TargetServer).Name
			$ReplicaOnThisServer = $pf.Replicas | where-object { $_.Name -eq $dbName }
			if ($ReplicaOnThisServer)
			{
				$HasReplica = $True
				$pfStats = get-publicfolderstatistics $pf.Identity -server $Server
				$LastUserAccessTime = $pfStats.LastUserAccessTime
				$LastUserModificationTime = $pfStats.LastUserModificationTime
			}
		}
		else
		{
			if ($pf.Replicas.Count -gt 0)
			{
				# Caller did not specify a particular server.
				# Walk all servers whose databases hold replicas of this pf and get stats from each one.
				# In particular, the LastUserAccessTime and LastUserModifiedTime may be different across replicas.
				$pf.Replicas | foreach-object {
					# The replicas are defined in terms of databases.
					# Map each database to its server.
					$replicaServer = (Get-PublicFolderDatabase $_).Server

					# Now ask that server for its statistics for this pf.
					$pfStatsAll += @(Get-PublicFolderStatistics $pf.Identity -Server $replicaServer)
				}

				if ($pfStatsAll.Count -gt 0)
				{
					# Pick any arbitrary stats. The timestamps should be the only thing that varies.
					$pfStats = $pfStatsAll[0];

					# Pick any arbitrary server.
					$Server = $replicaServer
					$HasReplica = $True

					# Get the most recent LastUserAccessTime
					$pfStatsAllByLastUserAccessTime = @($pfStatsAll | Sort-Object -Descending -Property LastUserAccessTime)
					$LastUserAccessTime = $pfStatsAllByLastUserAccessTime[0].LastUserAccessTime

					# Get the most recent LastUserModificationTime
					$pfStatsAllByLastUserModificationTime = @($pfStatsAll | Sort-Object -Descending -Property LastUserModificationTime)
					$LastUserModificationTime = $pfStatsAllByLastUserModificationTime[0].LastUserModificationTime
				}
			}
		}

		if ($HasReplica)
		{

			# Figure out who's the owner of this pf
			$ipfOwner = 0
			$pfOwnerString = ""
			$pfOwners = Get-PublicFolderClientPermission $pf.Identity -server $Server | where {$_.AccessRights[0].ToString() -eq "Owner"}
			$pfOwners | foreach-object { $pfOwnerString += $_.User; if ($ipfOwner -lt ($pfOwners.Count -1)) { $pfOwnerString += ", " } $ipfOwner = $ipfOwner + 1 }

			# Remove duplicate item types.
			$pfItemTypes = get-publicfolderItemStatistics $pf.Identity -server $Server | sort-object -Property ItemType -Unique

			# Convert the list of ItemTypes into a comma separated string
			$ipfItemType = 0
			$pfItemTypesString = ""
			$pfItemTypes | foreach-object { $pfItemTypesString += $_.ItemType; if ($ipfItemType -lt ($pfItemTypes.Count - 1)) { $pfItemTypesString += ", " } $ipfItemType = $ipfItemType + 1 }

			# Create a result object to return to the pipeline
			$configEntry = new-object System.Management.Automation.PSObject;
			add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:Identity -value:$pf.Identity
			add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:Owner -value:$pfOwnerString
			add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:ItemTypes -value:$pfItemTypesString
			add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:FolderType -value:$pf.FolderType
			add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:MailEnabled -value:$pf.MailEnabled
			add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:HasModerator -value:$pf.HasModerator
			add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:HasRules -value:$pf.HasRules
			add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:ItemCount -value:$pfStats.ItemCount
			add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:TotalItemSize -value:$pfStats.TotalItemSize
			add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:LastUserAccessTime -value:$LastUserAccessTime
			add-member -InputObject:$configEntry -MemberType:NoteProperty -Name:LastUserModificationTime -value:$LastUserModificationTime

			write-output $configEntry
		}
		# If there's no replica, just skip this folder
		# and let the user know we skipped some.
		else
		{

			if ($script:SkipCount -ge 25)
			{
				if ($script:SkipCount -eq 25)
				{
					$script:SkippedString += ", ..."
				}
			}
			else
			{
				if ($script:SkipCount -gt 0)
				{
					$script:SkippedString += ", "
				}

				$script:SkippedString += $pf.Identity.ToString()
			}
			$script:SkipCount = $script:SkipCount +  1
		}
	}
}

####################################################################################################
# Script starts here
####################################################################################################

if ($PublicFolder)
{
	if ($Server)
	{
		$pfs = get-publicfolder $PublicFolder -Server $Server
	}
	else
	{
		$pfs = get-publicfolder $PublicFolder
	}
}
else
{
	if ($Server)
	{
		$pfs = get-publicfolder "\" -Server $Server -Recurse
	}
	else
	{
		# build a list of pfs in the org, use the pf dbs as a starting point and eliminate duplicate pfs
		# do it in 2 steps because of remote powershell pipelining
		$pfdbs = get-publicfolderdatabase
		$pfdbs | foreach-object { $pfs += get-publicfolder "\" -Server $_.Server -Recurse; $pfs = $pfs | Sort-Object -Property ParentPath,Name -Unique  }
	}
}

# now dump the information for each pf
$pfs | foreach-object { dump-pf -pf $_ -TargetServer $Server }

if ($script:SkipCount -gt 0)
{
	Write-Warning $AggregatePFData_LocalizedStrings.res_0000
	Write-Warning $script:SkippedString
}

# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUl/9/sN2VQdynjC1ackeReTAW
# 1M2gghhqMIIE2jCCA8KgAwIBAgITMwAAASIn72vt4vugowAAAAABIjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzQw
# WhcNMjAwMTEwMjEwNzQwWjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046RUFDRS1FMzE2LUM5MUQxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQDRU0uHr30XLzP0oZTW7fCdslb6OXTQeoCa/IpqzTgDcXyf
# EqI0fdvhtqQ84neZE4vwJUAbJ2S+ajJirzzZIEU/JiTZpJgeeAMtN+MuAbzXrdyS
# ohyUDuGkuN+vVSeCnEZkeGcFf/zrNWWXmS7JsVK2BJR8YvXk0sBUbWVpdj0uvz68
# Y+HUyx8AKKE2nHRu54f6fC4eiwP/hs+L7NejJm+sNo7HXV4Y6edQI36FdY0Sotq8
# 7Lh3U96U4O6X9cD0iqKxr4lxYYkh98AzVUjiiSdWUt65DAMbdjBV6cepatwVVoET
# EtNK/f83bMS3sOL00QMWoyQM1F7+fLoz1TF7qlozAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUtlFVlkkUKuXnuF3JZxfDlHs2paYwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAZldsd7vjji5U30Zj
# cKiJvhDtcmx0b4s4s0E7gd8Lp4VnvAQAnpc3SkknUslyknvHGE77OSdxKdrO8qnn
# T0Tymqvf7/Re2xJcRVcM4f8TeE5hCaffCkB7Gtu90R+6+Eb1BnBDYMbj3b42Jq8K
# 42hnDG0ntrgv4/TmyJWIvmGQORWMCWyM/NraY3Ldi7pDpTfx9Z9s4eNE/cxipoST
# XHMIgPgDgbZcuFBANnWwF+/swj69cv87x+Jv/8HM/Naoawrr8+0yDjiJ90OzLGI5
# RScuGfQUlH0ESbzevO/9PFpoUywmNYhHoEPngLJVT2W6y13jFUx3IS9lnR0r1dCh
# mynB8jCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUGgnf30rsduYuF1SRMxUPlWPl1GQw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAGS7OUQOaTTC+rcBWPpittbxsxErPmdqbYGw0by45tSE
# jITl4gB/qz2RSHSJP6wFf/2GMxql1DtLjfanifVpNd5FdvG2LxC1tAj9mCnatwsi
# W85Zfbo89oI3qQ6PT15ExCIJPcGGjGc5ehBpOSMlYMQ0p+DbH45YkvOArPZME8TK
# y8OHYk80Bm5b/sWxv6cDhrs6OoZi1EPEYdoIEk9pGLoIraDefF7Sl3pDetf58Ps1
# HiYdn3BWxMrBK5/YpA5eRuQwt5/TSjKjNPRRmGHMS6qnc8Aenpm8acEIorIbcjUx
# 996b7XjRp5XSPTSErS0JEPrMJr7Uz2lLNtjBLtS2+8ehggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# Iifva+3i+6CjAAAAAAEiMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI2MDBaMCMGCSqGSIb3DQEJ
# BDEWBBS/LgJ44gD57xg3jAiqFDpaI2E+MDANBgkqhkiG9w0BAQUFAASCAQASpfAZ
# 2qO48Mljo9/MwP/2SM+H80St7eCEzhZWaNjDACmecFDgxvxcSZdxTxNklZJUIWEC
# 2imuwlsfzAUlSa2hg8d7cvVX7FH5jxZdHSslQsEVn1la2jk6rhuw5CJj8h5G7hWC
# uwCAKEAbKCGDFEbnlVAeAy0dusNNfVhpM4PX+6W6l3Mkhn2psYy0eqXN9SICQ4mY
# 9PhhEkXLTkBcFLIStgTrqvg0YV0sgaCi/zGMxVemK5xZHz8ukS7Hd1ZsjeL3Z6j9
# mJEOsYNPIXCJL86oFQoazES85OfXLcLnqFZP7bOqT5Syoc0MwMUV/X2MLeuZMjz9
# wRbFQFi4c5scap9f
# SIG # End signature block
