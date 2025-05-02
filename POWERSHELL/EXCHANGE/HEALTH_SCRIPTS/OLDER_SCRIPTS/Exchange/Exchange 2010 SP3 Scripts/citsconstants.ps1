# Copyright (c) 2009 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
# This file contains global constants used by CI Troubleshooter library.
#

# Global constants for performance counter names
#
$searchIndicesCounterObjectName = "MSExchange Search Indices"
$aolniCounterName = "age of the last notification indexed"
$tslniCounterName = "time since last notification was indexed"
$noirtCounterName = "number of items in retry table"
$nomltcCounterName = "number of mailboxes left to crawl"
$noirfrtCounterName = "number of items removed from retry table"

# counter array used for getting performance counter values
#
$counters = (
                "$searchIndicesCounterObjectName(*)\$aolniCounterName",
                "$searchIndicesCounterObjectName(*)\$tslniCounterName",
                "$searchIndicesCounterObjectName(*)\$noirtCounterName",
                "$searchIndicesCounterObjectName(*)\$nomltcCounterName",
                "$searchIndicesCounterObjectName(*)\$noirfrtCounterName"
            )

# Hash table used to store performance counter values for catalogs
#
$counterHashTable = @{}

# Hash table used to store the corresponding mailbox database objects values of catalogs
#
$mailboxDatabaseHashTable = @{}

$exchangeRegKey = "SOFTWARE\Microsoft\ExchangeServer\v14"

# Name of registry hive where exchange setup information is stored
#
$exchangeSetupRegKey= "HKLM:$exchangeRegKey\Setup"

# Name of registry hive where TS diagnostic information is stored
#
$troubleshooterRegKey= "$exchangeRegKey\CITroubleshooter\"

# Name of registry hive where contentIndex state is stored
#
$contentIndexKeyPath = "$exchangeRegKey\ContentIndex"

# Name of registry hive where catalog health values are stored
#
$copyHealthKeyPath = "$contentIndexKeyPath\CatalogHealth\"

# Last known Filters key Name
#
$lastKnownFiltersKeyName = "$contentIndexKeyPath\LastKnownFilters"

# Need Reindex Databases Key Name
#
$needReindexDatabasesKeyName = "$contentIndexKeyPath\NeedReindexDatabases"

# Name of the registry hive where the MSSearch IFilters are defined
#
$MsSearchFilterPath="$exchangeRegKey\MSSearch\Filters"

# Value to determine if we should automatically disable bad IFilters
# If the value is set 0, the troubleshooter will skip the logic of disabling IFilters
$disableBadIFilters = 1

# String used to report catalog corruption or other reset
# conditions in catalog health registry hive
#
$corruptionIndicator = "CatalogNeedReset"

# Stall threshold is 1 hour
#
$stallThresholdInSeconds = 1 * 60 * 60

# Extended period Stall threshold is 5 days
#
$extendedStallThresholdInSeconds = 1 * 60 * 60 * 24 * 5

# Stale health threshold is 30 minutes
#
$staleThresholdInSeconds = 30 * 60

# Backlog health threshold is 48 hours
#
$backlogThresholdInSeconds = 48 * 60 * 60

# Retry items threshold
#
$retryItemsThreshold = 10000

# The minimun number of times the number of items in the retry table is found to be above the set threshold before 
# checking if the retry feeder is stalled.
#
$minRetryTableIssueThreshold = 2

# The minimun number of items in the retry table that should be drained over the past hour for a catalog before assuming
# the retry feeder for that catalog is stalled. 
#
$retryTableDrainThreshold = 500

# MSFTESQL service name
#
$msftesqlServiceName = "msftesql-exchange"

# MSFTESQL process name
#
$msftesqlProcessName = "msftesql"

# MSFTEFD process name
#
$msftefdProcessName = "msftefd"

# ExSearch service name
#
$exsearchServiceName = "MSExchangeSearch"

# ExSearch process name
#
$exsearchProcessName = "Microsoft.Exchange.Search.ExSearch"

# ExSearch event source name
# 
$exsearchEventSource = "MSExchange Search Indexer"

# default timeout for restarting services
#
$defaultRestartTimeout = New-TimeSpan -Minutes 15

# Timespan to wait before taking the next crash dump
#
$minTimeBetweenCrashDumps = New-TimeSpan -Days 2

# name used for application log
#
$appLogName = "Application"

# Event log source name for application log
#
$appLogSourceName = "CI Troubleshooter"

# name used for crimson log
#
$crimsonLogName = "Microsoft-Exchange-Troubleshooters/Operational"

# Event log source name for crimson log
#
$crimsonLogSourceName = "Content Index"

# Indicates the number of minutes we would go back from current time
# to start checking events for bad disk blocks/msftesql crashes
#
$badDiskBlockCheckIntervalInMinutes = 30

# Indicates the number of minutes we would go back from current time
# to start checking events for IFilter error messages
#
$badIFilterCheckIntervalInMinutes = 60

# Event Id that indicates msftesql-exchange service found a bad Ifilter
#
$msftesqlBadIFilterEventIdEventId = 130

# Min number of Events required in the eventLog before an Ifilter is disabled
#
$msftesqlBadIFilterEventThreshold = 25

# CPU Affinity count for the msftesql process
#
$msftesqlCPUAffinityCount = 2

# Event Id that indicates msftesql-exchange service has crashed
#
$msftesqlCrashEventId = 1053

# Event Id that indicates msftesql-exchange master merge failed
#
$msftesqlMasterMergeFailedBadDiskEventId = 4104

# Source name to use for finding bad disk events in system log
#
$diskSourceName = "Disk"

# Event Id that indicates bad disk block error in system log
#
$badDiskEventId = 7

# Suffix that should be added by the TS to an Ifilter when it needs to disabled it
#
$disabledIFilterSuffix = "_DisabledByCITroubleshooter"

# The count threshold of the number of mailboxes left to crawl remains the same for consequetive runs before the TS will asssume that the service is stalled 
# and restart it
$stallDuringCrawlThreshold = 6

# Max threshold for the catalog size as a percentage of the overall database size before the 
# troubleshooter raises an alert
#
$maxPercentageCatalogSize = 20

# To avoid the MSFTESQL Process from consuming a lot of CPU we limit processor affinity.
# When the value is missing, it will use 33% of available logical CPU.
#
$affinityValue = $null

# Maximum msftefd cumulative memory consumption 
#
$MaxMsftefdMemoryConsumption = 6144
# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUhyicr6XirXbMjhaHbK6g0Wc8
# cqagghhqMIIE2jCCA8KgAwIBAgITMwAAAR4S9EGOKUc2xgAAAAABHjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzM2
# WhcNMjAwMTEwMjEwNzM2WjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046QUUyQy1FMzJCLTFBRkMxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQDzb28mkTnE/Qx2JfKv+ykRWkxyTx2Gt1TN7wBh/G2D9Z9J
# oWGDsHVyLMxbJzirFVKII+4mj6qKjCbMhWWhrUcKaez6q7hh6tpNL/knhCj48lZe
# FYMbSNAcY+cgHt5i+xV3kh6sZz1x8VrNopcripG3IjrbIJ5a47/NMUVCZyLwn+V8
# yz0elFnqH52amVPanbS0Re4Mku1U9IEOAdhlFd1AMfNL4kumj3GucM+W1rL9jsRO
# 9kgSsMDFwsM5lDAhn3toZcapx0yMi961g05xhpSmZ/hI4+szlAyqH0HN1CXjq2XQ
# 6PhRYcn4o+BdUzYbJ8rSrU3VwUhOzhv7hdTl5R2DAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQU3NGCgrg8lSUew8D8IjHvi1eXPDUwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAWR2z6mBmVxajzjhf
# 10GitwpRJTFbD1PjMopI0EPjoQiZUNQk4pxBLpQMSTv983jET+IHM6tU58aH8zI9
# 6GRYxgqVC4fNWWmZZq+OJ8+kLC95j65TbHarqzjZeW8x7nbZBd+l27sDbgyE99YA
# m9LwKecAYJY4IOcC2vl1CwdBzVMvnwN+mbgHw5X1hEdrjRODR0Fq0p/Yp6olDZ+4
# 8Wytf1U2gnOxM+3oMIg5OMnZ36pvAU05trHyX3/sx4vv/iKnuenE4tnK7MVqF4Jd
# u49bNdNxrivTf7UIluolvjaOIfnePwHajCAKRQLcHcD9LgtFg5PFEFhx64v52YnZ
# YAWspDCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU3/NHaJdXIbb0vMA8AONzhZEU3/Qw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAID/08y2RdemykmM12wfS//yf7jux2FkRe5FtFA0Dm+V
# 7euYcM+LgrinH4dBjZGeDoDCTytH/6GBiD/+tc8laRzEHtZ+3yBwQ7ikVZ/NCtbJ
# 1HjRNDuz5uuSfNbcQUOe2bk4ytGZXWG3KJyfruCjpydNJoZSiK1323h+P536V4Db
# 8+7VKdPM4VzVBIHr1/oTlx7Hk4VZMtLEPic4wy05ZG6wcngPKf2kYlshTz79OvL6
# bsIOUaJ0ewxVrnRPKcltj8MiESBAQwiwJsPfY38dcR0JuXygnZ74jUD3q6hO8t0c
# 7gKxmldL9bkcxl4KG0hA+l2p5MetdMlvFnJWZzNdOJOhggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# HhL0QY4pRzbGAAAAAAEeMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI1NTlaMCMGCSqGSIb3DQEJ
# BDEWBBSxghSvupx1XkaygJytH7rVk3+uRDANBgkqhkiG9w0BAQUFAASCAQDwpVNi
# TOxox80HWLTSmZ8yPQL7sYU5RhntrLTp0BMGWMt/M0GEHxit1POJnRDtHsMaq1O3
# B+C7zJlPlSHD3Igfe5Ly+UQAM/L2jY7nkbmWygVTFuuX7G8H2q+CtNF5PDfzAIdr
# CmN6sQOOLwbeOEXNTD4gyPMiU/qtWGZToZ6s8hh6G/Az70JqBId7yLGM4+a0Neiw
# 22quDSVDMnmUcRiT+H5rFmtb+5EBTD6pqHYGsQGlsoDoAEieV2zaR7Uxo3Kj8WMV
# xbx48jZ9BBNXEc94wFXpUKhwWusF8okpHCK8uM21OIz68XXT6LVqgwUdsWKinYgs
# IkBpVe6VQU98Fk83
# SIG # End signature block
