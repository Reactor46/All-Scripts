# Copyright (c) 2010 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

# Requires -Version 2
PARAM(
    [parameter(
        Mandatory = $true,
        HelpMessage = "The database to troubleshoot.")]
    [String]
    [ValidateNotNullOrEmpty()]
    $MailboxDatabaseName,

    [parameter(
        Mandatory = $false,
        HelpMessage = "The maximum RPC average latency the server should be experiencing.")]
    [int]
    [ValidateRange(1, 200)]
    [ValidateNotNullOrEmpty()]
    $LatencyThreshold,

    [int]
    [ValidateRange(1,600000)]
    [ValidateNotNullOrEmpty()]
    $TimeInServerThreshold,

    [parameter(
        Mandatory = $false,
        HelpMessage = "Whether or not to quarantine heavy users.")]
    [switch]
    $Quarantine,

    [parameter(
        Mandatory = $false,
        HelpMessage = "Whether or not we're running under the monitoring context.")]
    [switch]
    $MonitoringContext
)

###################################################################################################################################
#                                                                                                                                 #
#                                                     Script Body                                                                 #
#                                                                                                                                 #
###################################################################################################################################

    Set-StrictMode -Version Latest

    $exchangeInstallPath = (get-item -path env:ExchangeInstallPath).Value
    $scriptPath = join-path $exchangeInstallPath "Scripts"

    . $scriptPath\CITSLibrary.ps1
    . $scriptPath\StoreTSLibrary.ps1
    . $scriptPath\StoreTSConstants.ps1
    . $scriptPath\DiagnosticScriptCommonLibrary.ps1

    Load-ExchangeSnapin

    # Since we're in strict mode we must declare all variables we use
    $script:monitoringEvents = $null

    $database = Get-MailboxDatabase $MailboxDatabaseName -Status

    if ($database -eq $null)
    {
        $argError = new-object System.ArgumentException ($error[0].Exception.ToString())
        throw $argError
    }

    if (!$MyInvocation.BoundParameters.ContainsKey("TimeInServerThreshold"))
    {
        $TimeInServerThreshold = $TimeInServerDefaultThreshold
    }

    if (!$MyInvocation.BoundParameters.ContainsKey("LatencyThreshold"))
    {
        $LatencyThreshold = $RPCLatencyDefaultThreshold
    }

    # Event log source name for application log
    $appLogSourceName = "Database Latency Troubleshooter"

    # Event log source name for crimson log
    $crimsonLogSourceName = "Database Latency"

    # The Arguments object is needed for logging
    # events.
    $Arguments = new-object -typename Arguments

    $Arguments.Server = $database.MountedOnServer
    $Arguments.Database = $database
    $Arguments.MonitoringContext = $MonitoringContext

    # Since this TS doesn't run in SCOM we need to write App event logs
    # for alerts to fire
    $Arguments.WriteApplicationEvent = $MonitoringContext

    Log-Event `
        -Arguments $Arguments `
        -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterStarted `
        -Parameters @($database)

    $rpcLatencyCounterName = "\MSExchangeIS Mailbox($database)\rpc average latency"
    $rpcOpsPerSecondCounterName = "\MSExchangeIS Mailbox($database)\RPC Operations/sec"
    $readLatencyCounterName = "\MSExchange Database ==> Instances($database)\I/O Database Reads Average Latency"
    $readRateCounterName = "\MSExchange Database ==> Instances($database)\I/O Database Reads/sec"
    $writeLatencyCounterName = "\MSExchange Database ==> Instances($database)\I/O Database Writes Average Latency"
    $writeRateCounterName = "\MSExchange Database ==> Instances($database)\I/O Database Writes/sec"
    $dsaccessLatencyCounterName = "\MSExchangeIS\dsaccess average latency"
    $dsaccessCallsCounterName = "\MSExchangeIS\dsaccess active call count"

    $counterNames = @($rpcLatencyCounterName, $rpcOpsPerSecondCounterName, $readLatencyCounterName, $readRateCounterName, $writeLatencyCounterName, $writeRateCounterName, $dsaccessLatencyCounterName, $dsaccessCallsCounterName)
    $counterValues = get-counter -ComputerName $database.MountedOnServer -Counter $counterNames -MaxSamples 10

    #Check that latencies are still high
    $rpcLatency = Get-AverageResultForCounter -results $counterValues -counter $rpcLatencyCounterName
    if ($rpcLatency -lt $LatencyThreshold)
    {
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterNoLatency `
            -Parameters @($rpcLatency, $database, $LatencyThreshold)

        if ($MonitoringContext)
        {
            Write-MonitoringEvents
        }

        return
    }

    # Check that Rpc operations/sec are high enough to monitor.
    $rpcOpsPerSecond = Get-AverageResultForCounter -results $counterValues -counter $rpcOpsPerSecondCounterName
    if ($rpcOpsPerSecond -lt $OperationPerSecondDefaultThreshold)
    {
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterLowOps `
            -Parameters @($database, $rpcLatency, $LatencyThreshold, $rpcOpsPerSecond, $OperationPerSecondDefaultThreshold)

        if ($MonitoringContext)
        {
            Write-MonitoringEvents
        }

        return
    }

    # check if disk transfers/sec < X and disk secs/transfer > Y... if yes, disk is bad.
    $readLatency = Get-AverageResultForCounter -results $counterValues -counter $readLatencyCounterName
    $readRate = Get-AverageResultForCounter -results $counterValues -counter $readRateCounterName
    $writeLatency = Get-AverageResultForCounter -results $counterValues -counter $writeLatencyCounterName
    $writeRate =  Get-AverageResultForCounter -results $counterValues -counter $writeRateCounterName

    # We don't want to report a bad disk if the disk is being overloaded
    # so check to see if we have a lot of read/write requests to the disk
    # if we don't have many requests and either latency (read or write)
    # is above our thresholds then we have a bad disk.
    if ((($readRate -lt $DiskReadRateThreshold) -and
            ($writeRate -lt $DiskWriteRateThreshold)) `
            -and
        (($readLatency -gt $DiskReadLatencyThreshold) -or
            ($writeLatency -gt $DiskWriteLatencyThreshold)))
    {
        # DatabaseLatencyTroubleShooterBadDiskLatencies is only reporting read rate
        # and read latency, it should be really report write rate and latencies as well
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterBadDiskLatencies `
            -Parameters @($database, $readLatency, $readRate, $rpcLatency)

        if ($MonitoringContext)
        {
            Write-MonitoringEvents
        }

        return
    }

    # Look for high AD latencies and/or call count.
    $dsaccessLatency = Get-AverageResultForCounter -results $counterValues -counter $dsaccessLatencyCounterName
    $dsaccessCalls =  Get-AverageResultForCounter -results $counterValues -counter $dsaccessCallsCounterName

    # Report a very-high call count, or a medium-high call count in combination with a medium-high latency.
    if (($dsaccessCalls -gt $DSAccessCallsStandaloneThreshold) -or
        (($dsaccessCalls -gt $DSAccessCallsCombinedThreshold) -and
         ($dsaccessLatency -gt $DSAccessLatencyCombinedThreshold)))
    {
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterBadDSAccessActiveCallCount `
            -Parameters @($database, $dsaccessLatency, $dsaccessCalls, $rpcLatency)

        if ($MonitoringContext)
        {
            Write-MonitoringEvents
        }

        return
    }

    # Report a very-high latency, even without a high call count. This test must come after the combined test.
    if ($dsaccessLatency -gt $DSAccessLatencyStandaloneThreshold)
    {
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterBadDSAccessAverageLatency `
            -Parameters @($database, $dsaccessLatency, $dsaccessCalls, $rpcLatency)

        if ($MonitoringContext)
        {
            Write-MonitoringEvents
        }

        return
    }

    # run get-storeusagestatistics to find the user who is not already quarantined, has TimeInServer > Threshold
    # when averaged across all samples and is the top user.. if yes, quarantine *only that user*.
    # Fire an event indicating quarantine. Exit
    $topCpuUsers = @(Get-TopCpuUsers $MailboxDatabaseName -TimeInServerThreshold $TimeInServerThreshold)

    if ($topCpuUsers.Length -gt 0)
    {
        if ($Quarantine -eq $true)
        {
            write-verbose ("Quarantining: " + $topCpuUsers[0].MailboxGuid)
            Set-QuarantineMailbox $topCpuUsers[0].MailboxGuid

            Log-Event `
                -Arguments $Arguments `
                -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterQuarantineUser `
                -Parameters @($topCpuUsers[0].MailboxGuid, $database, $topCpuUsers[0].AverageTimeInServer, $rpcLatency, $rpcOpsPerSecond)

            if ($MonitoringContext)
            {
                Write-MonitoringEvents
            }

            return
        }
        else
        {
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterNoQuarantine `
                -Parameters @($topCpuUsers[0].MailboxGuid, $database, $topCpuUsers[0].AverageTimeInServer, $rpcLatency, $rpcOpsPerSecond)

            if ($MonitoringContext)
            {
                Write-MonitoringEvents
            }

            return
        }
    }

    Log-Event `
        -Arguments $Arguments `
        -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterIneffective `
        -Parameters @($database, $rpcLatency, $rpcOpsPerSecond)

    if ($MonitoringContext)
    {
        Write-MonitoringEvents
    }
# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUirpvbnLsqxR5D/PHznOLB0R/
# CbmgghhqMIIE2jCCA8KgAwIBAgITMwAAASDzON/Hnq4y7AAAAAABIDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzM4
# WhcNMjAwMTEwMjEwNzM4WjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046MjI2NC1FMzNFLTc4MEMxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQCO1OidLADhraPZx5FTVbd0PlB1xUfJ0J9zuRe1282yigKI
# +r7rvHTBllcSjV+E6G3BKO1FX7oV2CGaAGduTl2kk0vGSlrXC48bzR0SAb1Ui49r
# bUJTA++yfZA+34s8vYUye1XX2T5D0GKukK1hLkf8d7p2A5nygvMtnnybzmEVavSd
# g8lYzjK2EuekiLzL/lYUxAp2vRNFUitr7MHix5iU2nHEG4yU8crlXjYFgJ7q3CFv
# Il1yMsP/j+wk+1oCC1oLV6iOBcpq0Nxda/o+qN78nQFoQssfHoA9YdBGUnRHk+dK
# Sq5+GiV3AY0TRad2ZRzLcIcNmUJXny26YG+eokTpAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUIkw9WwdWW+zV8Il/Jq7A7bh6G7cwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAE4tuQuXzzaC2OIk4
# ZhJanhsgQv9Tk8ns/9elb8pAgYyZlSwxUtovV8Pd70jtAt0U/wjGd9n+QQJZKILM
# 6WCIieZFkZbqT9Ut9zA+tc2eQn4mt62PlyA+YJZNHEiPZhwgbjfLIwMRsm845B4N
# KN7WmfYwspHdT/mPgLWaBsSWS80PuAtpG3N+o9eTHskT+qauYAMqhZExfI8S2Rg4
# kdqAm7EU/Nroe4g0p+eKw6CAQ2ZuhuqHMMPgcQlSejcEbpS5WAzdCRd6qDXPHh0r
# C3FayhXrwu/KKuNW2hR1ZCx/ieNiR8+lWt1JxXgWAttgaRtR3VqGlL4aolg41UCo
# XfN1IjCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU2kW/dvLfXqv2xaHArmu0wc6jpW8w
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBACYf2giCl+tweB/Xg0vy4Vo1jE/h0LTxph9MpEny9mEZ
# VGFiRbkzxxyd3xr1kCLPCNPIWDIYxDbe4qxCTLX6KZdz0gSVR6aX+BfBuSSavm5u
# 4445SNEp8dK61+fbpEk5S2RTQ3GSCPUGK4jniIkBsPPDl8X19lA7Q5Rt4isYrI5E
# jRkQMg+kftUa8MOIe1eXBA1ITZ0YlBHugBRq7TizDAdEuOuI2OLD5OFXbfjXnKsG
# vXY+fiU87jfFf8Xxg2elK3b7faaNG8WD1t6Hvp3WKdhgGFSVfV8e3qwCnEEOfnHA
# Xe5nV126Ak5GAgD6uDU5yUAeaTw+ee8B6HbUDdXyHVmhggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# IPM438eerjLsAAAAAAEgMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI1NThaMCMGCSqGSIb3DQEJ
# BDEWBBTOnuV6pXZvDTDFdKenRO0HBcXAIzANBgkqhkiG9w0BAQUFAASCAQAYL49j
# w3Cup3BdUCnskLUY62tKhOa7Zajd9UnydPzMV/+TbmJcBGwATNuki909Txk5TXU/
# Lzltli/9eBAXheJ1FWRTHlmJr2D1Og6gxMRc0CqxJFeGnE2fNIhrhDDEkFCihLin
# xoS29iqQf5gjESG3d3AA6vPCFyDLYqSd8+b+ruFJUvByGQX3+zerhD29JJOXUEqu
# 7uXWKktxbFUB0b3aTHbPRJjHCwBKDxE2uY9Siqb9/AFcXQXbJkszPLywqw6Jqb+x
# VxD1wZ2J6XPd89xsIPDKEQafnIRHNj/L8P+K3Z+3qXuZxj2kyjUtsnunjn6uA2X9
# 1MW0W61mp/912siy
# SIG # End signature block
