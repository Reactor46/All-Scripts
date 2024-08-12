# Copyright (c) 2010 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

# Requires -Version 2
[CmdletBinding(DefaultParametersetName="Default")]
PARAM(
    [parameter(
        ParameterSetName = "Default",
        Mandatory = $true,
        HelpMessage = "The database to troubleshoot.")]
    [String]
    [ValidateNotNullOrEmpty()]
    $MailboxDatabaseName,

    [parameter(
        ParameterSetName = "Default",
        Mandatory = $true,
        HelpMessage = "The maximum RPC average latency the server should be experiencing. (Typical thresholds might be 70 or 150.)")]
    [int]
    [ValidateRange(1, 200)]
    [ValidateNotNullOrEmpty()]
    $LatencyThreshold,

    [parameter(
        ParameterSetName = "NoOp",
        Mandatory = $true,
        HelpMessage = "Loads helper libraries, then quits. Used to verify proper deployment/installation.")]
    [switch]
    $Deploy,

    [int]
    [ValidateRange(1,600000)]
    [ValidateNotNullOrEmpty()]
    $TimeInServerThreshold,

    [int]
    [ValidateRange(1,500)]
    [ValidateNotNullOrEmpty()]
    $OperationPerSecondThreshold,
    
    [parameter(
        Mandatory = $false,
        HelpMessage = "Minimum number of samples of StoreUsageStatistics a user should have to be able to positively identify a single user problem.")]
    [int]
    [ValidateRange(1,10)]
    [ValidateNotNullOrEmpty()]
    $MinimumUserSampleCount,
    
    [parameter(
        Mandatory = $false,
        HelpMessage = "Minimum number of samples of StoreUsageStatistics in the set that was collected.")]
    [int]
    [ValidateRange(1,250)]
    [ValidateNotNullOrEmpty()]
    $MinimumStoreUsageStatisticsSampleCount,
    
    [parameter(
        Mandatory = $false,
        HelpMessage = "Percent of samples below rop latency threshold at which we want to investigate further.")]
    [int]
    [ValidateRange(1,100)]
    [ValidateNotNullOrEmpty()]
    $PercentSampleBelowThresholdToAlert,
    
    [parameter(
        Mandatory = $false,
        HelpMessage = "Whether or not to quarantine heavy users.")]
    [switch]
    $Quarantine,

    [parameter(
        Mandatory = $false,
        HelpMessage = "Whether or not we're running under the monitoring context.")]
    [switch]
    $MonitoringContext,

    [parameter(
        Mandatory = $false,
        HelpMessage = "Whether or not to quarantine heavy users.")]
    [String]
    $QuarantineString,

    [parameter(
        Mandatory = $false,
        HelpMessage = "Whether or not we're running under the monitoring context.")]
    [String]
    $MonitoringContextString,
    
    [parameter(
        Mandatory = $false,
        HelpMessage = "Alert Guid if executed in response to an alert.")]
    [String]
    $AlertGUID
)

###################################################################################################################################
#                                                                                                                                 #
#                                                     Script Body                                                                 #
#                                                                                                                                 #
###################################################################################################################################

    Set-StrictMode -Version Latest

    if ($QuarantineString -eq "true")
    {
        $Quarantine = $true;
    }
    if ($MonitoringContextString -eq "true")
    {
        $MonitoringContext = $true;
    }
        
    $scriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)

    . $scriptPath\CITSLibrary.ps1
    . $scriptPath\StoreTSLibrary.ps1
    . $scriptPath\StoreTSConstants.ps1
    . $scriptPath\DiagnosticScriptCommonLibrary.ps1

    Load-ExchangeSnapin

    # Since we're in strict mode we must declare all variables we use
    $script:monitoringEvents = $null

    if ($PSCmdlet.ParameterSetName -eq "NoOp")
    {
        if ($MonitoringContext)
        {
            #E14:303295 Add a dummy monitoring event to supress SCOM failure alerts
            Add-MonitoringEvent -Id $StoreLogEntries.DatabaseSpaceTroubleShooterStarted[0] -Type $EVENT_TYPE_INFORMATION -Message "Latency TS in No-op mode"
            Write-MonitoringEvents
        }
        
        return
    }

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
        
    $retries = 0
    
    do
    {
        $success = $true
        $error.Clear()    
    
        $counterValues = get-counter -ComputerName $database.MountedOnServer -Counter $counterNames -MaxSamples 10 -ErrorAction SilentlyContinue
    
        #Verify there were no errors when trying to get the perf counters
        #Even if we are unsuccessful getting the perf counters and come
        #out of this loop, it will be handled properly when we try to get
        #the average results for the counters. This way, we handle the case
        #where we at least got some counters back and try to go as far into 
        #the TS as possible
        if ($error.Count -gt 0)
        {
            #Check to see if the error was related to the database being failed over,
            #in this case, we should get CounterPathIsInvalid (E14 Bug: 468424)
            if ($error[0].FullyQualifiedErrorId -match "CounterPathIsInvalid")
            {            
                #Get the computername database is currently on to verify that the
                #database indeed moved
                $currentDatabase = Get-MailboxDatabase $MailboxDatabaseName -Status
                if ($currentDatabase.MountedOnServer -ne $database.MountedOnServer)
                {                
                    #Log an event for TS not able to run due to database moved
                    $failureReason = ($StoreLocStrings.DatabaseMoved + $error[0])
                    Log-Event `
                        -Arguments $Arguments `
                        -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterNotRunDatabase `
                        -Parameters @($database, $failureReason)
                        
                    if ($MonitoringContext)
                    {
                        Write-MonitoringEvents
                    }
            
                    #Clear the error so that it does not get escalated
                    $error.Clear()
                    
                    return
                }
            }
            elseif ($error[0].FullyQualifiedErrorId -match "CounterApiError")
            {
                #This could be due to stale powershell session or some other transient issue
                #Let's try a few more times to get the counters after sleeping for a few
                #seconds between each retries, with a total sleep time of 30 seconds
                $retries++
                $success = $false
                
                if ($retries -lt 5)
                {
                    $sleepSeconds = $retries * 3
                    write-verbose ("Unable to get perf counters after {0} retries, Re-trying after sleeping for {1} seconds." -f $retries, $sleepSeconds)
                    Start-Sleep -Seconds $sleepSeconds
                }
            }
            else
            {
                #Not the expected error, throw so recovery workflow will escalate
                if ($MonitoringContext)
                {
                    Write-MonitoringEvents
                }
                        
                throw $error[0]
            }
        }    
    } while ((-not $success) -and $retries -lt 5)

    #Check to see if we got any perf counters back
    #If we didn't, does not make sense to continue
    #executing the troubleshooter
    if (!$counterValues)
    {
        if ($MonitoringContext)
        {
            Write-MonitoringEvents
        }
        
        $failureReason = "{0} Error: {1}" -f $StoreLocStrings.UnableToGetAnyCounters, $error[0]
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterNotRunDatabase `
            -Parameters @($database, $failureReason)
                        
        throw $failureReason
    }

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
    if (!$MyInvocation.BoundParameters.ContainsKey("OperationPerSecondThreshold"))
    {
        $OperationPerSecondThreshold = $OperationPerSecondDefaultThreshold
    }
    
    $rpcOpsPerSecond = Get-AverageResultForCounter -results $counterValues -counter $rpcOpsPerSecondCounterName
    if ($rpcOpsPerSecond -lt $OperationPerSecondThreshold)
    {
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterLowOps `
            -Parameters @($database, $rpcLatency, $LatencyThreshold, $rpcOpsPerSecond, $OperationPerSecondThreshold)

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

    # Run get-storeusagestatistics to find the user who is not already quarantined, has TimeInServer > Threshold
    # and RopLatency (TimeInServer/RopCount) > Threshold when averaged across all samples and is the top user
    if (!$MyInvocation.BoundParameters.ContainsKey("PercentSampleBelowThresholdToAlert"))
    {
        $PercentSampleBelowThresholdToAlert = $PercentSampleBelowThresholdToAlertDefault
    }
    
    if (!$MyInvocation.BoundParameters.ContainsKey("MinimumStoreUsageStatisticsSampleCount"))
    {
        $MinimumStoreUsageStatisticsSampleCount = $MinimumStoreUsageStatisticsSampleCountDefault
    }
        
    $topCpuUsers = @(Get-TopCpuUsers $database -TimeInServerThreshold $TimeInServerThreshold -ROPLatencyThreshold $ROPLatencyThreshold -PercentSampleBelowThresholdToAlert $PercentSampleBelowThresholdToAlert -MinimumStoreUsageStatisticsSampleCount $MinimumStoreUsageStatisticsSampleCount)
    
    # See if there is one or more user with TimeInServer and ROP Latency greater than the threshold
    # If no user in the list, there is no impact to anybody even with high latency
    # It could likely be SystemMailbox causing this. In any case, we don't care so
    # just suppress
    if ($topCpuUsers.Length -eq 0)
    {
        # Log an event and return
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterNoMailbox `
                -Parameters @($database, $rpcLatency)

            if ($MonitoringContext)
            {
                Write-MonitoringEvents
            }
            
            return
    }    
    else
    {
        # If we have a top user and we are down here, the only action left to take is quarantine *only that user*. 
        # Fire an event indicating quarantine. Exit
        # In production, we have quarantine set to false. So if we get this far, we are going to get a paging
        # or non-paging alert (based on RPC latency) with a pointer to saved SUS data.
        if ($Quarantine -eq $true)
        {
            # Quarantine the mailbox
            write-verbose ("Quarantining: " + $topCpuUsers[0].MailboxGuid + " due to following:")
            write-verbose ("TotalTimeInServer: " + $topCpuUsers[0].TotalTimeInServer)
            write-verbose ("AverageTimeInServer: " + $topCpuUsers[0].AverageTimeInServer)
            write-verbose ("TimeInServerThreshold: " + $TimeInServerThreshold)
            write-verbose ("RPC Operations per sec: " + $rpcOpsPerSecond)
                        
            Enable-MailboxQuarantine -Identity $topCpuUsers[0].MailboxGuid.ToString() -Confirm:$false -ErrorAction Stop

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
        }
    }

    # If we are down here, let's save store usage statistics and send mail to on-calls
    # so that it can be analysed irrespective of whether we quarantine the mailbox or not.
    # We only want to do this in prod.
    Log-Event `
        -Arguments $Arguments `
        -EventInfo $StoreLogEntries.DatabaseLatencyTroubleShooterIneffective `
        -Parameters @($database, $rpcLatency, $rpcOpsPerSecond)
        
    if ($MonitoringContext -and $global:StoreUsageStatsTS)
    {
        $exchangeInstallPath = (get-item HKLM:\SOFTWARE\Microsoft\ExchangeServer\V14\Setup).GetValue("MsiInstallPath")
        $StoreLibraryPath = Join-Path $exchangeInstallPath "Datacenter\StoreCommonLibrary.ps1"
        if (Test-Path $StoreLibraryPath)
        {
            # Dot-source common libraries to populate alert details.
            . ($StoreLibraryPath)
            
            # Get store usage statistics and send mail
            $statsPath = Output-StoreUsageStatistics $global:StoreUsageStatsTS $alertGUID
            
            # Because of E14:497368, we are going to just throw here so that proper alert mail gets sent
            # to the on-calls instead of us trying to populate alert details.
            $failureReason = "{0} {1}" -f $StoreLocStrings.UnableToAnalyze, $statsPath
            Write-MonitoringEvents
            throw $failureReason
        }
    }
    
    if ($MonitoringContext)
    {
        Write-MonitoringEvents

        # Throw an exception to make the workflow that invoked us go to its failure state.
        # This should result in an escalation - either urgent or non-urgent, depending on what
        # the health manifest specified in the RecoveryAction.
        throw "Troubleshoot-DatabaseLatency failed. See the event log for further details."
    }


# SIG # Begin signature block
# MIIdwAYJKoZIhvcNAQcCoIIdsTCCHa0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUpW6nGzvZZwYSous7ihM4fagt
# kU+gghhkMIIEwzCCA6ugAwIBAgITMwAAAJqamxbCg9rVwgAAAAAAmjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI5
# WhcNMTcwNjMwMTkyMTI5WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkIxQjctRjY3Ri1GRUMyMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEApkZzIcoArX4o
# w+UTmzOJxzgIkiUmrRH8nxQVgnNiYyXy7kx7X5moPKzmIIBX5ocSdQ/eegetpDxH
# sNeFhKBOl13fmCi+AFExanGCE0d7+8l79hdJSSTOF7ZNeUeETWOP47QlDKScLir2
# qLZ1xxx48MYAqbSO30y5xwb9cCr4jtAhHoOBZQycQKKUriomKVqMSp5bYUycVJ6w
# POqSJ3BeTuMnYuLgNkqc9eH9Wzfez10Bywp1zPze29i0g1TLe4MphlEQI0fBK3HM
# r5bOXHzKmsVcAMGPasrUkqfYr+u+FZu0qB3Ea4R8WHSwNmSP0oIs+Ay5LApWeh/o
# CYepBt8c1QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFCaaBu+RdPA6CKfbWxTt3QcK
# IC8JMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAIl6HAYUhsO/7lN8D/8YoxYAbFTD0plm82rFs1Mff9WBX1Hz
# /PouqK/RjREf2rdEo3ACEE2whPaeNVeTg94mrJvjzziyQ4gry+VXS9ZSa1xtMBEC
# 76lRlsHigr9nq5oQIIQUqfL86uiYglJ1fAPe3FEkrW6ZeyG6oSos9WPEATTX5aAM
# SdQK3W4BC7EvaXFT8Y8Rw+XbDQt9LJSGTWcXedgoeuWg7lS8N3LxmovUdzhgU6+D
# ZJwyXr5XLp2l5nvx6Xo0d5EedEyqx0vn3GrheVrJWiDRM5vl9+OjuXrudZhSj9WI
# 4qu3Kqx+ioEpG9FwqQ8Ps2alWrWOvVy891W8+RAwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBMYwggTCAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB2jAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUzC9go82zaE7pbBM711DNkZRXZzowegYKKwYB
# BAGCNwIBDDFsMGqgQoBAAFQAcgBvAHUAYgBsAGUAcwBoAG8AbwB0AC0ARABhAHQA
# YQBiAGEAcwBlAEwAYQB0AGUAbgBjAHkALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBACidwn2umiPc
# 8CQOE8qSTOpZNepdNTBR5q3P3S8gcv3wKHCuVgugWofVFnpncail5INcTMFZtMTe
# 7wVc/AFVklQQniPmIU1drHc6o76rI2RTtBLDeKL4+NQ0VzOPiCJ1Sjq39Ca4knai
# PErsxc7sUTD3L130ssDk3sqW1j6sSTu/M7FDZTeFK0Pjme1Pf8NbKxYWbqlfOHAX
# 9iBC4XIauKdApqos9WvMVvexv43zG8P59UtonaXChanxVRgqbkJ0xPaxYo+Tlxui
# KR+SgEJzcxkeGfuAc/Ygd3IBuOlhMow3I+ezlOrXIF2wyoSaIZeGTtAS2cvsbchc
# gaLY7pVbZ/ehggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAmpqbFsKD2tXCAAAAAACaMAkGBSsOAwIa
# BQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0x
# NjA5MDMxODQ0NTVaMCMGCSqGSIb3DQEJBDEWBBTFrAJdir9r4nDgtw9ZNspIn2UY
# /jANBgkqhkiG9w0BAQUFAASCAQCSOwLAIDPvT8AdbxEWOICVyJEHlABuNLRhGa3+
# aOlYwsJDraY2G3Q5r1EZPxht1HMMjsprm8zwuSeKXMzAZAIlrQN8HXXOB+PaPhTP
# ajGNttl2Yfd/V8gXLnlVY5il5sdYCEJfcCZuzwZiVVwZt42Fx0yJtYBuJB7yei3n
# l5TylOXjXkwWP2CJoZQPgAs036LmJpatqmmysZ+s1vXMaomajY+pLcp7NosoThnR
# dJcgfNcUXXdW4x5SBqcRdh0N9xO8AFA9C6JQZNcgwyssb11Np8Vqurzd9DiOlpyo
# iv3614l8YCHWulwAC92/xn8lH5tQdHSNj/HJn8gVWDKTyN8y
# SIG # End signature block
