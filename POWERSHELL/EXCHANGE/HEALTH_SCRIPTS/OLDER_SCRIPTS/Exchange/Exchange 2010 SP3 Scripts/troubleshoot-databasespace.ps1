# Copyright (c) 2010 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

# Requires -Version 2
PARAM(
    [parameter(
        ParameterSetName = "Default",
        Mandatory = $true,
        HelpMessage = "The database to troubleshoot.")]
    [String]
    [ValidateNotNullOrEmpty()]
    $MailboxDatabaseName,
    
    [parameter(
        ParameterSetName = "Server",
        Mandatory = $true,
        HelpMessage = "The mailbox server to troubleshoot.")]
    [String]
    [ValidateNotNullOrEmpty()]
    $Server,

    [parameter(
        Mandatory = $false,
        HelpMessage = "The percentage of disk space for the EDB file at which we should start quarantining users.")]
    [int]
    [ValidateRange(1, 99)]
    [ValidateNotNullOrEmpty()]
    $PercentEdbFreeSpaceThreshold,
    
    [parameter(
        Mandatory = $false,
        HelpMessage = "The percentage of disk space for the logs at which we should start quarantining users.")]
    [int]
    [ValidateRange(1, 99)]
    [ValidateNotNullOrEmpty()]
    $PercentLogFreeSpaceThreshold,

    [parameter(
        Mandatory = $false,
        HelpMessage = "The number of hours we can wait before running out of space.")]
    [int]
    [ValidateRange(1, 1000000000)]
    [ValidateNotNullOrEmpty()]
    $HourThreshold,

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
    
    #
    # Check a single database for free space
    #
    function Troubleshoot-DatabaseCopySpace([Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $database, [int] $PercentEdbFreeSpaceThreshold, [int] $PercentLogFreeSpaceThreshold, [int] $HourThreshold, [bool] $Quarantine, [bool] $MonitoringContext)
    {
        if ($database -eq $null)
        {
            $argError = new-object System.ArgumentException ($error[0].Exception.ToString())
            throw $argError
        }
        
        # Event log source name for application log
        $appLogSourceName = "Database Space Troubleshooter"

        # Event log source name for crimson log
        $crimsonLogSourceName = "Database Space"

        # The Arguments object is needed for logging 
        # events.
        $Arguments = new-object -typename Arguments
        $Arguments.Server = $database.MountedOnServer
        $Arguments.Database = $database
        $Arguments.MonitoringContext = $MonitoringContext
        $Arguments.WriteApplicationEvent = $false
            
        # Check the Edb drive for space
        $edbGrowthRateThreshold = -1
        $edbVolume = Get-WmiVolumeFromPath $database.EdbFilePath $database.MountedOnServer
        $Arguments.InstanceName = $edbVolume.Name -replace "\\$", ""	# Remove any trailing backslash if exists since SCOM discovery removes it from the instance name
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $StoreLogEntries.DatabaseSpaceTroubleShooterStarted `
            -Parameters @($Arguments.InstanceName, $database)
        $edbVolumeFreeSpace = $edbVolume.FreeSpace + $database.AvailableNewMailboxSpace
        $edbPercentFreeSpace = ($edbVolumeFreeSpace * 100) / $edbVolume.Capacity

        # Check the Log drive for space
        $logGrowthRateThreshold = -1
        $logVolume = Get-WmiVolumeFromPath $database.LogFolderPath $database.MountedOnServer
        $logVolumeFreeSpace = $logVolume.FreeSpace
        $logPercentFreeSpace = ($logVolumeFreeSpace * 100) / $logVolume.Capacity
        if ($logPercentFreeSpace -lt $PercentLogFreeSpaceThreshold)
        {
            $logGrowthRateThreshold = $logVolume.FreeSpace / $HourThreshold
        }
        
        $growthRateThreshold = -1

        # Figure out which of the 2 thresholds is lower
        # to determine whether or not we need to quarantine 
        # some users
        if ($logGrowthRateThreshold -ne -1) 
        {
            $growthRateThreshold = $logGrowthRateThreshold
        }
        
        if (($edbGrowthRateThreshold -ne -1) -and
            ($edbGrowthRateThreshold -lt $logGrowthRateThreshold))
        {
            $growthRateThreshold = $edbGrowthRateThreshold 
        }   
        
        $counterValue = 0
        $counterName = "\MSExchangeIS Mailbox($database)\JET Log Bytes Generated/hour"
        $logBytesCounter = get-counter -ComputerName $database.MountedOnServer -Counter $counterName -MaxSamples 10
        $countervalue = Get-AverageResultForCounter -results $logBytesCounter -counter $counterName

        if ($null -eq $countervalue)
        {
            $countervalue = 0
        }
        
        write-verbose ("Current Growth Rate: " + $counterValue)
        write-verbose ("Growth Rate Threshold: " + $growthRateThreshold)
                
        # Check if we are low on space compared to any of the thresholds
        $problemDetected = $false
        if ($edbPercentFreeSpace -lt $PercentEdbFreeSpaceCriticalThreshold)
        {
            # Log an event that will trigger paging alert for critical space issue
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $StoreLogEntries.DatabaseSpaceTroubleDetectedCriticalSpaceIssue `
                -Parameters @($Arguments.InstanceName, $database, $PercentEdbFreeSpaceCriticalThreshold)
                
            $problemDetected = $true
        }
        elseif ($edbPercentFreeSpace -lt $PercentEdbFreeSpaceAlertThreshold)
        {
            # Log an event that will trigger non-paging alert for low space issue
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $StoreLogEntries.DatabaseSpaceTroubleDetectedAlertSpaceIssue `
                -Parameters @($Arguments.InstanceName, $database, $PercentEdbFreeSpaceAlertThreshold)
               
            $problemDetected = $true
        }
        elseif ($edbPercentFreeSpace -lt $PercentEdbFreeSpaceThreshold)
        {
            $edbGrowthRateThreshold = $edbVolumeFreeSpace / $HourThreshold
            
            # Log a warning event for low space issue
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $StoreLogEntries.DatabaseSpaceTroubleDetectedWarningSpaceIssue `
                -Parameters @($Arguments.InstanceName, $database, $PercentEdbFreeSpaceThreshold)
                
            $problemDetected = $true
        }
        
        # Format to show only 2 decimal places in event log
        $edbPercentFreeSpace = "{0:N2}" -f $edbPercentFreeSpace
        $logPercentFreeSpace = "{0:N2}" -f $logPercentFreeSpace
        $edbVolumeFreeSpaceInGB = "{0:N2}" -f ($edbVolumeFreeSpace/1GB)
        $logVolumeFreeSpaceInGB = "{0:N2}" -f ($logVolumeFreeSpace/1GB)
        
        if ($problemDetected)
        {
            # Exclude the database from provisioning as we are below one of the space thresholds
            Set-MailboxDatabase $database -IsExcludedFromProvisioning $true
        }
        else
        {
            # We have enough free space that we are not going to even bother looking for outliers causing rapid growth to try to qurantine.
            if ($growthRateThreshold -eq -1)
            {
                Log-Event `
                    -Arguments $Arguments `
                    -EventInfo $StoreLogEntries.DatabaseSpaceTroubleShooterNoProblemDetected `
                    -Parameters @($Arguments.InstanceName, $database, $edbVolumeFreeSpaceInGB, $edbPercentFreeSpace, $logVolumeFreeSpaceInGB, $logPercentFreeSpace, $PercentEdbFreeSpaceThreshold, $PercentLogFreeSpaceThreshold, $HourThreshold, $counterValue, $PercentEdbFreeSpaceCriticalThreshold, $PercentEdbFreeSpaceAlertThreshold, $QuarantineValue)
                
                return
            }
        }
    	
	    $cUsersQuarantined = 0
        $currentGrowthRate = $counterValue

        # Do we need to quarantine users?
        if (($currentGrowthRate -lt $growthRateThreshold))
        {   
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $StoreLogEntries.DatabaseSpaceTroubleShooterFoundLowSpaceNoQuarantine `
                -Parameters @($Arguments.InstanceName, $database, $edbVolumeFreeSpaceInGB, $edbPercentFreeSpace, $logVolumeFreeSpaceInGB, $logPercentFreeSpace, $PercentEdbFreeSpaceThreshold, $PercentLogFreeSpaceThreshold, $HourThreshold, $growthRateThreshold, $counterValue, $currentGrowthRate, $PercentEdbFreeSpaceCriticalThreshold, $PercentEdbFreeSpaceAlertThreshold, $QuarantineValue)
                
            return
        }

        # We have to quarantine some users now, so
        # get the least of the most active users
        # and figure out which ones we need to quarantine
        # so that the current log generation rate drops
        # below our maximum calculated growth rate
        $topLogGenerators = @(Get-TopLogGenerators $database)
        $iNextUserToQuarantine = 0

        if ($Quarantine)
        {
            for($iLogGenerator = 0; $iLogGenerator -lt $topLogGenerators.Length; $iLogGenerator++)
            {   
                write-verbose ("Current Growth Rate: " + $currentGrowthRate)
                write-verbose ("Growth Rate Threshold: " + $growthRateThreshold)
                write-verbose ("Top user: " + $topLogGenerators[$iLogGenerator].MailboxGuid)
                write-verbose ("Top user logs: " + $topLogGenerators[$iLogGenerator].TotalLogBytes)
                
                if ($currentGrowthRate -lt $growthRateThreshold)
                {
                    Log-Event `
                        -Arguments $Arguments `
                        -EventInfo $StoreLogEntries.DatabaseSpaceTroubleShooterFoundLowSpace `
                        -Parameters @($Arguments.InstanceName, $database, $edbVolumeFreeSpaceInGB, $edbPercentFreeSpace, $logVolumeFreeSpaceInGB, $logPercentFreeSpace, $PercentEdbFreeSpaceThreshold, $PercentLogFreeSpaceThreshold, $HourThreshold, $growthRateThreshold, $counterValue, $currentGrowthRate, $cUsersQuarantined, $PercentEdbFreeSpaceCriticalThreshold, $PercentEdbFreeSpaceAlertThreshold, $QuarantineValue)
                
                    return
                }
                
                write-verbose ("Quarantining: " + $topLogGenerators[$iNextUserToQuarantine].MailboxGuid)
                Set-QuarantineMailbox $topLogGenerators[$iNextUserToQuarantine].MailboxGuid
                
                Log-Event `
                    -Arguments $Arguments `
                    -EventInfo $StoreLogEntries.DatabaseSpaceTroubleShooterQuarantineUser `
                    -Parameters @($topLogGenerators[$iNextUserToQuarantine].MailboxGuid, $database)

                $currentGrowthRate -= $topLogGenerators[$iNextUserToQuarantine].TotalLogBytes
                $iNextUserToQuarantine++
                $cUsersQuarantined++
            }
        }

        Log-Event `
            -Arguments $Arguments `
            -EventInfo $StoreLogEntries.DatabaseSpaceTroubleShooterFinishedInsufficient `
            -Parameters @($Arguments.InstanceName, $database, $edbVolumeFreeSpaceInGB, $edbPercentFreeSpace, $logVolumeFreeSpaceInGB, $logPercentFreeSpace, $PercentEdbFreeSpaceThreshold, $PercentLogFreeSpaceThreshold, $HourThreshold, $growthRateThreshold, $counterValue, $currentGrowthRate, $cUsersQuarantined, $PercentEdbFreeSpaceCriticalThreshold, $PercentEdbFreeSpaceAlertThreshold, $QuarantineValue)
    }

###################################################################################################################################
#                                                                                                                                 #
#                                                     Script Body                                                                 #
#                                                                                                                                 #
###################################################################################################################################

    Set-StrictMode -Version Latest
        
    $scriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)

    . $scriptPath\CITSLibrary.ps1
    . $scriptPath\StoreTSLibrary.ps1
    . $scriptPath\StoreTSConstants.ps1
    . $scriptPath\DiagnosticScriptCommonLibrary.ps1

    Load-ExchangeSnapin
    
    $QuarantineValue = $false
    $MonitoringContextValue = $false
    $MailboxServer = $null
    $InvocationGuid = $null

    # Since we're in strict mode we must declare all variables we use
    $script:monitoringEvents = $null 

    if (!$MyInvocation.BoundParameters.ContainsKey("PercentEdbFreeSpaceThreshold"))
	{
        $PercentEdbFreeSpaceThreshold = $PercentEdbFreeSpaceDefaultThreshold
    }
    
    if (!$MyInvocation.BoundParameters.ContainsKey("PercentLogFreeSpaceThreshold"))
	{
        $PercentLogFreeSpaceThreshold = $PercentLogFreeSpaceDefaultThreshold
    }

    if (!$MyInvocation.BoundParameters.ContainsKey("HourThreshold"))
    {
        $HourThreshold = $HourDefaultThreshold
    }
    
    if ($Quarantine)
    {
        $QuarantineValue = $true
    }
    
    if ($MonitoringContext)
    {
        $MonitoringContextValue = $true
        
        #E14:303295 Add a monitoring event to suppress SCOM failure alerts
        #Use a guid for each invocation so that we can identify
        #start and finish events for each unique invocation
        $InvocationGuid = [System.Guid]::NewGuid().ToString()
        $messageStart = "Database Space TS started successfully for Invocation Guid {0}" -f $InvocationGuid
        Add-MonitoringEvent -Id $StoreLogEntries.DatabaseSpaceTroubleShooterStarted[0] -Type $EVENT_TYPE_INFORMATION -Message $messageStart
    }
    
    if ($PSCmdlet.ParameterSetName -eq "Server")
    {
        $MailboxServer = $Server
        $databases = @(Get-MailboxDatabase -Server $MailboxServer)
    }
    else
    {
        $databases = @(Get-MailboxDatabase $MailboxDatabaseName)
    }
    
    foreach($database in $databases)
    {
        #Get the status right before invoking the function
        #so we can narrow down failover related issues
        $database = Get-MailboxDatabase $database -Status
        
        #Run only if the database is specified or in case of server,
        #on the active copy to avoid running it multiple times
        #against the same database
        if (($MailboxServer -eq $null) -or 
            ($database.MountedOnServer -match $MailboxServer))
        {
            Troubleshoot-DatabaseCopySpace `
                -database $database `
                -PercentEdbFreeSpaceThreshold $PercentEdbFreeSpaceThreshold `
                -PercentLogFreeSpaceThreshold $PercentLogFreeSpaceThreshold `
                -HourThreshold $HourThreshold `
                -Quarantine $QuarantineValue `
                -MonitoringContext $MonitoringContextValue
        }
    }

    if ($MonitoringContext)
    {
        #Monitoring event to suppress SCOM failure alerts
        #Also lets us know that TS did finish successfully
        $messageFinish = "Database Space TS finished successfully for Invocation Guid {0}" -f $InvocationGuid
        Add-MonitoringEvent -Id 5102 -Type $EVENT_TYPE_INFORMATION -Message $messageFinish
        
        Write-MonitoringEvents
    }
# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUkEAK20YFYl1kWzuuvo4W7wpY
# SO2gghhqMIIE2jCCA8KgAwIBAgITMwAAASMwQ40kSDyg1wAAAAABIzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzQw
# WhcNMjAwMTEwMjEwNzQwWjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046MUE4Ri1FM0MzLUQ2OUQxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQCyYt3Mdjll12pYYKRUadMqDK0tJrlPK1MhEo75C/BI1y2i
# r4HnxDl/BSAvi44lI4IKUFDS40WJPlnCHfEuwgvEUNuFrseh7bDkaWezo6W/F7i9
# PO7GAxmM139zh5p2a1XwfzqYpZE2hEucIkiOR82fg7I0s8itzwl1jWQmdAI4XAZN
# LeeXXof9Qm80uuUfEn6x/pANst2N+WKLRLnCqWLR7o6ZKqofshoYFpPukVLPsvU/
# ik/ch1kj2Ja53Zb+KHctMCk/CpN2p7fNArpLcUA3H7/mdJjlaUFYLY9yy5TBndFF
# I1kBbZEB/Z1kYVnjRsIsV8W2CCp1RCxiIkx6AhIzAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQU2zl1LgtoHHcQXPImRhW0WL0hxPAwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAauzxVwcRLuLSItcW
# CHqZFtDSR5Ci4pgS+WrhLSmhfKXQRJekrOwZR7keC/bS7lyqai7y4NK9+ZHc2+F7
# dG3Ngym/92H45M/fRYtP63ObzWY9SNXBxEaZ1l8UP/hvv3uJnPf5/92mws50THX8
# tlvKAkBMWikcuA5y4s6yYy2GBFZIypm+ChZGtswTCst+uZhG8SBeE+U342Tbb3fG
# 5MLS+xuHrvSWdRqVHrWHpPKESBTStNPzR/dJ7pgtmF7RFKAWYLcEpPhr9hjUcf9q
# SJa7D5aghTY2UNFmn3BvKBSON+Dy5nDJA81RyZ/lU9iCOG+hGdpsGsJfvKT5WxsJ
# vEwdjzCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUTTUA0vq/7oznl9uyecCtVyW+xjow
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBADWRWQcm5v2lzAEoDqw9od+RAqmX1DG4Jfa5axB6idJn
# H/R21OFiRCCRW/2aZYBgkA7hHYhXrT3FZRGMr4ux4QYLWQ3yTkvO4VeLkhMj7KbD
# IkKGGbrc00mG9oMFEKU8xv9HGp17l7hFz70F5Cjq75ZRG1StiMKlJiijix8bcmMQ
# pqCj/CZdaWufpuPOA/vj+xE/+HNG/LzBiHfOe2bd6FpeO5vV0xZ+nmGlYKRBmJOt
# AVpcr4en6O5+miSQaTAX4ET0ChZjIZp4ZJdgfaj2NwnVcsQCvR7Qqv8EsWaeG2dF
# 4UCHWnk4sw2sXo4ZMcjEKps7Ftrt5plK44YTGziFG9qhggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# IzBDjSRIPKDXAAAAAAEjMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI1NTdaMCMGCSqGSIb3DQEJ
# BDEWBBSiLNrTaCxLkUYs7oZ4lxoFg22cVTANBgkqhkiG9w0BAQUFAASCAQB90ZRf
# 1je4O9AwYJg6ZcD4uBGbVfffL5oUG2UWQzZiby/v6fXy6nPhHPC/dME0poaBLzEW
# HBy2Uu3vSn5RzpkN3WyUU7s+XIuJGvwL/+kKW11VzCxk+tOVJEKV0vCp76eC6H62
# dGgNwwlOapGEsGmSDxIsA6J3efCZ9iZoyaCK1kVBn2fREUy1pPmoCHtRdeZ82lEn
# 41edlO+WuH1uY7RsdC2jaJHoyb+Da531P0OMXzM6SAtFzcXGkh6LrzUX+j09xiXF
# cunJkJIZVglZ24QpmrgR2oVY9Q95C8HyJlFz3RSt07M59miABzdR+YjLN7mm2rM5
# hnCbqf9LG6bNsdrf
# SIG # End signature block
