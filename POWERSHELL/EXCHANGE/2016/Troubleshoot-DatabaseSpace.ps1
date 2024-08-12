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
        $Arguments.InstanceName = $edbVolume.Name
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $StoreLogEntries.DatabaseSpaceTroubleShooterStarted `
            -Parameters @($edbVolume.Name, $database)
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
                -Parameters @($edbVolume.Name, $database, $PercentEdbFreeSpaceCriticalThreshold)
                
            $problemDetected = $true
        }
        elseif ($edbPercentFreeSpace -lt $PercentEdbFreeSpaceAlertThreshold)
        {
            # Log an event that will trigger non-paging alert for low space issue
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $StoreLogEntries.DatabaseSpaceTroubleDetectedAlertSpaceIssue `
                -Parameters @($edbVolume.Name, $database, $PercentEdbFreeSpaceAlertThreshold)
               
            $problemDetected = $true
        }
        elseif ($edbPercentFreeSpace -lt $PercentEdbFreeSpaceThreshold)
        {
            $edbGrowthRateThreshold = $edbVolumeFreeSpace / $HourThreshold
            
            # Log a warning event for low space issue
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $StoreLogEntries.DatabaseSpaceTroubleDetectedWarningSpaceIssue `
                -Parameters @($edbVolume.Name, $database, $PercentEdbFreeSpaceThreshold)
                
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
                    -Parameters @($edbVolume.Name, $database, $edbVolumeFreeSpaceInGB, $edbPercentFreeSpace, $logVolumeFreeSpaceInGB, $logPercentFreeSpace, $PercentEdbFreeSpaceThreshold, $PercentLogFreeSpaceThreshold, $HourThreshold, $counterValue, $PercentEdbFreeSpaceCriticalThreshold, $PercentEdbFreeSpaceAlertThreshold, $QuarantineValue)
                
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
                -Parameters @($edbVolume.Name, $database, $edbVolumeFreeSpaceInGB, $edbPercentFreeSpace, $logVolumeFreeSpaceInGB, $logPercentFreeSpace, $PercentEdbFreeSpaceThreshold, $PercentLogFreeSpaceThreshold, $HourThreshold, $growthRateThreshold, $counterValue, $currentGrowthRate, $PercentEdbFreeSpaceCriticalThreshold, $PercentEdbFreeSpaceAlertThreshold, $QuarantineValue)
                
            return
        }

        # We have to quarantine some users now, so
        # get the least of the most active users
        # and figure out which ones we need to quarantine
        # so that the current log generation rate drops
        # below our maximum calculated growth rate
        $topLogGenerators = @(Get-TopLogGenerators $database)
        $iNextUserToQuarantine = 0

        #If we are down here, let's save store usage statistics and send mail to on-calls
        #so that it can be analysed irrespective of whether we quarantine the mailbox or not.
        #We only want to do this in prod.
        if ($MonitoringContext -and $global:StoreUsageStatsLB)
        {
            $exchangeInstallPath = (get-item HKLM:\SOFTWARE\Microsoft\ExchangeServer\V14\Setup).GetValue("MsiInstallPath")
            $StoreLibraryPath = Join-Path $exchangeInstallPath "Datacenter\StoreCommonLibrary.ps1"
            if (Test-Path $StoreLibraryPath)
            {
                # Dot-source common libraries to populate alert details.
                . ($StoreLibraryPath)
                
                #Get store usage statistics and send mail
                #we are passing a null alertGuid since the space TS does not run in
                #response to an alert hence there is no alertGuid.
                $statsPath = Output-StoreUsageStatistics $global:StoreUsageStatsLB $null
                Send-DataCollectionMail $null $statsPath
            }
        }

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
                        -Parameters @($edbVolume.Name, $database, $edbVolumeFreeSpaceInGB, $edbPercentFreeSpace, $logVolumeFreeSpaceInGB, $logPercentFreeSpace, $PercentEdbFreeSpaceThreshold, $PercentLogFreeSpaceThreshold, $HourThreshold, $growthRateThreshold, $counterValue, $currentGrowthRate, $cUsersQuarantined, $PercentEdbFreeSpaceCriticalThreshold, $PercentEdbFreeSpaceAlertThreshold, $QuarantineValue)
                
                    return
                }
                
                #Quarantine the mailbox
                write-verbose ("Quarantining: " + $topLogGenerators[$iNextUserToQuarantine].MailboxGuid)
                Enable-MailboxQuarantine -Identity $topLogGenerators[$iNextUserToQuarantine].MailboxGuid.ToString() -Confirm:$false -ErrorAction Stop
                
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
            -Parameters @($edbVolume.Name, $database, $edbVolumeFreeSpaceInGB, $edbPercentFreeSpace, $logVolumeFreeSpaceInGB, $logPercentFreeSpace, $PercentEdbFreeSpaceThreshold, $PercentLogFreeSpaceThreshold, $HourThreshold, $growthRateThreshold, $counterValue, $currentGrowthRate, $cUsersQuarantined, $PercentEdbFreeSpaceCriticalThreshold, $PercentEdbFreeSpaceAlertThreshold, $QuarantineValue)
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
# MIIdvAYJKoZIhvcNAQcCoIIdrTCCHakCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUNXI/XtD1kNF5Gip8hvLtLiei
# RcygghhkMIIEwzCCA6ugAwIBAgITMwAAAKxjFufjRlWzHAAAAAAArDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwNTAzMTcxMzIz
# WhcNMTcwODAzMTcxMzIzWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkMwRjQtMzA4Ni1ERUY4MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAnyHdhNxySctX
# +G+LSGICEA1/VhPVm19x14FGBQCUqQ1ATOa8zP1ZGmU6JOUj8QLHm4SAwlKvosGL
# 8o03VcpCNsN+015jMXbhhP7wMTZpADTl5Ew876dSqgKRxEtuaHj4sJu3W1fhJ9Yq
# mwep+Vz5+jcUQV2IZLBw41mmWMaGLahpaLbul+XOZ7wi2+qfTrPVYpB3vhVMwapL
# EkM32hsOUfl+oZvuAfRwPBFxY/Gm0nZcTbB12jSr8QrBF7yf1e/3KSiqleci3GbS
# ZT896LOcr7bfm5nNX8fEWow6WZWBrI6LKPx9t3cey4tz0pAddX2N6LASt3Q0Hg7N
# /zsgOYvrlwIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFCFXLAHtg1Boad3BTWmrjatP
# lDdiMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAEY2iloCmeBNdm4IPV1pQi7f4EsNmotUMen5D8Dg4rOLE9Jk
# d0lNOL5chmWK+d9BLG5SqsP0R/gqph4hHFZM4LVHUrSxQcQLWBEifrM2BeN0G6Yp
# RiGB7nnQqq86+NwX91pLhJ5LBzJo+EucWFKFmEBXLMBL85fyCusCk0RowdHpqh5s
# 3zhkMgjFX+cXWzJXULfGfEPvCXDKIgxsc5kUalYie/mkCKbpWXEW6gN+FNPKTbvj
# HcCxtcf9mVeqlA5joTFe+JbMygtOTeX0Mlf4rTvCrf3kA0zsRJL/y5JdihdxSP8n
# KX5H0Q2CWmDDY+xvbx9tLeqs/bETpaMz7K//Af4wggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBMIwggS+AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB1jAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUv67mOkL+aiLGqozlNB/aGt5p6Z0wdgYKKwYB
# BAGCNwIBDDFoMGagPoA8AFQAcgBvAHUAYgBsAGUAcwBoAG8AbwB0AC0ARABhAHQA
# YQBiAGEAcwBlAFMAcABhAGMAZQAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAYLkm4clxmCsee5I4
# Bf+4BaSAWCEj6zUqM4KTVZeFpx5NHXRNNcJec5v1TpIrIENn0pxrT82vuVDRLwBl
# 2iRRnCMwhefafwN/r2UpyCKXKtM8MfMX9NFSp5KndsSaaBlxfrU0Iz2rGmWjEdwO
# LWNCqezU7Zx9rQctuY00dvxy62JS8+Hkiei9c4u8FubW1dbYhWcPGpeWIPAppfdA
# L+mXiiJsoBXbJSkIW/1QpFYyIk/sJ18SZHWw9HwithubJyfPMzkG6+Nlr5j4Uywd
# 2NnaBroLld1jOKEPbitDuVRW+dxzkEutic7lJYbwsetyAmrqwGrnRFvgcCvURTyu
# Z2pDiaGCAigwggIkBgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0ECEzMAAACsYxbn40ZVsxwAAAAAAKwwCQYFKw4DAhoFAKBd
# MBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE2MDkw
# MzE4NDQ1NVowIwYJKoZIhvcNAQkEMRYEFB2gsHxxRfUYSiMeGRG4tJ3CheapMA0G
# CSqGSIb3DQEBBQUABIIBAFXPNeK8omwdfstK8FqYt/g9QePDqq9WMmG46VYDA8Id
# VoDPkjxVkLOWwghqLHNkoLCaZTHuGFmiHkWyoj1MOh8YvKWlS6oVqrCxTW9a+Nu/
# cljAJP0mBRCFeQgzckvef+JY5TsL9+0/i8iGge3cPPoEtACWJ5j6CGGqIsSaDPHg
# dceCk/g8tdqBNtlezuWVZoPqcb1gLleIA8O8wraVfrpRbFbKvu3KdwXy1KXGJ/LE
# Nj/Bt7nEOrvOH/CKlwYNlWdVzJx5xZ6ZViEfXzVx+HHm/ug/07m7UWM9GvMiP4B8
# i0K1rDewCeT+fKQtWsnkIJsVxbikAo6sJZgcK4mBqbY=
# SIG # End signature block
