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
# MIIaggYJKoZIhvcNAQcCoIIaczCCGm8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUkEAK20YFYl1kWzuuvo4W7wpY
# SO2gghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# a8kTo/0xggS9MIIEuQIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIHWMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBRNNQDS+r/ujOeX27J5wK1XJb7GOjB2BgorBgEEAYI3AgEMMWgw
# ZqA+gDwAVAByAG8AdQBiAGwAZQBzAGgAbwBvAHQALQBEAGEAdABhAGIAYQBzAGUA
# UwBwAGEAYwBlAC4AcABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL2V4
# Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQANSt9kspMsNwWVwVrxjhfDBsesinZy
# 6MGbuB5a/YynqdMU/XL2aqzatgE3P+NajdgFjdw7Q61J5NYvlqBPbhNWMH/ss35z
# lQuwqoY89MnZsYqcCQP3Vy45MUbiD2RkjoTknTy4LM7PVqP+I0UIKRiIACRk8igp
# 1mLvi+JgLwew4ze8XTOBnqQuU0K5ER6D4EwbK3dO/1Ne2B/8g5+f2ZbMgC+jrUkX
# Y3ohYbJ0NWR72BxnwVSeWgZQFc6gqcLLuAFki3ZNkrDyR2kVFOnGSvu8XJzXHNsU
# d9N1e0A2A5YSnTgH4bZUg7bYMjHKZVm1huRgbuG7roqoRlSx+sdkOQH5oYICKDCC
# AiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQQITMwAAACs5MkjBsslI8wAAAAAAKzAJBgUrDgMCGgUAoF0wGAYJKoZIhvcN
# AQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTMwMjA2MDAxMDAxWjAj
# BgkqhkiG9w0BCQQxFgQUKFIe+A96TJEtGO/aBC7ERmV6VyYwDQYJKoZIhvcNAQEF
# BQAEggEABx4ic0Vh1oZ+7haQ+ohg66vP6uzPkI7/AfNh6Tw27zYUPIAe/p4ED+BK
# 08MGtHT2U3Kjoxaj43gpwruO8DWkNSXhNlV/fMmAF2tK6ZwrUs+ji4XM2nYaNbK5
# icPZPzQdooOVk50P9OAvqP8OYw9j0ZGlJ7+LKE5BLfPcxVUZzoNMp5Cp0hCGwhGR
# 6roBR0LZhP6yE8AmfPc2Jluy7FDuGDP2jytgToqIIoT/RCyrqc6ufwbCtpcZpJud
# dw3MbYJi+Y6aJrCj2VEfOIaTKwkQs/5fNqlvy/Iva3mckLgvo98fmw0v2MYCsD+N
# cfSEc4ukYy5GSXVhHQ6NlqRZnBp4aA==
# SIG # End signature block
