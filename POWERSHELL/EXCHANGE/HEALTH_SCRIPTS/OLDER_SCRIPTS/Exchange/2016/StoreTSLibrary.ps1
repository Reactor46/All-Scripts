# Copyright (c) 2010 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

# Requires -Version 2

# Figures out on which volume our path is residing. This is necessary
# because customers mount spindles into a folder.
function Get-WmiVolumeFromPath([string] $FilePath, [string] $Server)
{
    do
    {
        $FilePath = $FilePath.Substring(0, $FilePath.LastIndexOf('\') + 1)

        $wmiFilter = ('Name="{0}"' -f $FilePath.Replace("\", "\\"))

        $volume = get-wmiobject -class win32_volume -computername $Server -filter $wmiFilter

        $FilePath = $FilePath.Substring(0, $FilePath.LastIndexOf('\'))

    } while ($volume -eq $null)

    return $volume
}

# Returns a descending list of the users generating the most log bytes for a given database
# based on the output of Get-StoreUsageStatistics the list contains the MailboxGuid and the
# number of bytes generated during the captured sampling periods (~ 1 hour)
function Get-TopLogGenerators([string] $DatabaseIdentity)
{
    # The Filter parameter doesn't accept complex filters, so filter on the category
    # and then use a where clause to filter out the Quarantined mailboxes.
    $topGenerators = New-Object Collections.ArrayList
    
    $global:StoreUsageStatsLB = Get-StoreUsageStatistics -Database $DatabaseIdentity `
                    -Filter "DigestCategory -eq 'LogBytes'"
         
    if ($global:StoreUsageStatsLB)
    {                               
        $stats = $global:StoreUsageStatsLB | where {$_.IsQuarantined -eq $false} `
                        | group MailboxGuid

        if ($null -ne $stats)
        {
            foreach($mailboxStats in $stats)
            {
                $total = 0
                $statSummary = new-object PSObject

                foreach($stat in $mailboxStats.Group)
                {
                    $total += $stat.LogRecordBytes
                }

                Add-Member -in $statSummary -Name TotalLogBytes -MemberType NoteProperty -Value $total
                Add-Member -in $statSummary -Name MailboxGuid -MemberType NoteProperty -Value $mailboxStats.Group[0].MailboxGuid

                # Tell PS we don't care about the return value for this function
                # otherwise these values will be output to the pipeline!
                [void]$topGenerators.Add($statSummary)
            }
            $topGenerators = Sort-Object -InputObject $topGenerators -Property TotalLogBytes -Descending
        }
    }

    return $topGenerators
}

# Returns a descending list of the users using up the most time in server for a given database
# based on the output of Get-StoreUsageStatistics the list contains the MailboxGuid and the
# time in server used up during the captured sampling periods (10 min)
function Get-TopCpuUsers([string] $DatabaseIdentity, [int] $TimeInServerThreshold, [int] $RopLatencyThreshold, [int] $PercentSampleBelowThresholdToAlert, [int] $MinimumStoreUsageStatisticsSampleCount)
{
    $topUsers = New-Object Collections.ArrayList

    $global:StoreUsageStatsTS = Get-StoreUsageStatistics -Database $DatabaseIdentity `
                    -Filter "DigestCategory -eq 'TimeInServer'"

    if ($global:StoreUsageStatsTS)
    {
        # First let's see if most users are being affected by high latency
        # or if it's just a small percent of users
        $totalCountAboveThreshold = 0
        $totalPctBelowThreshold = 0
        $totalCountAboveThreshold = ($global:StoreUsageStatsTS | ? {$_.TimeInServer/$_.RopCount -ge $RopLatencyThreshold} | measure).count; 
        $totalPercentBelowThreshold = [Math]::Round(100*($global:StoreUsageStatsTS.count - $totalCountAboveThreshold)/($global:StoreUsageStatsTS.count),1);
        
        # This is a condition that requires further investigation - If we have enough samples
        # and if users above the totalPercentBelowThreshold are having a bad experience
        If ($global:StoreUsageStatsTS.count -ge $MinimumStoreUsageStatisticsSampleCount -and $totalPercentBelowThreshold -lt $PercentSampleBelowThresholdToAlert)
        {
            $stats = $global:StoreUsageStatsTS | where {$_.IsQuarantined -eq $false} `
                        | group MailboxGuid    
                    
            if ($null -ne $stats)
            {
                foreach($mailboxStats in $stats)
                {
                    $statSummary = new-object PSObject
                    $totalTimeInServer = 0
                    $totalROPCount = 0
                    
                    # Get Total TimeInServer and ROPCount
                    foreach($stat in $mailboxStats.Group)
                    {
                        $totalTimeInServer += $stat.TimeInServer
                        $totalROPCount += $stat.ROPCount
                    }
                
                    # Calculate averages
                    $averageTimeInServer = $totalTimeInServer / $mailboxStats.Count
                    $averageROPCount = $totalROPCount / $mailboxStats.Count
                    $ropLatency = $averageTimeInServer / $averageROPCount
                    
                    # If either TimeInServer or RopLatency is above the threshold, include that mailbox in the list
                    if (($averageTimeInServer -ge $TimeInServerThreshold) -or ($ropLatency -ge $RopLatencyThreshold))
                    {
                        Add-Member -in $statSummary -Name TotalTimeInServer -MemberType NoteProperty -Value $totalTimeInServer
                        Add-Member -in $statSummary -Name AverageTimeInServer -MemberType NoteProperty -Value $averageTimeInServer
                        Add-Member -in $statSummary -Name RopLatency -MemberType NoteProperty -Value $ropLatency
                        Add-Member -in $statSummary -Name MailboxGuid -MemberType NoteProperty -Value $mailboxStats.Group[0].MailboxGuid
                        Add-Member -in $statSummary -Name Count -MemberType NoteProperty -Value $mailboxStats.Count

                        # Tell PS we don't care about the return value for this function
                        # otherwise these values will be output to the pipeline!
                        [void]$topUsers.Add($statSummary)
                    }
                }
                $topUsers = Sort-Object -InputObject $topUsers -Property TotalTimeInServer -Descending
            }
        }        
    }

    return $topUsers
}

function Get-RegistryValue
{
	param(
		[string]$Server = ".",
		[Parameter(Mandatory=$true)][Microsoft.Win32.RegistryHive]$Hive,
		[Parameter(Mandatory=$true)][string]$KeyName,
		[Parameter(Mandatory=$true)][string]$ValueName,
		[object]$DefaultValue = $null
	)

	try
	{
	    $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Server)
	    $key = $baseKey.OpenSubKey($KeyName, $false)
	    if ($key -ne $null)
	    {
    		$key.GetValue($ValueName, $DefaultValue)
	    }
    }
    finally
    {
        if ($key -ne $null)
        {
            $key.Close()
        }
	    if ($baseKey -ne $null)
	    {
    	    $baseKey.Close()
	    }
	}
}

function Set-RegistryValue
{
	param(
		[string]$Server = ".",
		[Parameter(Mandatory=$true)][Microsoft.Win32.RegistryHive]$Hive,
		[Parameter(Mandatory=$true)][string]$KeyName,
		[Parameter(Mandatory=$true)][string]$ValueName,
		[Parameter(Mandatory=$true)][object]$Value,
		[Microsoft.Win32.RegistryValueKind]$ValueKind = "Unknown"
	)

    try
    {
	    $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Server)
	    $key = $baseKey.OpenSubKey($KeyName, $true)
	    if ($key -eq $null)
	    {
    		$key = $baseKey.CreateSubKey($KeyName)
    	}

    	$key.SetValue($ValueName, $Value, $ValueKind)
	}
    finally
    {
        if ($key -ne $null)
        {
            $key.Close()
        }
	    if ($baseKey -ne $null)
	    {
    	    $baseKey.Close()
	    }
	}
}

function Remove-RegistryValue
{
	param(
		[string]$Server = ".",
		[Parameter(Mandatory=$true)][Microsoft.Win32.RegistryHive]$Hive,
		[Parameter(Mandatory=$true)][string]$KeyName,
		[Parameter(Mandatory=$true)][string]$ValueName
	)

    try
    {
	    $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Server)
	    $key = $baseKey.OpenSubKey($KeyName, $true)
	    if ($key -ne $null)
	    {
    		$key.DeleteValue($ValueName, $false)
    	}
    }
    finally
    {
        if ($key -ne $null)
        {
            $key.Close()
        }
	    if ($baseKey -ne $null)
	    {
    	    $baseKey.Close()
	    }
	}
}

function Remove-RegistryKey
{
	param(
		[string]$Server = ".",
		[Parameter(Mandatory=$true)][Microsoft.Win32.RegistryHive]$Hive,
		[Parameter(Mandatory=$true)][string]$KeyName
	)

    try
    {
	    $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Server)
	    $baseKey.DeleteSubKeyTree($KeyName)
	}
    finally
    {
        if ($key -ne $null)
        {
            $key.Close()
        }
	    if ($baseKey -ne $null)
	    {
    	    $baseKey.Close()
	    }
	}
}

function Get-RegistrySubKeyNames
{
	param(
		[string]$Server = ".",
		[Parameter(Mandatory=$true)][Microsoft.Win32.RegistryHive]$Hive,
		[Parameter(Mandatory=$true)][string]$KeyName
	)

    try
    {
	    $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Server)
	    $key = $baseKey.OpenSubKey($KeyName, $false)
    	if ($key -ne $null)
    	{
		    $key.GetSubKeyNames()
	    }
	}
    finally
    {
        if ($key -ne $null)
        {
            $key.Close()
        }
	    if ($baseKey -ne $null)
	    {
    	    $baseKey.Close()
	    }
	}
}

function Get-AverageResultForCounter($results, $counter, $instance)
{
    #Turn strictmode off so that if we don't have cookedvalue
    #for a particular instance, we don't get an exception. 
    #This way, we can get average from whatever counters are 
    #available and continue as far as we can. Turning strict mode
    #off is only within the scope of this function.
    Set-StrictMode -off
    
    $count = 0
    $total = $null
    $counterName = $counter
    
    if ($instance)
    {
        $counterName = $counterName.replace("*", $instance)
    }

    foreach ($sample in $results)
    {
        $cookedValue = ($sample.CounterSamples | ?{$_.Path -like "*$counterName*"}).CookedValue
        
        if ($cookedValue -ne $null)
        {
            #We want to increment the number of samples here so that
            #we can calculate the avg properly rather than assuming
            #that we got samples equal to the length of results
            $total += $cookedValue
            $count++
        }
    }
    
    If ($total -ne $null -and $count -ne 0)
    {
    	return $total/($count)
    }
    else
    {
        $failureReason = ($StoreLocStrings.FailureToGetCounter + $counterName)
        if ($Arguments -ne $null)
        {
            #Get the invoking scriptname so we can put in the event
            $script = $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf('\') + 1)
            if ($script -ne $null)
            {
                $script = $script.Substring(0,$script.Length - 4)
            }
            else
            {
                $script = "Troubleshooter"
            }
            
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $StoreLogEntries.DatabaseTroubleShooterNotRun `
                -Parameters @($script, $failureReason)
        }
        
        throw $failureReason
    }
}

Import-LocalizedData -BindingVariable StoreLocStrings -FileName StoreTSLibrary.strings.psd1

$StoreLogEntries = @{
#
# Events logged to application log and windows event (crimson) log
# Information: 5100-5199; Warning: 5400-5499; Error: 5700-5799;
#
#   Informational events
#
    DatabaseSpaceTroubleShooterStarted=(5100,"Information", $StoreLocStrings.DatabaseSpaceTroubleShooterStarted)
    DatabaseSpaceTroubleShooterNoProblemDetected=(5101,"Information", $StoreLocStrings.DatabaseSpaceTroubleShooterNoProblemDetected)
    DatabaseLatencyTroubleShooterStarted=(5110,"Information", $StoreLocStrings.DatabaseLatencyTroubleShooterStarted)
    DatabaseLatencyTroubleShooterNoLatency=(5111,"Information", $StoreLocStrings.DatabaseLatencyTroubleShooterNoLatency)
    DatabaseLatencyTroubleShooterLowOps=(5112,"Information", $StoreLocStrings.DatabaseLatencyTroubleShooterLowOps)
    DatabaseTroubleShooterNotRun=(5113,"Information", $StoreLocStrings.DatabaseTroubleShooterNotRun)
    DatabaseLatencyTroubleShooterNotRunDatabase=(5114,"Information", $StoreLocStrings.DatabaseLatencyTroubleShooterNotRunDatabase)
    DatabaseLatencyTroubleShooterUniqueMailbox=(5115,"Information", $StoreLocStrings.DatabaseLatencyTroubleShooterUniqueMailbox)
    DatabaseLatencyTroubleShooterNoMailbox=(5116,"Information", $StoreLocStrings.DatabaseLatencyTroubleShooterNoMailbox)

    DatabaseSpaceTroubleShooterFoundLowSpace=(5400,"Warning", $StoreLocStrings.DatabaseSpaceTroubleShooterFoundLowSpace)
    DatabaseSpaceTroubleShooterFoundLowSpaceNoQuarantine=(5401,"Warning", $StoreLocStrings.DatabaseSpaceTroubleShooterFoundLowSpaceNoQuarantine)
    DatabaseSpaceTroubleDetectedWarningSpaceIssue=(5402,"Warning", $StoreLocStrings.DatabaseSpaceTroubleDetectedWarningSpaceIssue)
    DatabaseSpaceTroubleShooterQuarantineUser=(5410,"Warning", $StoreLocStrings.DatabaseSpaceTroubleShooterQuarantineUser)
    DatabaseLatencyTroubleShooterQuarantineUser=(5411,"Warning", $StoreLocStrings.DatabaseLatencyTroubleShooterQuarantineUser)
    DatabaseLatencyTroubleShooterNoQuarantine=(5412,"Warning", $StoreLocStrings.DatabaseLatencyTroubleShooterNoQuarantine)

    DatabaseSpaceTroubleShooterFinishedInsufficient=(5700,"Error", $StoreLocStrings.DatabaseSpaceTroubleShooterFinishedInsufficient)
    DatabaseSpaceTroubleDetectedAlertSpaceIssue=(5701,"Error", $StoreLocStrings.DatabaseSpaceTroubleDetectedAlertSpaceIssue)
    DatabaseSpaceTroubleDetectedCriticalSpaceIssue=(5702,"Error", $StoreLocStrings.DatabaseSpaceTroubleDetectedCriticalSpaceIssue)
    DatabaseLatencyTroubleShooterBadDiskLatencies=(5710,"Error", $StoreLocStrings.DatabaseLatencyTroubleShooterBadDiskLatencies)
    DatabaseLatencyTroubleShooterBadDSAccessActiveCallCount=(5711,"Error", $StoreLocStrings.DatabaseLatencyTroubleShooterBadDSAccessActiveCallCount)
    DatabaseLatencyTroubleShooterIneffective=(5712,"Error", $StoreLocStrings.DatabaseLatencyTroubleShooterIneffective)
    DatabaseLatencyTroubleShooterBadDSAccessAverageLatency=(5713,"Error", $StoreLocStrings.DatabaseLatencyTroubleShooterBadDSAccessAverageLatency)
}

# SIG # Begin signature block
# MIIdpAYJKoZIhvcNAQcCoIIdlTCCHZECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUZBLU0SgWPq6EnyVMP0qu1JGk
# VlqgghhkMIIEwzCCA6ugAwIBAgITMwAAAJvgdDfLPU2NLgAAAAAAmzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI5
# WhcNMTcwNjMwMTkyMTI5WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjcyOEQtQzQ1Ri1GOUVCMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAjaPiz4GL18u/
# A6Jg9jtt4tQYsDcF1Y02nA5zzk1/ohCyfEN7LBhXvKynpoZ9eaG13jJm+Y78IM2r
# c3fPd51vYJxrePPFram9W0wrVapSgEFDQWaZpfAwaIa6DyFyH8N1P5J2wQDXmSyo
# WT/BYpFtCfbO0yK6LQCfZstT0cpWOlhMIbKFo5hljMeJSkVYe6tTQJ+MarIFxf4e
# 4v8Koaii28shjXyVMN4xF4oN6V/MQnDKpBUUboQPwsL9bAJMk7FMts627OK1zZoa
# EPVI5VcQd+qB3V+EQjJwRMnKvLD790g52GB1Sa2zv2h0LpQOHL7BcHJ0EA7M22tQ
# HzHqNPpsPQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFJaVsZ4TU7pYIUY04nzHOUps
# IPB3MB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBACEds1PpO0aBofoqE+NaICS6dqU7tnfIkXIE1ur+0psiL5MI
# orBu7wKluVZe/WX2jRJ96ifeP6C4LjMy15ZaP8N0OckPqba62v4QaM+I/Y8g3rKx
# 1l0okye3wgekRyVlu1LVcU0paegLUMeMlZagXqw3OQLVXvNUKHlx2xfDQ/zNaiv5
# DzlARHwsaMjSgeiZIqsgVubk7ySGm2ZWTjvi7rhk9+WfynUK7nyWn1nhrKC31mm9
# QibS9aWHUgHsKX77BbTm2Jd8E4BxNV+TJufkX3SVcXwDjbUfdfWitmE97sRsiV5k
# BH8pS2zUSOpKSkzngm61Or9XJhHIeIDVgM0Ou2QwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBKowggSmAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBvjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUnCpd7AnNArbaE3Xrni8AVdjtHJMwXgYKKwYB
# BAGCNwIBDDFQME6gJoAkAFMAdABvAHIAZQBUAFMATABpAGIAcgBhAHIAeQAuAHAA
# cwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZI
# hvcNAQEBBQAEggEAQBi0gH8roSLLwPQQW8jbUR16bP5P6h2iYj7MCgFSkcleoLOu
# VGGz+wf8p8qH+NL4fbWhDy62bv1IpaIk707Y13EubikdIFbmTBRjcaPKiLaUq1kz
# ZV9IIEHna5fNKV3aFSr1w9RRiYv/BWcsEKyZxlRTYYM2hWg5wk97ohclFUCdHp+H
# ojMWUh0vntIhxuzV00AJdjpIKZrKNO4Q7zfJXcR88dVfwS4UWHIsC11izBLLIYep
# kD0cb8cIUHrhNURkxQG6v8jSPzXt9T+OXKnOF5Au/zVJRXYsq7z7s5IQdXAHQ2Es
# 0r6bATTiF/5jN2sF6PUFrXDBnctmGD6hTx1hl6GCAigwggIkBgkqhkiG9w0BCQYx
# ggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAACb4HQ3
# yz1NjS4AAAAAAJswCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0B
# BwEwHAYJKoZIhvcNAQkFMQ8XDTE2MDkwMzE4NDQ1NlowIwYJKoZIhvcNAQkEMRYE
# FPcoIzSWRQYNuRTluTupcfwC5bEwMA0GCSqGSIb3DQEBBQUABIIBAIbmv0rOb4sw
# zjMwCRkWJEB46vEnd1Z3f+bPrfmmvtIbo6F8kLn7LVXM0aaXpSOvv8Uk0k3LAeXL
# duSXkCS/kkAXRQGeAjFweCdjFbTXJ6uVVeKrBi1nE0dGIrcU2aKxB55mKzpE+b1S
# Tjiqf0D97bcz9MDTHpDdBOAjFEMMrRKVTqz5QQ1pWSqZMWCCh1eHuyUViadAH1+g
# nI7ofPjP0FeUMN9qltIdq4GFzyUXQFdSz3f0N/Qw6XKP0wnK5ewYZHdDmnbsUyuC
# 9qjTt8TG1dxQxqX6RK3biNbpRglv40G3O7gXDeZRqfasZ3Cazd14ITRz9KNOhTQq
# yMAElKl5vE4=
# SIG # End signature block
