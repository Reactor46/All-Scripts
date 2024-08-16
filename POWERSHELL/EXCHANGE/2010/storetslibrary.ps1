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
    $stats = Get-StoreUsageStatistics -Database $DatabaseIdentity `
                    -Filter "DigestCategory -eq 'LogBytes'" `
                    | where {$_.IsQuarantined -eq $false} `
                    | group MailboxGuid

    $topGenerators = New-Object Collections.ArrayList

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

    return $topGenerators
}

# Returns a descending list of the users using up the most time in server for a given database
# based on the output of Get-StoreUsageStatistics the list contains the MailboxGuid and the
# time in server used up during the captured sampling periods (10 min)
function Get-TopCpuUsers([string] $DatabaseIdentity, [int] $TimeInServerThreshold)
{
    # StoreUsageStatistics return 10 samples of 1 minute each for TimeInServer
    $samples = 10
    $stats = Get-StoreUsageStatistics -Database $DatabaseIdentity `
                    -Filter "DigestCategory -eq 'TimeInServer'" `
                    | where {$_.IsQuarantined -eq $false} `
                    | group MailboxGuid

    $topUsers = New-Object Collections.ArrayList

    if ($null -ne $stats)
    {
        foreach($mailboxStats in $stats)
        {
            $statSummary = new-object PSObject

            $total = 0

            foreach($stat in $mailboxStats.Group)
            {
                $total += $stat.TimeInServer
            }

            $average = $total / $samples

            if ($average -ge $TimeInServerThreshold)
            {
                Add-Member -in $statSummary -Name TotalTimeInServer -MemberType NoteProperty -Value $total
                Add-Member -in $statSummary -Name AverageTimeInServer -MemberType NoteProperty -Value $average
                Add-Member -in $statSummary -Name MailboxGuid -MemberType NoteProperty -Value $mailboxStats.Group[0].MailboxGuid

                # Tell PS we don't care about the return value for this function
                # otherwise these values will be output to the pipeline!
                [void]$topUsers.Add($statSummary)
            }
        }
        $topUsers = Sort-Object -InputObject $topUsers -Property TotalTimeInServer -Descending
    }

    return $topUsers
}

function Set-QuarantineMailbox([string] $Identity)
{
    $mbx = @(Get-Mailbox -Identity $Identity -ErrorAction Stop)
    if ($mbx.Count -ne 1)
    {
        $error = "Get-Mailbox returned more than one Mailbox for Identity " + $Identity + ". Please specify a unique Identity."
        throw $error
    }

    $db = Get-MailboxDatabase -Identity $mbx[0].Database -Status -ErrorAction Stop
    $srv = Get-ExchangeServer -Identity $db.MountedOnServer -ErrorAction Stop

    $databaseKey = [System.String]::Format(
            "SYSTEM\CurrentControlSet\Services\MSExchangeIS\{0}\Private-{1}",
            $srv.Name,
            $db.Guid)
            $quarantinedMailboxKey = [System.String]::Format(
            "{0}\QuarantinedMailboxes\{1}",
            $databaseKey,
            $mbx[0].ExchangeGuid)

    #Read the crash threshold so we can set the right value to quarantine
    $crashThreshold = Get-RegistryValue -Server $srv.Fqdn -Hive "LocalMachine" -KeyName $databaseKey -ValueName "MailboxQuarantineCrashThreshold"
    if ($crashThreshold -eq $null)
    {
        $crashThreshold = 3
    }

    $quarantineDuration = Get-RegistryValue -Server $srv.Fqdn -Hive "LocalMachine" -KeyName $databaseKey -ValueName "MailboxQuarantineDurationInSeconds"
    if ($quarantineDuration -eq $null)
    {
        $quarantineDuration = 21600
    }

    Set-RegistryValue -Server $srv.Fqdn -Hive "LocalMachine" -KeyName $quarantinedMailboxKey -ValueName "CrashCount" -Value $crashThreshold -ValueKind "DWord"
    Set-RegistryValue -Server $srv.Fqdn -Hive "LocalMachine" -KeyName $quarantinedMailboxKey -ValueName "LastCrashTime" -Value (Get-Date).ToFileTime() -ValueKind "QWord"
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
    $total = 0
    $counterName = $counter
    if ($instance)
    {
        $counterName = $counterName.replace("*", $instance)
    }

    foreach ($sample in $results)
    {
        $total += ($sample.CounterSamples | ?{$_.Path -like "*$counterName*"}).CookedValue
    }
    return $total/($results.length)
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
# MIIaagYJKoZIhvcNAQcCoIIaWzCCGlcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUzMzeKL0udnuGpXpII9+XDJZ4
# WWagghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# a8kTo/0xggSlMIIEoQIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIG+MBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBTeBN6shuOnh8xhQ96GTSdq3pk6kzBeBgorBgEEAYI3AgEMMVAw
# TqAmgCQAUwB0AG8AcgBlAFQAUwBMAGkAYgByAGEAcgB5AC4AcABzADGhJIAiaHR0
# cDovL3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASC
# AQAdz5WUAEb1kXQHD1zuryLq41VSUUaGvKGFCtWBtabLOuOfIQcpdGYjGcj4lo3z
# X/F5zkozDRSAsVMQbiIkMvT0oebnMsA5K+mA9ZCXh6cv+gwse/nhsmbhcwlrWkEw
# zasWI3ncNDGaBvwIUr0M+rITwv+hM7kFD7aof6m4Anerqjv35mtBTKzD8by8nVTn
# 9N0XRv7GW1NisMMagsXH3Rb1qCXlaoOFfkDWi/XDIJo19GSSnPmu2JdULEhTV9Br
# vFzMLpbXosW+5Lu/y188P/2KtLRFnlc3QMFnki+qeJOBwWzafdb2F7nuGpyb4AqJ
# c83pO8Zean7YT+0WK6S1uMTkoYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEB
# MIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNV
# BAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAACs5MkjBsslI8wAAAAAA
# KzAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTMwMjA2MDAxMDAxWjAjBgkqhkiG9w0BCQQxFgQUiLZUspmukLYX
# KnHHf5J95qskQwQwDQYJKoZIhvcNAQEFBQAEggEADrOYVhkXDNOBjPLIQb7CkMQz
# qumNLwTM55TjfpACeitNxQS0eUS/qCXL+4BLcTbOsBhXcW69OxvSmOTThSXHG2Hp
# 2Q43hCBoOL3G9VCb+Wa+c2pU2s3IVzIPZAzZTw3nF7UAbYPgSKyXGCzoI4LwxwKA
# YTE5QcTDCUzeXi8Z5aLUCfwD6Ot2MeXZPOHITu7Ir7Ry35kFU+PP4Dn0oR+rV+tv
# slR3EUQVjbbrKEWAsPCq/YJObOd3oUPfC9bS/qSBmQQIziGqQcow2Qu4oBFLolg5
# UKVjOC+kABkqGcrYTWt0JW2hrDMvGiaczIHzPP5yqEJS8wYmz6kAEIMshzf7VQ==
# SIG # End signature block
