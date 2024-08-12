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
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUzMzeKL0udnuGpXpII9+XDJZ4
# WWagghhqMIIE2jCCA8KgAwIBAgITMwAAARzbbpm3tnP6bwAAAAABHDANBgkqhkiG
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU3gTerIbjp4fMYUPehk0nat6ZOpMw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAG+UpLWTfzSWdRhQljXIWjiQGfT/HfxO3VHBdni5ZO1I
# L1uJxIxC1DgOzzuPhmACqWFt+CbvGk+dQKOyNrF/EKKK7ioghQc+WGVmO8lK3GQF
# 4fJu3GvqOPv9stwE8EW+w9fJCl1ZbckdO2WORhxurIU713yh+vUgihO7CKfVLoj3
# c6rAsAFEKvV7RL0BrkDG3Cha9ZahV/+5VW7wmIf3QJUn/L+E28Oq6v7s/B4PdJ/m
# jtASuFtgQ9h5wpHzgpzCwzSAzBqGSf0Hgz32fQ0MFR+pwB4KXumMvZomzIYdeuNC
# TIV+i0ytx+LaKwi8gfwz/IKAweruvOMiTYElLpmqEJOhggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# HNtumbe2c/pvAAAAAAEcMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI1NTlaMCMGCSqGSIb3DQEJ
# BDEWBBR9ZxDwgw6fYH7iA9ifLBwDuE3FbDANBgkqhkiG9w0BAQUFAASCAQBgzYTm
# HXcF7RZmUvyqAOT8NoFL1qD6cYV3vOSTBYBxYlRdfRAMNgLYzcjE+LPnsDKJK+U1
# JvYjwpCsPfsGfBDNTh+yolAyf80pieP95TWZmGyn0pscTDdu4dAYD81ArrjOcfSq
# Oz4Z9iTnD5UP2+819tFshwCD2E2ozqric0FCeBUc3uJKwDsB3Wkhe04Mk/glF58j
# /mYvLrwwEY3gulZTIPGGU+2GlkPxqvirJW33m82xWxFojRI5Rysi2K6q+xiJwF0+
# DqMNXx0XPwuqF2HrFAOBf/uxmvCiH5an9EEBre3+rZOfpx13fkG56NDSEU2G8rCp
# mYmcj1MUQLmzKtg1
# SIG # End signature block
