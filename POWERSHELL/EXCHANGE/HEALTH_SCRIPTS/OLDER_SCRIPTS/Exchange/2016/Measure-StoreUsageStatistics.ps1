param (
    [parameter(Mandatory=$true,
              ParameterSetName="ServerSet",
              HelpMessage="Retrieve statistics from server specified")]
    [String]
    $Server, 
    
    [parameter(ParameterSetName="ServerSet",
              HelpMessage="Retrieve statistics server hosting copy of database specified")]
    [String]
    $AlertGuid, 

    [parameter(Mandatory=$true,
              ParameterSetName="DatabaseSet",
              HelpMessage="Retrieve statistics from database specified")]
    [String]
    $Database,     
    
    [parameter(ParameterSetName="FileSet",
              HelpMessage="Retrieve events from all servers hosting copy of database specified")]
    [String]
    $FileName, 
    
    [parameter(HelpMessage="Measure statistics using sample provided (debug)")]
    $Stats, 
    
    [parameter(HelpMessage="Export statistics to global variable")]
    [Switch]
    $Export, 
    
    [parameter(HelpMessage="RPC Opeation latency threshold to count percentile of healthy operations")]
    [Int]
    $RopLatencyThreshold=15, 
    
    [parameter(HelpMessage="Return Top N users with IncludeDetailedAnalysis option")]
    [Int]
    $TopUsers=20, 
    
    [parameter(HelpMessage="Minimum sample count to be considered for alerting")]
    [Int]
    $MinimumSampleCount=125, 
    
    [parameter(HelpMessage="Threshold used for alerting")]
    [Int]
    $PctSampleBelowThresholdToAlert=80, 
    
    [parameter(HelpMessage="Return detailed Top N users in several pivots")]
    [Switch]
    $IncludeDetailedAnalysis, 
    
    [parameter(HelpMessage="Use DigestCategory to change statistics type to analyze (LogBytes useful for bytes generated)")]
    [ValidateSet("TimeInServer", "LogBytes")]
    [String]
    $DigestCategory="TimeInServer",

    [parameter(HelpMessage="Folder path to working directory to download StoreUsageStatistics file")]
    [String]
    $WorkingPath="$Home\AppData\Local\Temp"
)

function Get-TraceFileList
{
    param(
        [Parameter(Mandatory=$true)]
        $ExServer,

        [Parameter(Mandatory=$true)]
        [bool]
        $UseTorus
    )

    if ($UseTorus)
    {
        write-host "$(get-date) Looking for trace files using Get-MachineLog"
        Get-MachineLog -Target $ExServer.Name -Log StoreUsageStatistics
    }
    else
    {
        If (test-path "\\$($ExServer.FQDN)\Exchange\Diagnostics\StoreUsageStatistics")
        {
            Get-ChildItem -Path "\\$($ExServer.FQDN)\Exchange\Diagnostics\StoreUsageStatistics"
        }
        Else
        {
            write-host -foregroundcolor:red "path not found (\\$($ExServer.FQDN)\Exchange\Diagnostics\StoreUsageStatistics)"  
            return          
        }
    }
}

$useTorus = $false
$torus = Get-Variable TorusConnection -ErrorAction SilentlyContinue
if ($torus -ne $null)
{
    if ($torus.Value.CapacityOnly.ToBool() -or -not $torus.Value.ConnectCapacity)
    {
        Write-Error "Measure-StoreUsageStatistics.ps1 only works with Torus in management-capacity dual sessions."
        return
    }
    else
    {
        $useTorus = $true
    }
}

If ($Server)
{
    $ExServer = get-ExchangeServer $Server

    If ([String]$ExServer.AdminDisplayVersion -match "Version 14")
    {
        If ($AlertGuid)
        {
            If (test-path "\\$($ExServer.FQDN)\Exchange\Diagnostics\$AlertGuid")
            {
                $File = get-Item -path "\\$($ExServer.FQDN)\Exchange\Diagnostics\$AlertGuid\StoreUsageStatistics*.csv"
                $sus = Import-csv $File.FullName | ? {$_.DigestCategory -eq $DigestCategory}
            }
            Else
            {
                write-host -foregroundcolor:red "path not found (\\$($ExServer.FQDN)\Exchange\Diagnostics\$AlertGuid)"
                return
            }
        }
        Else
        {
            $sus = Get-StoreUsageStatistics -server $Server | ? {$_.DigestCategory -eq $DigestCategory}
        }
    }
    Else
    {
        $files = @(Get-TraceFileList -ExServer $ExServer -UseTorus $useTorus | ? {$_.Name.StartsWith("StoreUsageStatisticsData") -and $_.Extension -eq '.csv'})
        If ($Files.count -gt 0)
        {
            write-host "Analyzing most recent of $($Files.count) files at \\$($ExServer.FQDN)\Exchange\Diagnostics\StoreUsageStatistics"
            $recentfile = $Files | ? {$_.Length -gt -0} | Sort LastWriteTime | Select -last 1 

            if ($UseTorus)
            {
                $MLfilePath=@(Get-MachineLog -Target $ExServer.Name -Log StoreUsageStatistics -Filter $recentfile.name -DownloadToLocalFolderPath $WorkingPath | ? {$_.length -gt 0})
                Foreach ($filepath in $MLfilePath)
                {
                    If (Test-Path $filepath)
                    {
                        $file = Get-Item $filePath
                        break;
                    }
                }
                If ($file -eq $Null)
                {
                    write-error "Get-MachineLog failed to return contents of '$recentfile'"
                    return;
                }
            }
            Else
            {
                $file = $recentfile
            }

            $sus = Import-csv $File.FullName | ? {$_.DigestCategory -eq $DigestCategory} | % { `
                if (-not $_.ServerName) {Add-Member -InputObject $_ -MemberType NoteProperty -Name ServerName -Value ($ExServer.Name);}
                Write-Output $_;} 

            if ($useTorus) {remove-item $file}

        }
        Else
        {
            write-host -foregroundcolor:red "no files in path (\\$($ExServer.FQDN)\Exchange\Diagnostics\StoreUsageStatistics)"
            return                
        }
    }
}
ElseIf ($Database)
{
    $sus = Get-StoreUsageStatistics -Database $Database | ? {$_.DigestCategory -eq $DigestCategory}
}
ElseIf ($FileName)
{
    If ($Filename -match ".xml")
    {
        $sus = Import-Clixml $FileName | ? {$_.DigestCategory -eq $DigestCategory}
    }
    Else
    {
        $sus = Import-csv $FileName | ? {$_.DigestCategory -eq $DigestCategory}
    }
}
ElseIf ($Stats)
{
    $sus = $stats | ? {$_.DigestCategory -eq $DigestCategory}
}
Else
{
    help Measure-StoreUsageStatistics.ps1
    return
}

if ($Export)
{
    write-host "`n$(get-date) exported `$sus variable"
    $Global:sus = $sus
}

#exclude statistics that have no RPC Operations
$Sus = $Sus | ? {$_.RopCount -gt 0}

$TotalCountAboveThreshold = ($sus | ? {$_.TimeInServer/$_.RopCount -ge $RopLatencyThreshold} | measure).count; 
$TotalPctBelowThreshold = [Math]::Round(100*($sus.count - $TotalCountAboveThreshold)/($sus.count),1); 

$Groups = $sus | Group MailboxGuid
Foreach ($Group in $Groups) { `
    $TotalTimeInServer = (($group.group | measure TimeInServer -sum).Sum); `
    $TotalRopCount = (($group.group | measure RopCount -sum).Sum); `
    $Roplatency = [Math]::Round($TotalTimeInServer/$TotalRopCount,1); `
    $CountAboveThreshold = ($group.group | ? {$_.RopCount -gt 0} | ? {$_.TimeInServer/$_.RopCount -ge $RopLatencyThreshold} | measure).count; `
    $PctBelowThreshold = [Math]::Round(100*($group.count - $CountAboveThreshold)/($group.count),1); `
    $TotalRopSamplesBelowThreshold = @($group.group | ? {$_.RopCount -gt 0} | ? {$_.TimeInServer/$_.RopCount -lt $RopLatencyThreshold}); `
    If ($TotalRopSamplesBelowThreshold.count -gt 0)
    {
        $TotalRopCountBelowThreshold = ($TotalRopSamplesBelowThreshold | measure RopCount -sum).Sum; `
    }
    Else
    {
        $TotalRopCountBelowThreshold = 0; `
    }
    $PctRopsBelowThreshold = [Math]::Round(100*($TotalRopCountBelowThreshold/$TotalRopCount),1); `
    $TotalPageRead = (($group.group | measure PageRead -sum).Sum); `
    $TotalPagePreRead = (($group.group | measure PagePreRead -sum).Sum); `
    $TotalLogRecordBytes = (($group.group | measure LogRecordBytes -sum).Sum); `
    add-member -InputObject $group -Name RopCount -Value $TotalRopCount -MemberType NoteProperty; `
    add-member -InputObject $group -Name TimeInServer -Value $TotalTimeInServer -MemberType NoteProperty; `
    add-member -InputObject $group -Name RopLatency -Value $Roplatency -MemberType NoteProperty; `
    $PctSamplePropName = "PctSampleLt$($RopLatencyThreshold)ms"
    add-member -InputObject $group -Name $PctSamplePropName -Value $PctBelowThreshold -MemberType NoteProperty; `
    $PctRopPropName = "PctRopLt$($RopLatencyThreshold)ms"
    add-member -InputObject $group -Name $PctRopPropName -Value $PctRopsBelowThreshold -MemberType NoteProperty; `
    add-member -InputObject $group -Name PageRead -Value $TotalPageRead -MemberType NoteProperty; `
    add-member -InputObject $group -Name PagePreRead -Value $TotalPagePreRead -MemberType NoteProperty; `
    add-member -InputObject $group -Name LogRecordBytes -Value $TotalLogRecordBytes -MemberType NoteProperty; `
}; `

$SampleTimeGroups = $sus | Group SampleTime | Sort Name; `
Foreach ($Group in $SampleTimeGroups) { `
    $TotalTimeInServer = (($group.group | measure TimeInServer -sum).Sum); `
    $TotalRopCount = (($group.group | measure RopCount -sum).Sum); `
    $Roplatency = [Math]::Round($TotalTimeInServer/$TotalRopCount,1); `
    $CountAboveThreshold = ($group.group | ? {$_.RopCount -gt 0} | ? {$_.TimeInServer/$_.RopCount -ge $RopLatencyThreshold} | measure).count; `
    $PctBelowThreshold = [Math]::Round(100*($group.count - $CountAboveThreshold)/($group.count),1); `
    $TotalRopCountBelowThreshold = ($group.group | ? {$_.RopCount -gt 0} | ? {$_.TimeInServer/$_.RopCount -lt $RopLatencyThreshold} | measure RopCount -sum).Sum; `
    $PctRopsBelowThreshold = [Math]::Round(100*($TotalRopCountBelowThreshold/$TotalRopCount),1); `
    $TotalLogRecordBytes = (($group.group | measure LogRecordBytes -sum).Sum); `
    add-member -InputObject $group -Name RopCount -Value $TotalRopCount -MemberType NoteProperty; `
    add-member -InputObject $group -Name TimeInServer -Value $TotalTimeInServer -MemberType NoteProperty; `
    add-member -InputObject $group -Name RopLatency -Value $Roplatency -MemberType NoteProperty; `
    $PctSamplePropName = "PctSampleLt$($RopLatencyThreshold)ms"
    add-member -InputObject $group -Name $PctSamplePropName -Value $PctBelowThreshold -MemberType NoteProperty; `
    $PctRopPropName = "PctRopLt$($RopLatencyThreshold)ms"
    add-member -InputObject $group -Name $PctRopPropName -Value $PctRopsBelowThreshold -MemberType NoteProperty; `
    add-member -InputObject $group -Name LogRecordBytes -Value $TotalLogRecordBytes -MemberType NoteProperty; `
}; `


$TotalRopCount = ($Sus | measure RopCount -sum).sum; `
$TotalRopLatency = [math]::Round(($Sus | measure TimeInServer -sum).Sum/($Sus | measure RopCount -sum).sum,1); `
$start = (($sus | sort SampleTime)[0]).SampleTime; `
$end = (($sus | sort SampleTime)[$sus.count-1]).SampleTime; `

$ServerName = $sus[0].ServerName

$DBGroups = @($sus | Group DatabaseName)

If ($DBGroups.count -eq 1)
{
    write-host "Database $($DBGroups[0].group[0].Databasename) was active on $ServerName between $Start and $End"
}
Else
{
    write-host "$($DBGroups.count) databases were active on $ServerName between $Start and $End"
}

write-host "$TotalRopCount operations had average latency $TotalRopLatency msec for $($Groups.count) unique mailboxes between $Start and $End"; `
If ($sus.count -ge $MinimumSampleCount -and $TotalPctBelowThreshold -lt $PctSampleBelowThresholdToAlert -and $TotalRopLatency -gt $RopLatencyThreshold)
{
    write-host -foregroundColor red "*** This is a condition that requires further investigation ***"
    write-host -foregroundColor red "$TotalPctBelowThreshold% of $($sus.count) samples for $($Groups.count) unique mailboxes had average rop latency below $RopLatencyThreshold msec"
}
ElseIf ($sus.count -lt $MinimumSampleCount)
{
    write-host -foregroundColor green "*** This is a condition that DOES NOT require further investigation ***"
    write-host -foregroundColor green "$TotalPctBelowThreshold% of $($sus.count) samples for $($Groups.count) unique mailboxes had average rop latency below $RopLatencyThreshold msec (not enough samples)"
}
ElseIf ($TotalRopLatency -le $RopLatencyThreshold)
{
    write-host -foregroundColor green "*** This is a condition that DOES NOT require further investigation ***"
    write-host -foregroundColor green "$TotalRopCount operations had average latency $TotalRopLatency msec ($TotalPctBelowThreshold% of $($sus.count) samples for $($Groups.count) unique mailboxes had average rop latency below $RopLatencyThreshold msec)"
}
Else
{
    write-host -foregroundColor green "*** This is a condition that DOES NOT require further investigation ***"
    write-host -foregroundColor green "$TotalPctBelowThreshold% of $($sus.count) samples for $($Groups.count) unique mailboxes had average rop latency below $RopLatencyThreshold msec"
}

write-host "`n"

If ($IncludeDetailedAnalysis)
{
    Foreach ($group in $DBGroups) 
    {
        $TotalTimeInServer = (($group.group | measure TimeInServer -sum).Sum); 
        $TotalRopCount = (($group.group | measure RopCount -sum).Sum); 
        $Roplatency = [Math]::Round($TotalTimeInServer/$TotalRopCount,1);
        $CountAboveThreshold = ($group.group | ? {$_.RopCount -gt 0} | ? {$_.TimeInServer/$_.RopCount -ge $RopLatencyThreshold} | measure).count; `
        $PctBelowThreshold = [Math]::Round(100*($group.count - $CountAboveThreshold)/($group.count),1); `
        $TotalRopCountBelowThreshold = ($group.group | ? {$_.RopCount -gt 0} | ? {$_.TimeInServer/$_.RopCount -lt $RopLatencyThreshold} | measure RopCount -sum).Sum; `
        $PctRopsBelowThreshold = [Math]::Round(100*($TotalRopCountBelowThreshold/$TotalRopCount),1); `
        add-member -InputObject $group -Name RopCount -Value $TotalRopCount -MemberType NoteProperty; 
        add-member -InputObject $group -Name TimeInServer -Value $TotalTimeInServer -MemberType NoteProperty; 
        add-member -InputObject $group -Name RopLatency -Value $Roplatency -MemberType NoteProperty; 
        $PctSamplePropName = "PctSampleLt$($RopLatencyThreshold)ms"
        add-member -InputObject $group -Name $PctSamplePropName -Value $PctBelowThreshold -MemberType NoteProperty; `
        $PctRopPropName = "PctRopLt$($RopLatencyThreshold)ms"
        add-member -InputObject $group -Name $PctRopPropName -Value $PctRopsBelowThreshold -MemberType NoteProperty; `
    }; 

    write-host "Summary Performance by Database between $Start and $End"; 
    $DBgroups | sort Name | FT -a Name,Count,TimeInServer,RopCount,RopLatency,$PctSamplePropName,$PctRopPropName; `

    write-host "Top $TopUsers Mailbox by RopLatency between $Start and $End"; `
    $groups | sort RopLatency -descending | select -first $TopUsers | FT -a Name,Count,TimeInServer,RopCount,RopLatency,$PctSamplePropName,$PctRopPropName; `
    write-host "Top $TopUsers Mailbox by RopCount between $Start and $End"; `
    $groups | sort RopCount -descending | select -first $TopUsers | FT -a Name,Count,TimeInServer,RopCount,RopLatency,$PctSamplePropName,$PctRopPropName; `
    write-host "Top $TopUsers Mailbox by TimeInServer between $Start and $End"; `
    $groups | sort TimeInServer -descending | select -first $TopUsers | FT -a Name,Count,TimeInServer,RopCount,RopLatency,$PctSamplePropName,$PctRopPropName; `
    write-host "Top $TopUsers Mailbox by PageRead between $Start and $End"; `
    $groups | sort PageRead -descending | select -first $TopUsers | FT -a Name,Count,TimeInServer,PageRead,PagePreRead,LogRecordBytes,$PctSamplePropName,$PctRopPropName; `
    write-host "Top $TopUsers Mailbox by PagePreRead between $Start and $End"; `
    $groups | sort PagePreRead -descending | select -first $TopUsers | FT -a Name,Count,TimeInServer,PageRead,PagePreRead,LogRecordBytes,$PctSamplePropName,$PctRopPropName; `
    write-host "Top $TopUsers Mailbox by LogRecordBytes between $Start and $End"; `
    $groups | sort LogRecordBytes -descending | select -first $TopUsers | FT -a Name,Count,TimeInServer,PageRead,PagePreRead,LogRecordBytes,$PctSamplePropName,$PctRopPropName; `
    write-host "Grouped by SampleTime between $Start and $End"; `
    $SampleTimeGroups | sort Name | FT -a Name,Count,TimeInServer,RopCount,RopLatency,$PctSamplePropName,$PctRopPropName,LogRecordBytes; 
}

# SIG # Begin signature block
# MIIdwAYJKoZIhvcNAQcCoIIdsTCCHa0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+56/UICb0anq87ac8kEmARcV
# r7ugghhkMIIEwzCCA6ugAwIBAgITMwAAAJ1CaO4xHNdWvQAAAAAAnTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTMw
# WhcNMTcwNjMwMTkyMTMwWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjE0OEMtQzRCOS0yMDY2MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAy8PvNqh/8yl1
# MrZGvO1190vNqP7QS1rpo+Hg9+f2VOf/LWTsQoG0FDOwsQKDBCyrNu5TVc4+A4Zu
# vqN+7up2ZIr3FtVQsAf1K6TJSBp2JWunjswVBu47UAfP49PDIBLoDt1Y4aXzI+9N
# JbiaTwXjos6zYDKQ+v63NO6YEyfHfOpebr79gqbNghPv1hi9thBtvHMbXwkUZRmk
# ravqvD8DKiFGmBMOg/IuN8G/MPEhdImnlkYFBdnW4P0K9RFzvrABWmH3w2GEunax
# cOAmob9xbZZR8VftrfYCNkfHTFYGnaNNgRqV1rEFt866re8uexyNjOVfmR9+JBKU
# FbA0ELMPlQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFGTqT/M8KvKECWB0BhVGDK52
# +fM6MB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAD9dHEh+Ry/aDJ1YARzBsTGeptnRBO73F/P7wF8dC7nTPNFU
# qtZhOyakS8NA/Zww74n4gvm1AWfHGjN1Ao8NiL3J6wFmmON/PEUdXA2zWFYhgeRe
# CPmATbwNN043ecHiGjWO+SeMYpvl1G4ma0NIUJau9DmTkfaMvNMK+/rNljr3MR8b
# xsSOZxx2iUiatN0ceMmIP5gS9vUpDxTZkxVsMfA5n63j18TOd4MJz+G0I62yqIvt
# Yy7GTx38SF56454wqMngiYcqM2Bjv6xu1GyHTUH7v/l21JBceIt03gmsIhlLNo8z
# Ii26X6D1sGCBEZV1YUyQC9IV2H625rVUyFZk8f4wggYHMIID76ADAgECAgphFmg0
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUZln7ZI2yebsWwHppekZSKsYLov0wegYKKwYB
# BAGCNwIBDDFsMGqgQoBAAE0AZQBhAHMAdQByAGUALQBTAHQAbwByAGUAVQBzAGEA
# ZwBlAFMAdABhAHQAaQBzAHQAaQBjAHMALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBADyRyVG5nre3
# AZy4VklrIpa9yk7kI/AadhV7waXPIGSfZyYmo01B1WbZ+R5ILk/5GIIwRTvH8dBf
# 0Z2wAPkNTwLrwkHzNfXuKWU60HH5i1qBhiw+zumiYKBpRv+urcSCMNkGCEnbgS5o
# JMjpMxv+abPZ9e7j9FrLlFUowjkiPGrG2jitYrQMVMC5sdnAI58G/IHNqA1brTuc
# antbYXyvTumPKnBEE+3qzxJeGGizYjvfWz+UNBvE6VYRdTu+pkiDt3kW4R1FAOcf
# OYfNNf8auK6q1oNKrLHIPFxZ81ZfqbLWrA/8MqtuI0rCFMs9POtWHGqPoHZl1sAr
# QVnHvbwP+HqhggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAnUJo7jEc11a9AAAAAACdMAkGBSsOAwIa
# BQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0x
# NjA5MDMxODQ1MDZaMCMGCSqGSIb3DQEJBDEWBBSGAmC8IQEFiMT2U2n+hm8Ejm1F
# oDANBgkqhkiG9w0BAQUFAASCAQCrdmv+30M4G6sApwm3m1M7FKKI7GIw1C9IgVB5
# HVQmXYpLNm85LosI4smdzfn3a7ZbXj6vXN2ji0Rg7cp8+je4FFac6pkPMc10Ih4z
# MOob6HoCrycx2hfejqdwy6d+GeOsxR+zaFiYxi1xoZGJQbED/NzEqvVcQe8yXqMM
# 2L+ODRirlVmLw999yndum1dJkm9v9WnA99PCnRYN0izm1mr7p0nvSfFkoo1EgmhP
# wMO1wQPr44a75oLb8zvM3EsQ4uDQ2m9SuKdSbpe4PW5K3qpv9YDuKdx1LF32mUtS
# GC/TivXCpAD0Y1p4Pt1AqiJRpqkZGZLMl17T8RTi7SL99RVh
# SIG # End signature block
