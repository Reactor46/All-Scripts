# Copyright (c) 2011 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

<#
.SYNOPSIS 
Get Managed Availability results from Exchange 15 servers.
.DESCRIPTION
Gets the specified set of Managed Availability results from the specified machines.
.EXAMPLE
Get-ExchangeServer | ? { $_.Name -like 'PIO-MBX-0*' } | Get-MADefinitions -Type Probe -Filter @{Name="OWA.Protocol/OwaSelfTestProbe" } | Get-MAResults -MaxResults 1

Gets the last result for each of the probe definitions with the given name from each of the machines.
.EXAMPLE
Get-ExchangeServer | ? { $_.Name -like 'PIO-MBX-0*' } | Get-MADefinitions -Type Probe -Filter @{Name="OWA.Protocol/OwaDeepTestProbe" } | Get-MAResults -Filter @{ResultType=4} -MaxResults 1

Gets the last failed result for each of the probe definitions with the given name from each of the machines.
.PARAMETER Definitions
The definitions to get results for. These definitions must be output from a single call to Get-MADefinitions.
.PARAMETER FromUtc
The start of the time range to query, in UTC. Defaults to 15 minutes ago.
.PARAMETER ToUtc
The end of the time range to query, in UTC. Defaults to now.
.PARAMETER MaxResults
The maximum number of results to query, per definition. Defaults to 1.
.PARAMETER Filter
The filtered set of results to query, in hash table format. Defaults to empty (i.e. queries all results).
#>
function Get-MAResults
{
param(
    [parameter(Mandatory=$true,
    ValueFromPipeline=$true)]$Definitions,
    [parameter(Mandatory=$false)][DateTime]$FromUtc = ([DateTime]::UtcNow.AddMinutes(-15)),
    [parameter(Mandatory=$false)][DateTime]$ToUtc = ([DateTime]::UtcNow),
    [parameter(Mandatory=$false)][int]$MaxResults = 1,
    [parameter(Mandatory=$false)][Hashtable]$Filter = @{}
)

    begin
    {
        $DefinitionsList = @()
    }

    process
    {
        # save off the Definitions into a separate list for later access
        foreach ($def in $Definitions)
        {
            $DefinitionsList += $def
        }
    }

    end
    {
        $machines = @()
        $Type = $null
        $queries = @{}

        # In testing, I have found that the XPath can only include about 22 conditions.
        # Two of the conditions are standard (begin and end time), so we can only have 20 other conditions.
        # Each definition translates to a condition (WorkItemId = blah).
        # So if there are more definitions than can fit into a single query, we need to split the query into multiple calls.

        $MaxDefinitionsPerQuery = 20 - $Filter.Count

        $iterations = [Math]::Floor(($DefinitionsList.Count + $MaxDefinitionsPerQuery - 1) / $MaxDefinitionsPerQuery)

        Write-Verbose ("Number of definitions = '{0}', definitions per query = '{1}' and iterations = '{2}'" -f $DefinitionsList.Count, $MaxDefinitionsPerQuery, $iterations)

        for ($i = 0; $i -lt $iterations; ++$i)
        {
            $xpath = "*[("
            $totalResults = 0

            for ($j = 0; $j -lt $MaxDefinitionsPerQuery; ++$j)
            {
                $index = $i * $MaxDefinitionsPerQuery + $j
                if ($DefinitionsList.Count -le $index)
                {
                    break;
                }

                $def = $DefinitionsList[$index]

                $xpath += ("(UserData/EventXML/WorkItemId = '{0}') or " -f $def.Id)

                if ($machines -notcontains $def.MachineName)
                {
                    $machines += $def.MachineName
                }

                if ($Type -eq $null)
                {
                    # Because the input is from the output of Get-MADefinitions
                    # and Get-MADefinitions will only look at one channel at a time
                    # inspecting the WorkItemType of the first definition is enough.
                    $Type = $def.WorkItemType
                }

                $totalResults += $MaxResults
            }

            $xpath += ")"
            $xpath = $xpath.Replace(" or )", ") and ")

            foreach ($key in $Filter.Keys)
            {
                $xpath += ("(UserData/EventXML/{0} = '{1}') and " -f $key, $Filter[$key])
            }

            $xpath += ("(System[TimeCreated[@SystemTime > '{0}']]) and (System[TimeCreated[@SystemTime < '{1}']])]" -f $FromUtc.ToString("o"), $ToUtc.ToString("o"))
            $queries.Add($xpath, $totalResults)
        }

        foreach ($machine in $machines)
        {
            foreach ($query in $queries.Keys)
            {
                Write-Verbose ("Executing query '{0}' on machine '{1}' with MaxEvents set to '{2}'" -f $query, $machine, $totalResults)
                $events = Get-WinEvent -ComputerName $machine -LogName ("Microsoft-Exchange-ActiveMonitoring/{0}Result" -f $Type) -FilterXPath $query -MaxEvents $queries[$query] -ErrorAction SilentlyContinue
                if ($events -eq $null -or $events.Count -eq 0)
                {
                    Write-Warning "No results matching the specified criteria found on server '$machine'."
                }
                else
                {
                    foreach ($event in $events)
                    {
                        $xml = [xml]$event.ToXml()
                        $xml.Event.UserData.EventXML
                    }
                }
            }
        }
    }
}

<#
.SYNOPSIS
Get Managed Availability definitions from Exchange 15 servers.
.DESCRIPTION
Gets the specified set of Managed Availability definitions from the specified machines.
.EXAMPLE
Get-ExchangeServer | ? { $_.Name -like 'PIO-CF15-*' } | Get-MADefinitions -Type Probe -Filter @{Name='HealthManagerObserverProbe'} | ft MachineName, Enabled, *Seconds, TargetResource -a

Gets the MachineName, Enabled, RecurrenceIntervalSeconds, TimeoutSeconds and TargetResource fields for all definitions with the given name.
.PARAMETER Identity
The Exchange 15 server(s) to query.
.PARAMETER Type
The type of definition to query - should be one of Probe, Monitor, Responder or Maintenance.
.PARAMETER Filter
The filtered set of results to query, in hash table format. Defaults to empty (i.e. queries all definitions).
#>

function Get-MADefinitions
{
param(
    [parameter(Mandatory=$true,
    ValueFromPipeline=$true,
    ValueFromPipelineByPropertyName=$true)]$Identity,
    [parameter(Mandatory=$true)][ValidateSet("Probe", "Monitor", "Responder", "Maintenance")][string]$Type,
    [parameter(Mandatory=$false)][Hashtable]$Filter = @{}
)

    process
    {
        $xpath = $null

        if ($Filter.Count -gt 0)
        {
            $xpath = "*["

            foreach ($key in $Filter.Keys)
            {
                $xpath += ("(UserData/EventXML/{0} = '{1}') and " -f $key, $Filter[$key])
            }

            $xpath += "]"
            $xpath = $xpath.Replace(" and ]", "]")
        }

        foreach ($id in $Identity)
        {
            if ($xpath -eq $null)
            {
                Write-Verbose ("Getting all definitions from machine '{0}'" -f $id)
                $events = Get-WinEvent -ComputerName $id -LogName ("Microsoft-Exchange-ActiveMonitoring/{0}Definition" -f $Type) -ErrorAction SilentlyContinue
            }
            else
            {
                Write-Verbose ("Getting all definitions matching filter '{0}' from machine '{1}'" -f $xpath, $id)
                $events = Get-WinEvent -ComputerName $id -LogName ("Microsoft-Exchange-ActiveMonitoring/{0}Definition" -f $Type) -FilterXPath $xpath -ErrorAction SilentlyContinue
            }

            if ($events -eq $null -or $events.Count -eq 0)
            {
                Write-Warning "No definitions matching the specified criteria found on server '$id'."
            }
            else
            {
                foreach ($event in $events)
                {
                    $xml = [xml]$event.ToXml()
                    $xml.Event.UserData.EventXML `
                        | Add-Member -MemberType NoteProperty -Name MachineName -Value $id -PassThru `
                        | Add-Member -MemberType NoteProperty -Name WorkItemType -Value $Type -PassThru
                }
            }
        }
    }
}

<#
.SYNOPSIS 
Start Real Time Event Watcher for Probe, Monitor, Responder Results
.DESCRIPTION 
Start Real Time Event Watcher for Probe, Monitor, Responder Results
.EXAMPLE 
. ./ActiveMonitoringCommonLibrary.ps1
Start-ResultWatcher -RemoteMachine "PIO-CF15-01"
.NOTES 
Default Values
$InputQuery="*[System[(EventID=2)]]"
$Type="Probe"
#>

function Start-ResultWatcher([String]$InputQuery="*[System[(EventID=2)]]", [String]$RemoteMachine, [String]$Type="Probe", [String]$Component)
{
    if(![String]::IsNullOrEmpty($Component))
    {
        switch ($Component.Trim().ToLower())
        {
            "ewsgen" { $InputQuery="*[UserData/EventXML/ResultName='EWSGeneric']" }
            "hostservice" { $InputQuery="*[UserData/EventXML/ResultName='Host Controller Service Node Availability Probe/HostControllerService']" }
            "health" { $InputQuery="*[UserData/EventXML/ResultName='Health Manager Heartbeat Probe']" }
            "autodiscoverxml" { $InputQuery="*[UserData/EventXML/ResultName='AutoDiscoverXml']" }
            "autodiscoversvc" { $InputQuery="*[UserData/EventXML/ResultName='AutoDiscoverSvc']" }
            "oab" { $InputQuery="*[UserData/EventXML/ResultName='OABDownload']" }
            "availability" { $InputQuery="*[UserData/EventXML/ResultName='PROD/AvailabilityService']" }
            "eas" { $InputQuery="*[UserData/EventXML/ResultName='ActiveSync']" }
            "owa" { $InputQuery="*[UserData/EventXML/ResultName='OWA']" }
            "emsmdb" { $InputQuery="*[UserData/EventXML/ResultName='Emsmdb']" }
            "topology" { $InputQuery="*[UserData/EventXML/ResultName='MSExchangeTopologyService-Microsoft.Office.Datacenter.ActiveMonitoring.GenericServiceProbe/MSExchangeTopologyService']" }
            "mailflow" { $InputQuery="*[UserData/EventXML/ResultName='MailFlow']" }
        }
    }    
    
    $LogName = @'
Microsoft-Exchange-ActiveMonitoring/{0}Result
'@ -f $Type

    [System.Diagnostics.Eventing.Reader.PathType]$PathType = [System.Diagnostics.Eventing.Reader.PathType]::LogName
    $EventLogQuery = New-Object System.Diagnostics.Eventing.Reader.EventLogQuery $LogName, $PathType, $InputQuery
    $EventLogQuery.TolerateQueryErrors = $True
    
    if(![String]::IsNullOrEmpty($RemoteMachine))
    {
        $Session = New-Object System.Diagnostics.Eventing.Reader.EventLogSession $RemoteMachine
        $EventLogQuery.Session = $Session
    }
    
    $EventLogWatcher = New-Object System.Diagnostics.Eventing.Reader.EventLogWatcher $EventLogQuery
    
    [ScriptBlock]$Action = 
    { 
        Write-Host ("---------------------------------------------------------------------")
        Write-Host ("[ {0:g} ]"-f $Event.TimeGenerated)
        Write-Host ("[ Name: {0} ]"-f $EventRecordXML.Event.UserData.EventXML.ResultName)
        Write-Host ("[ Event Data: {0} "-f $EventRecordXML.Event.UserData.EventXML.Exception)
    }
                        
    [String]$SourceIdentifier = "NewEventLog"
    [PSObject]$MessageData = $Null

    [ScriptBlock]$ObjectEventAction = 
    {    
        $EventRecord = $EventArgs.EventRecord
        [XML]$EventRecordXML = $EventRecord.ToXML()
    }
        
    If ($Action -ne $Null)
    {
        $ObjectEventAction = [ScriptBlock]::Create($ObjectEventAction.ToString() + "`n" + $Action.ToString())
    }
        
    $ObjectEventParams = 
    @{
        'InputObject' = $EventLogWatcher;
        'SourceIdentifier' = $SourceIdentifier;
        'EventName' = 'EventRecordWritten';
        'Action' = $ObjectEventAction;
        'MessageData' = $MessageData;
    }

    try
    {
        Register-ObjectEvent @ObjectEventParams
    }
    catch
    {
        Write-Error ("Error registereing EventRecordWritten Event`nERROR: {0}" -f $_.Exception.Message)
    }
    
    $EventLogWatcher.Enabled = $True
    return $EventLogWatcher
}
# SIG # Begin signature block
# MIIdwgYJKoZIhvcNAQcCoIIdszCCHa8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUbWoSpCSzFK5IXL8eSFYJIHHY
# M0KgghhkMIIEwzCCA6ugAwIBAgITMwAAAJqamxbCg9rVwgAAAAAAmjANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBMgwggTEAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB3DAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUaaZih4aUNcicf4NkRNpn7M0586MwfAYKKwYB
# BAGCNwIBDDFuMGygRIBCAEEAYwB0AGkAdgBlAE0AbwBuAGkAdABvAHIAaQBuAGcA
# QwBvAG0AbQBvAG4ATABpAGIAcgBhAHIAeQAuAHAAcwAxoSSAImh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEASv7IzbAh
# VgKMW5yxROOb89xEuHImIfnka1sYqE5UvORzPQDYpPmCFVBSUzTYIe5ud9qhkUjc
# wXZ5SdHhDdbyN4MPX0HuhGem7yGsvsoGV2GaLonp27XKFnresTPRogALkNvwb+sP
# HcGmL0jrcyLa059GKhs8DAfqyTKoljvzJYtMivhMVi2yHR1bC/ma1V8NNU0ly0ik
# rJcZFkaJxYqghgP3gcJ5E38KyubVRQv675x1ia/+XiL9FfSBQ4SmDuBuTyfR+nuP
# lK454CPL8hzgzP7Y+fIO0DDwx8cr9N6m/MCCBQALTH+Cof3TIPnDJA/En9F6X4km
# Lyw7m/Z/qhrBiKGCAigwggIkBgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAACampsWwoPa1cIAAAAAAJowCQYFKw4D
# AhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8X
# DTE2MDkwMzE4NDQ1NVowIwYJKoZIhvcNAQkEMRYEFEa/r3dgrNCBQ6p/RPZhCpZC
# JlEtMA0GCSqGSIb3DQEBBQUABIIBAKXDVE7mDkJ10ubshSlqhxwPp0+FaVN+L/0J
# mlP0CnzBDzEmEQl/qoynQPe3O9Ex1RXFFJIRZ3xIARXjdhjJdu1Vfy7RGKacXZiV
# 3v8glgnwg47ZMxfBW+HKrhEX8+gMsKi+C5MgG1Uqqut4Vsi3qKNAPPZv4WgY1ckt
# bmo143yzpRnyx5NbEJmbKboBbVrppuwZ2c0zXxJTrvC0PZWBkvDBFMlwmfswOeny
# V4uujWsmchPbDt6Kq16ShOogKkKfoRw0sBylptp2JmWTHkFnwhUnRcZu7aIuBhbI
# nbe+2Ixkh3BAHTOdO6rXuFKsG9cjJNrmU0KDRVSYmRsWwIS8+24=
# SIG # End signature block
