# Copyright (c) 2009 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

# Requires -Version 2

<#
   .SYNOPSIS 
   Performs troubleshooting on Content Index (CI) catalogs. 

   .DESCRIPTION
   The Troubleshoot-CI.ps1 script detects problems with content index
   catalogs and optionally attempts resolutions to the problems.

   .PARAMETER Server
   The simple NETBIOS name of mailbox server on which troubleshooting
   should be atempted for CI catalogs. If this optional parameter is
   not specified, local server is assumed. 

   .PARAMETER Database
   The name of database to troubleshoot. If this optional parameter is
   not specified, catalogs for all databases on the server specified
   by the Server parameter are troubleshooted.
   
   .PARAMETER Symptom
   Specifies the symptom to detect. Possible values are:
   'Deadlock', 'Corruption', 'Stall', 'Backlog' and 'All'.
   When 'All' is specified, all the first four symptoms in
   the list are checked.
   
   If this optional parameter is not specified, 'All' is assumed.
   
    .PARAMETER Action
   Specifies the action to be performed to resolve a symptom. The
   possible values are 'Detect', 'DetectAndResolve', 'Resolve'.
   The default value is 'Detect'
     
    .PARAMETER MonitoringContext
   Specifies if the command is being run in a monitoring context.
   The possible values are $true and $false. Default is $false.
   If the value is $true, warning/failure events are logged to the
   application event log. Otherwise, they are not logged.
   
   .PARAMETER FailureCountBeforeAlert
   Specifies the number of failures the troubleshooter will allow
   before raising an Error in the event log, leading to a SCOM alert.
   The allowed range for this parameter is [1,100], both inclusive.
   No alerts are rasised if MonitoringContext is $false.
   
   .PARAMETER FailureTimeSpanMinutes
   Specifies the number of minutes in the time span during which
   the troubleshooter will check the history of failures to count
   the failures and alert. If the failure count during this time
   span exceeds the value for FailureCountBeforeAlert, an alert
   is raised. No alerts are rasised if MonitoringContext is $false.
   The default value for this parameter is 600, which is equivalent
   to 10 hours.

   .INPUTS
   None. You cannot pipe objects to Troubleshoot-CI.ps1.

   .OUTPUTS
   Returns status information about each catalog, problems detected
   and resolution actions performed, if any

   .EXAMPLE
   C:\PS> .\Troubleshoot-CI.ps1 –database DB01
   Detects and reports if there’s any problem with catalog for 
   database DB01. Does not attempt any Resolution. 

   .EXAMPLE
   C:\PS> .\Troubleshoot-CI.ps1 –database DB01 –symptom IndexingStall	
   Detects if indexing on catalog for database DB01 is stalled. Does not attempt any Resolution.
   
   .EXAMPLE
   C:\PS> .\Troubleshoot-CI.ps1 –Server <S001>	
   Detects and reports problems with all catalogs on server S001, if any. Does not attempt any Resolution.
   
   .EXAMPLE
   C:\PS> .\Troubleshoot-CI.ps1 –database DB01 –Action DetectAndResolve	
   Detects and reports if there’s any problem with catalog for database DB01. 
   Attempts a Resolution of the problem.
   
   .EXAMPLE
   C:\PS> .\Troubleshoot-CI.ps1 –database DB01 –Symptom Corruption –Action Resolve	
   Attempts a Resolution action for catalog corruption for database DB01. 
#>

[CmdletBinding()]
PARAM(
    [parameter( 
        Mandatory=$false, 
# $Research$ For some reason, we can not do import-localizeddata before PARAM.
# So, we can not use localized strings in this help message. This help message
# is used only in prompting for mandatory parameters, so, this is not a big 
# issue for now.
        HelpMessage = "The server to troubleshoot." 
       )] 
      [String] 
      [ValidateNotNullOrEmpty()] 
      $Server,
      
    [parameter(
        Mandatory=$false, 
        HelpMessage = "The database of catalog to troubleshoot." 
       )
      ] 
      [ValidateNotNullOrEmpty()] 
      [String] 
      $Database,
      
    [parameter(
        Mandatory=$false, 
        HelpMessage = "The symptom to detect and/or recover from." 
       )
      ] 
      [String] 
      [ValidateSet("Deadlock", "Corruption", "Stall", "All")] 
      $Symptom = "All",
      
    [parameter(
        Mandatory=$false, 
        HelpMessage = "The action to perform for each symptom." 
       )
      ] 
      [String] 
      [ValidateSet("Detect", "DetectAndResolve", "Resolve")] 
      $Action = "Detect",

    [parameter(
        Mandatory=$false, 
        HelpMessage = "Indicates if command is being run in a monitoring context. This flag is used to determine if we need to log warnings/failures to the application event log." 
       )
      ] 
      [Switch] 
      $MonitoringContext = $false,
      
    [parameter(
        Mandatory=$false, 
        HelpMessage = "Number of failures allowed before raising an error in the event log, leading to a SCOM alert, if applicable." 
       )
      ] 
      [Int32]
      [ValidateRange(1,100)]
      $FailureCountBeforeAlert = 3,
      
    [parameter(
        Mandatory=$false, 
        HelpMessage = "The number of minutes back in time during which we will check the history of failures to count total failures. This is related to the argument 'FailureCountBeforeAlert'." 
       )
      ] 
      [Int32]
      # between 10 minutes and 7 days
      [ValidateRange(10, 10080)]
      # default is 10 hours
      $FailureTimeSpanMinutes = 600
 )

$scriptDir = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)

. $scriptDir\CITSLibrary.ps1

# PS 303295
. $scriptDir\DiagnosticScriptCommonLibrary.ps1

try
{
    if ($Verbose)
    {
        $VerbosePreference = "Continue"
    }

    # Load Exchange Snapin. This is needed when
    # we want to run troubleshooter from a raw
    # powershell session or as a scheduled task.
    # Note: This function will only load it
    # if the snapin is not already loaded.
    #
    Load-ExchangeSnapin    
        
    write-verbose `
        ("Server=" + $Server + `
         " Database=" + $Database + `
         " Symptom=" + $Symptom + `
         " Action=" + $Action + `
         " MonitoringContext=" + $MonitoringContext + `
         " FailureCountBeforeAlert=" + $FailureCountBeforeAlert + `
         " FailureTimeSpanMinutes=" + $FailureTimeSpanMinutes)
         
    # Validate arguments
    #
    if ($MonitoringContext)
    {
        $Arguments = Validate-Arguments `
            -Server $Server `
            -Database $Database `
            -Symptom $Symptom `
            -Action $Action `
            -MonitoringContext `
            -FailureCountBeforeAlert $FailureCountBeforeAlert `
            -FailureTimeSpanMinutes $FailureTimeSpanMinutes
    }
    else
    {
        $Arguments = Validate-Arguments `
            -Server $Server `
            -Database $Database `
            -Symptom $Symptom `
            -Action $Action `
            -FailureCountBeforeAlert $FailureCountBeforeAlert `
            -FailureTimeSpanMinutes $FailureTimeSpanMinutes
    }
    
    
    # Log 'Troubleshooter Started' event
    #
    Log-Event -Arguments $Arguments -EventInfo $LogEntries.TSStarted
       
    if ($Arguments.Action -ieq "Resolve")
    {
        # build server status object to reflect
        # the given sysmptom
        #
        $serverStatus = Build-ServerStatus `
            -Server $Arguments.Server `
            -Database $Arguments.Database `
            -Symptom $Symptom
    }
    else
    {
        # Detect any problems of all catalogs 
        # on the specified server
        #
        $serverStatus = Detect-Problems $Arguments.Server $Arguments.Database
        
        # output the detection status
        #
        $serverStatus
        
        # Log the detection results. This will turn
        # on/off specific alerts based on the issues.
        #
        Log-DetectionResults $Arguments $ServerStatus
    }
        
    # If Action=='DetectAndResolve' proceed to resolution
    # of all symptoms detected.
    #
    if (($Arguments.Action -ieq "DetectAndResolve") -or
        ($Arguments.Action -ieq "Resolve"))
    {
        Resolve-Problems $Arguments $serverStatus
    }
    
    # If we are here, log that we 
    # have successfully finished troubleshooting
    #
    Log-Event -Arguments $Arguments -EventInfo $LogEntries.TSSuccess
    if ($MonitoringContext)
    {
	#PS 303295 - add a monitoring event to supress SCOM failure alerts.
	#
    	Add-MonitoringEvent -Id $LogEntries.TSSuccess[0] -Type $EVENT_TYPE_INFORMATION -Message $LocStrings.TSStarted
    }
}
catch [System.Exception]
{
    $message = ($LocStrings.TroubleshooterFailed + $error[0].Exception.ToString() + $error[0].InvocationInfo.PositionMessage)
    write-host $message
    Log-Event `
        -Arguments $Arguments `
        -EventInfo $LogEntries.TSFailed `
        -Parameters @($message)
        
    if ($MonitoringContext)
    {
        #PS 303295 - add a monitoring event to supress SCOM failure alerts.
        #
        Add-MonitoringEvent -Id $LogEntries.TsFailed[0] -Type $EVENT_TYPE_ERROR -Message $message
    }        
}

if ($MonitoringContext)
{
    # PS 303295 Output monitoring events.
    #
    Write-MonitoringEvents
}
      

# SIG # Begin signature block
# MIIdpgYJKoZIhvcNAQcCoIIdlzCCHZMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+k15WALkaPP7IxJEVXQ8J72E
# O/CgghhkMIIEwzCCA6ugAwIBAgITMwAAAJzu/hRVqV01UAAAAAAAnDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTMw
# WhcNMTcwNjMwMTkyMTMwWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjU4NDctRjc2MS00RjcwMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzCWGyX6IegSP
# ++SVT16lMsBpvrGtTUZZ0+2uLReVdcIwd3bT3UQH3dR9/wYxrSxJ/vzq0xTU3jz4
# zbfSbJKIPYuHCpM4f5a2tzu/nnkDrh+0eAHdNzsu7K96u4mJZTuIYjXlUTt3rilc
# LCYVmzgr0xu9s8G0Eq67vqDyuXuMbanyjuUSP9/bOHNm3FVbRdOcsKDbLfjOJxyf
# iJ67vyfbEc96bBVulRm/6FNvX57B6PN4wzCJRE0zihAsp0dEOoNxxpZ05T6JBuGB
# SyGFbN2aXCetF9s+9LR7OKPXMATgae+My0bFEsDy3sJ8z8nUVbuS2805OEV2+plV
# EVhsxCyJiQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFD1fOIkoA1OIvleYxmn+9gVc
# lksuMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAFb2avJYCtNDBNG3nxss1ZqZEsphEErtXj+MVS/RHeO3TbsT
# CBRhr8sRayldNpxO7Dp95B/86/rwFG6S0ODh4svuwwEWX6hK4rvitPj6tUYO3dkv
# iWKRofIuh+JsWeXEIdr3z3cG/AhCurw47JP6PaXl/u16xqLa+uFLuSs7ct7sf4Og
# kz5u9lz3/0r5bJUWkepj3Beo0tMFfSuqXX2RZ3PDdY0fOS6LzqDybDVPh7PTtOwk
# QeorOkQC//yPm8gmyv6H4enX1R1RwM+0TGJdckqghwsUtjFMtnZrEvDG4VLA6rDO
# lI08byxadhQa6k9MFsTfubxQ4cLbGbuIWH5d6O4wggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBKwwggSoAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBwDAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUxpJ27zrDdof6xPo5RJPzoIQMo00wYAYKKwYB
# BAGCNwIBDDFSMFCgKIAmAFQAcgBvAHUAYgBsAGUAcwBoAG8AbwB0AC0AQwBJAC4A
# cABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkq
# hkiG9w0BAQEFAASCAQAo9IXmAlzAUPZnkcSEwOpjcW6WmVGs59Bd1sv0aJCpUXiB
# QmWKmhAe4LjVHCXpQUv1fHumj6PlwjsmZKlBUVIA2+qC+8yrrOXTtJYA3slVuVV3
# U2NB+ztMA94qhPP0TYHT7Ks+HN306bpdYAqZLjh4Hgk7Gk+fsKCwIIeaMBkiC/G6
# I9JWKqtKik2bcYYLFQvV+TnhcPbZTHKdUBlXb00MBOqFIJrdQkpRN3JD+CSymoC5
# ACs1Jof44WqxM+NozRZqyGK1ODBK7f9pkzdt6TLahqbq+4EJhUaD6eY3FXvADhR8
# 2uG+qGXrUT7S/NBHhKE6rjscfqrQ3VHUc3MG1GuwoYICKDCCAiQGCSqGSIb3DQEJ
# BjGCAhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
# dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
# YXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAJzu
# /hRVqV01UAAAAAAAnDAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3
# DQEHATAcBgkqhkiG9w0BCQUxDxcNMTYwOTAzMTg0NDU4WjAjBgkqhkiG9w0BCQQx
# FgQUX0/b1La1tKipd4s6UDyiR/2s2AswDQYJKoZIhvcNAQEFBQAEggEAyeO1SC8p
# JApUI4/+1HOJj4y/hPNtCLT/suClGVjuUe5Z0VPeudJgfkTymMdgpPEy5j4zWMLG
# VLKDlZFMwMnstmqIOJyJVP0TSAsi52pr8PXe2Urb9hu9lH5ZUhRJTWW7Htla6CHZ
# Ujd/egTKF63vB49U7H5KxzzUZgQ5y5saW4Eilz563mKYBz1Mh9pcF1Bi7IGNWGY/
# I2SAvI08wIRDsfPJ2hQ6UBC4EDGD1mKc+K+mOZ8vswUqmStMC41MqXd6srImYJwI
# c+5EXmNYeNaHX9urBFWP4jBPIVl91Ytpq90ZpA8C7P4/9wbplvNjNTZOz53uCDH1
# +To1hPg2+aMCdQ==
# SIG # End signature block
