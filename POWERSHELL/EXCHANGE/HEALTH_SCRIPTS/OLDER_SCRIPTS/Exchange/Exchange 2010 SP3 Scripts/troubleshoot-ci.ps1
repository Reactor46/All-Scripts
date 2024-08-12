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
      [ValidateSet("Deadlock", "Corruption", "Stall", "MsftefdHealth", "All")] 
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
      $FailureTimeSpanMinutes = 600,

       [parameter(
        Mandatory=$false, 
        HelpMessage = "Indicates if the TS is allowed to take crash dumps" 
       )
      ] 
      [bool] 
      $CanTakeProcessCrashDumps = $True
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
         " FailureTimeSpanMinutes=" + $FailureTimeSpanMinutes + `
         " MaxCumulativeMsftefdMemoryConsumption=" + $MaxCumulativeMsftefdMemoryConsumption + `
         " CanTakeProcessCrashDumps=" + $CanTakeProcessCrashDumps)
         
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
            -FailureTimeSpanMinutes $FailureTimeSpanMinutes `
            -CanTakeProcessCrashDumps $CanTakeProcessCrashDumps
    }
    else
    {
        $Arguments = Validate-Arguments `
            -Server $Server `
            -Database $Database `
            -Symptom $Symptom `
            -Action $Action `
            -FailureCountBeforeAlert $FailureCountBeforeAlert `
            -FailureTimeSpanMinutes $FailureTimeSpanMinutes `
            -CanTakeProcessCrashDumps $CanTakeProcessCrashDumps
    }
    
    
    # Log 'Troubleshooter Started' event
    #
    $managementPackVersion = Get-ManagementPackVersion
    Log-Event -Arguments $Arguments -EventInfo $LogEntries.TSStarted -Parameters @($managementPackVersion)
    
    if($Arguments.TroubleshooterDisabled)
    {
        Log-Event -Arguments $Arguments -EventInfo $LogEntries.TroubleshooterDisabled -Parameters @($Arguments.Server)
        if ($MonitoringContext)
        {
            #PS 303295 - add a monitoring event to supress SCOM failure alerts.
            #
            Add-MonitoringEvent -Id $LogEntries.TsFailed[0] -Type $EVENT_TYPE_ERROR -Message $LocStrings.TroubleshooterDisabled
        }
    }
    else
    {
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
            $serverStatus = Detect-Problems `
                                -Server:$Arguments.Server `
                                -Database:$Arguments.Database `
                                -Symptom:$Arguments.Symptom 
        
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
    
        # The CI troubleshooter has successfully run. Reset the Troubleshooter failed retry counter
        #
        Reset-EventRetryCounter -Server $Arguments.Server -EventId $LogEntries.TSFailed[0]
        if ($MonitoringContext)
        {
            #PS 303295 - add a monitoring event to supress SCOM failure alerts.
            #
            Add-MonitoringEvent -Id $LogEntries.TSSuccess[0] -Type $EVENT_TYPE_INFORMATION -Message $LocStrings.TSStarted
        }
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
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQURnF/AtQ6FzWwW1t99ddzj8yI
# 8bmgghhqMIIE2jCCA8KgAwIBAgITMwAAASIn72vt4vugowAAAAABIjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzQw
# WhcNMjAwMTEwMjEwNzQwWjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046RUFDRS1FMzE2LUM5MUQxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQDRU0uHr30XLzP0oZTW7fCdslb6OXTQeoCa/IpqzTgDcXyf
# EqI0fdvhtqQ84neZE4vwJUAbJ2S+ajJirzzZIEU/JiTZpJgeeAMtN+MuAbzXrdyS
# ohyUDuGkuN+vVSeCnEZkeGcFf/zrNWWXmS7JsVK2BJR8YvXk0sBUbWVpdj0uvz68
# Y+HUyx8AKKE2nHRu54f6fC4eiwP/hs+L7NejJm+sNo7HXV4Y6edQI36FdY0Sotq8
# 7Lh3U96U4O6X9cD0iqKxr4lxYYkh98AzVUjiiSdWUt65DAMbdjBV6cepatwVVoET
# EtNK/f83bMS3sOL00QMWoyQM1F7+fLoz1TF7qlozAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUtlFVlkkUKuXnuF3JZxfDlHs2paYwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAZldsd7vjji5U30Zj
# cKiJvhDtcmx0b4s4s0E7gd8Lp4VnvAQAnpc3SkknUslyknvHGE77OSdxKdrO8qnn
# T0Tymqvf7/Re2xJcRVcM4f8TeE5hCaffCkB7Gtu90R+6+Eb1BnBDYMbj3b42Jq8K
# 42hnDG0ntrgv4/TmyJWIvmGQORWMCWyM/NraY3Ldi7pDpTfx9Z9s4eNE/cxipoST
# XHMIgPgDgbZcuFBANnWwF+/swj69cv87x+Jv/8HM/Naoawrr8+0yDjiJ90OzLGI5
# RScuGfQUlH0ESbzevO/9PFpoUywmNYhHoEPngLJVT2W6y13jFUx3IS9lnR0r1dCh
# mynB8jCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQULBqos8mcDlDW7xEW+qnTFH+5m3Iw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAIcoDE2jDE25eDsvIj95ln0HUy++GwGGzgUSMBHzBi5u
# kfzcmGj4nev9yNQHDWgd1eas96J99Kq9kQhvUon/YWA42nsNVPCVR3+aNb4dKTeu
# lFBNGpaapV3jT3nqxI9RkX+EqgaDDrf0W5ETtaHwfXpgTCMEWvHETT3GYf3fPTX+
# tJhMShTkBFKsKNjXhBcCelgyJP4T9wIBNTintPwkKTYwlvWyYEwaOvC1jSOpZolT
# WJNUXNfS4k3r7VZM5Q2bkZ7W0r6ehLR+U+Y72Vf31yyJ4git/DHE4bRpNk4UsjpK
# FSEOdy8B27UUCcP3ClYR8aEFBalJbRV3Pf4q3k7VGayhggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# Iifva+3i+6CjAAAAAAEiMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI1NThaMCMGCSqGSIb3DQEJ
# BDEWBBSEX9lO3AMMZp3VNYuoEcdRYE//FzANBgkqhkiG9w0BAQUFAASCAQDICrAY
# VfRRZ60vQzAKpokisczzMgjmYuM6wqFRgb0goam6SGDXbuyU3TI1qQhb3fJvFYuA
# rpfGhsNLPEbWaPJ6txjLK+ED4ahum+OOdNGafbNk84yQsA7nM4Nnrj0nsLHlE/d/
# fA7xs75ZGScd0OhybQMqAL7+2U7CmfTMHZDECSqjSXa+dzow3fKlnXmkLPD0u34Y
# 6m8Hja5buJu8NzaLZ6+XdX147+SEHfINd9LgY3OF8078ukC9/jIqrRA7LQvFZh4i
# x+Si2kc2ZazPf2teVKg0UBrA7uWb3HYvZlnSNJcxF2CYKlV8/fgIJyou2Et+MvXX
# hsV9xnJm9VxuW7SQ
# SIG # End signature block
