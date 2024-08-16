# Copyright (c) 2005 Microsoft Corporation.  All rights reserved.
#
# InstallWindowsComponent.ps1
#
# This script calls Add-WindowsFeature cmdlet to add Windows components for specific role.

########################
##  Input parameters  ##
########################
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "None", DefaultParameterSetName="Default")]
param
(
    # TEST is for testing purpose only
    # AdminTools, ClientAccess and MailBox are external server role names that are used by customers.
    # The rest are internal server role names that are used by Exchange internally.
    [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
    [ValidateSet("TEST", "AdminTools", "ClientAccess", "MailBox", "Internal-HUB", "Internal-CAS", "Edge", "Internal-MBX", "Internal-UM", "Internal-CAFE")]
    [string[]]
    $ShortNameForRole = "AdminTools",

    [Parameter(Position = 1, Mandatory = $false, ValueFromPipeline = $true)]
    [bool]
    $ADToolsNeeded = $false,

    # This is for testing purpose only
    [Parameter(Position = 2, Mandatory = $false, ValueFromPipeline = $true)]
    [string[]]
    $TestFeature = @()
)

############################
## Script level variables ##
############################
# role name mapping from external role names to internal role names.
$script:roles = @{}
$script:ClientAccessRoleName = "ClientAccess"
$script:MailBoxRoleName      = "MailBox"
$script:roles = @{}
$script:roles[$script:ClientAccessRoleName]  = @("Internal-HUB", "Internal-CAFE")
$script:roles[$script:MailBoxRoleName]       = @("Internal-HUB", "Internal-CAS", "Internal-MBX", "Internal-UM")

# Features to install if $ADToolsNeeded is $true
$script:featureForADTools = @( "RSAT-ADDS-Tools" )

# Features to install per role. Valid roles are validated via Param ValidateSet.
$script:features = @{}

$script:features['TEST']         = $TestFeature

$script:features['AdminTools']   = @( 'NET-Framework',
                                      'Server-Gui-Mgmt-Infra',
                                      'Web-Mgmt-Console',
                                      'Web-Lgcy-Mgmt-Console',
                                      'Web-Metabase',
                                      'Windows-Identity-Foundation' )

$script:features['Internal-HUB'] = @( 'NET-Framework',
                                      'Server-Gui-Mgmt-Infra',
                                      'Web-Server',
                                      'Web-Basic-Auth',
                                      'Web-Windows-Auth',
                                      'Web-Metabase',
                                      'Web-Net-Ext',
                                      'Web-Lgcy-Mgmt-Console',
                                      'WAS-Process-Model',
                                      'RSAT-Web-Server',
                                      'Web-Mgmt-Service',
                                      'Windows-Identity-Foundation' )

$script:features['Internal-CAS'] = @( 'NET-Framework'
                                      'Server-Gui-Mgmt-Infra',
                                      'Web-Server',
                                      'Web-Basic-Auth',
                                      'Web-Windows-Auth',
                                      'Web-Metabase',
                                      'Web-Net-Ext',
                                      'Web-Lgcy-Mgmt-Console',
                                      'WAS-Process-Model',
                                      'RSAT-Clustering',
                                      'RSAT-Clustering-Mgmt',
                                      'RSAT-Clustering-PowerShell',
                                      'RSAT-Clustering-CmdInterface',
                                      'RSAT-Web-Server',
                                      'Desktop-Experience',
                                      'Web-ISAPI-Ext',
                                      'Web-Digest-Auth',
                                      'Web-Dyn-Compression',
                                      'Web-Stat-Compression',
                                      'Web-WMI',
                                      'Web-Asp-Net',
                                      'Web-ISAPI-Filter',
                                      'Web-Client-Auth',
                                      'Web-Dir-Browsing',
                                      'Web-Http-Errors',
                                      'Web-Http-Logging',
                                      'Web-Http-Redirect',
                                      'Web-Http-Tracing',
                                      'Web-Request-Monitor',
                                      'Web-Static-Content',
                                      'NET-HTTP-Activation',
                                      'RPC-over-HTTP-proxy',
                                      'Web-Mgmt-Service',
                                      'Windows-Identity-Foundation' )

$script:features['Edge']         = @( 'NET-Framework',
                                      'Server-Gui-Mgmt-Infra',
                                      'ADLDS' )

$script:features['Internal-MBX'] = @( 'NET-Framework',
                                      'Server-Gui-Mgmt-Infra',
                                      'Web-Server',
                                      'Web-Basic-Auth',
                                      'Web-Windows-Auth',
                                      'Web-Metabase',
                                      'Web-Net-Ext',
                                      'Web-Lgcy-Mgmt-Console',
                                      'WAS-Process-Model',
                                      'RSAT-Clustering',
                                      'RSAT-Clustering-Mgmt',
                                      'RSAT-Clustering-PowerShell',
                                      'RSAT-Clustering-CmdInterface',
                                      'RSAT-Web-Server',
                                      'Web-Mgmt-Service',
                                      'Windows-Identity-Foundation',
                                      'NET-Non-HTTP-Activ' )

$script:features['Internal-UM']  = @( 'NET-Framework',
                                      'Server-Gui-Mgmt-Infra',
                                      'Web-Server',
                                      'Web-Basic-Auth',
                                      'Web-Windows-Auth',
                                      'Web-Metabase',
                                      'Web-Net-Ext',
                                      'Web-Lgcy-Mgmt-Console',
                                      'WAS-Process-Model',
                                      'RSAT-Web-Server',
                                      'Desktop-Experience',
                                      'Windows-Identity-Foundation' )

$script:features['Internal-CAFE']= @( 'NET-Framework'
                                      'Server-Gui-Mgmt-Infra',
                                      'Web-Server',
                                      'Web-Basic-Auth',
                                      'Web-Windows-Auth',
                                      'Web-Metabase',
                                      'Web-Net-Ext',
                                      'Web-Lgcy-Mgmt-Console',
                                      'WAS-Process-Model',
                                      'RSAT-Clustering',
                                      'RSAT-Clustering-Mgmt',
                                      'RSAT-Clustering-PowerShell',
                                      'RSAT-Clustering-CmdInterface',
                                      'RSAT-Web-Server',
                                      'Desktop-Experience',
                                      'Web-ISAPI-Ext',
                                      'Web-Digest-Auth',
                                      'Web-Dyn-Compression',
                                      'Web-Stat-Compression',
                                      'Web-WMI',
                                      'Web-Asp-Net',
                                      'Web-ISAPI-Filter',
                                      'Web-Client-Auth',
                                      'Web-Dir-Browsing',
                                      'Web-Http-Errors',
                                      'Web-Http-Logging',
                                      'Web-Http-Redirect',
                                      'Web-Http-Tracing',
                                      'Web-Request-Monitor',
                                      'Web-Static-Content',
                                      'NET-HTTP-Activation',
                                      'RPC-over-HTTP-proxy',
                                      'Web-Mgmt-Service',
                                      'Windows-Identity-Foundation' )

# The corresponding components in Windows 8
$featuresWin8Substitutes = @{	'NET-Framework' = 'NET-Framework-45-Features';
                                'Web-Net-Ext' = 'Web-Net-Ext45';
                                'RSAT-Web-Server' = 'Web-Mgmt-Console';
                                'Web-Asp-Net' = 'Web-Asp-Net45';
                                'NET-HTTP-Activation' = 'NET-WCF-HTTP-Activation45';
                                'NET-Non-HTTP-Activ' = 'NET-WCF-Pipe-Activation45';
                            }

# Logging variables
$script:logDir = "$env:SYSTEMDRIVE\ExchangeSetupLogs"
$script:mainLogFile = "$script:logDir\" + "InstallWindowsComponent.log"
$script:roleLogFile = "$script:logDir\" + $script:ShortNameForRole + "Prereqs.log"

# Variables to assist to log everything into the file
$script:errorVariable = $null
$script:warningVariable = $null

# Backup environment variables
$script:backupOFS = $null

# Result of InstallWindowsFeature
$script:installWindowsFeatureResult = $null

#######################
## Exception Handler ##
#######################
# Trap exceptions for terminating errors
Trap
{
    Log-Exception $_.Exception

    Restore-EnvironmentVariable

    # stop and rethrow
    break;
}

####################
## Log functions  ##
####################
#Log as info
function Log-Info
{
    $entry = $Args[0]
    
    if($entry)
    {
        if ($script:writeToExchangeSetupLog -eq $true)
        {
            $line = "[{0}] {1}" -F "InstallWindowsComponent.ps1", $entry
            Write-ExchangeSetupLog -Info $line
        }
        else
        {
            $line = Format-Entry "Info" $entry
            Log-File $line
            Write-Verbose $line
        }
    }
}

#Log as warning
function Log-Warning
{
    $entry = $Args[0]
    
    if($entry)
    {
        if ($script:writeToExchangeSetupLog -eq $true)
        {
            Write-ExchangeSetupLog -Warning $entry
        }
        else
        {
            $line = Format-Entry "Warning" $entry
            Log-File $line
            Write-Warning $line
        }
    }
}

# Log exception as error
function Log-Exception
{
    $entry = $Args[0]

    if($entry -and ($entry -is [System.Exception]))
    {
        if ($script:writeToExchangeSetupLog -eq $true)
        {
            Write-ExchangeSetupLog -Error $entry
        }
        else
        {
            $line = Format-Entry "Error" $entry.ToString()
            Log-File $line
            Write-Error $line
        }
    }
}

# Log warning or error as warnings
function Log-WarningErrorVariable
{
    if (Check-WarningVariable)
    {
        Log-Warning "warningVariable: $($script:warningVariable)"
    }

    if (Check-ErrorVariable)
    {
        Log-Warning "errorVariable: $($script:errorVariable)"
    }
}

# Log result of AddWindowsFeature
function Log-AddWindowsFeatureResult($resultToLog)
{
    if ($resultToLog -ne $null)
    {
        Log-Info "Success: `"$($result.Success)`""
        Log-Info "RestartNeeded: `"$($result.RestartNeeded)`""
        Log-Info "ExitCode: `"$($result.ExitCode)`""
        Log-Info "Totally $($result.FeatureResult.Length) feature(s) were added: `"$($result.FeatureResult)`""
    }
    else
    {
        Log-Info "AddWindowsFeature result is null"
    }
}

# Write a line to file
function Log-File([string] $line)
{
    Add-Content -Path $script:mainLogFile -Value $line
}

# Format log entry
function Format-Entry([string]$type, [string] $entry)
{
    return "[{0}] [{1}] {2}" -F $(get-date).ToString("HH:mm:ss"), $type, $entry
}

######################
## Helper functions ##
######################
# Install windows feature per role
function Install-WindowsFeaturePerRole
{
    $result = $null

    Log-Info "Enter Install-WindowsFeaturePerRole"

    $allRoleNames = Get-AllRoleNames

    $allFeatures = Get-AllFeatures $allRoleNames

    $validFeatureNames = Get-ValidFeatureNames $allFeatures

    if ($validFeatureNames -eq $null)
    {
        Log-Info "The valid feature name array is null. Nothing to install."
    }
    elseif ($validFeatureNames.Count -eq 0)
    {
        Log-Info "The valid feature name array is emtpy. Nothing to install."
    }
    else
    {
        Log-Info "Totally $($validFeatureNames.Count) valid feature(s) to be installed: `"$($validFeatureNames)`""
        Log-Info "Begin installing features for internal server role(s): `"$allRoleNames`""
                
        $result = Add-WindowsFeature $validFeatureNames -logPath $script:roleLogFile -ErrorVariable script:errorVariable -WarningVariable script:warningVariable -ErrorAction SilentlyContinue
        Log-WarningErrorVariable
        Log-AddWindowsFeatureResult $result
      
        Log-Info "End installing features for internal server role(s) `"$allRoleNames`""
    }

    Log-Info "Exit Install-WindowsFeaturePerRole"

    return $result
}

# Add string if array does not contain it yet
function Add-UniqueStringToArray([REF][string[]] $all, [string] $unique)
{
    if ($all -eq $null)
    {
        Log-Info "Add-UniqueStringToArray: array is null"
        return
    }

    if ($unique -eq $null)
    {
        Log-Info "Add-UniqueStringToArray: string is null"
        return
    }

    if ($all.Value -inotcontains $unique)
    {
        $all.Value += $unique
    }
    else
    {
        Log-Info "Add-UniqueStringToArray: `"$unique`" is already in the array: `"$($all.Value)`""
    }
}

# Get the all role names
function Get-AllRoleNames
{
    [string[]] $allRoleNames = @()

    Log-Info "Get-AllRoleNames: Totally $($ShortNameForRole.Count) server role(s) as parameter: `"$($ShortNameForRole)`""

    foreach ($roleName in $ShortNameForRole)
    {
        if ($roleName -ieq $script:ClientAccessRoleName)
        {
            Log-Info "Get-AllRoleNames: For role `"$roleName`", adding internal server role(s): `"$($script:roles[$script:ClientAccessRoleName])`""

            foreach($name in $script:roles[$script:ClientAccessRoleName])
            {
                 Add-UniqueStringToArray ([REF]$allRoleNames) $name
            }
        }
        elseif ($roleName -ieq $script:MailBoxRoleName)
        {
            Log-Info "Get-AllRoleNames: For role `"$roleName`", adding internal server role(s): `"$($script:roles[$script:MailBoxRoleName])`""

            foreach($name in $script:roles[$script:MailBoxRoleName])
            {
                 Add-UniqueStringToArray ([REF]$allRoleNames) $name
            }
        }
        else
        {
            Log-Info "Get-AllRoleNames: Adding internal role `"$roleName`""

            Add-UniqueStringToArray ([REF]$allRoleNames) $roleName
        }
    }

    Log-Info "Get-AllRoleNames: Totally $($allRoleNames.Count) unique internal server role(s): `"$($allRoleNames)`""

    return $allRoleNames
}

# Get the all feature names
function Get-AllFeatures([string[]] $roleNames)
{
    if ($roleNames -eq $null)
    {
        Log-Info "Get-AllFeatures: the role name array is null."
        return $null
    }
    if ($roleNames.Count -eq 0)
    {
        Log-Info "Get-AllFeatures: the role name array is empty."
        return $null
    }

    [string[]] $allFeatures = @()

    foreach ($roleName in $roleNames)
    {
        Log-Info "Get-AllFeatures: For role `"$roleName`", adding features: `"$($script:features[$roleName])`""

        foreach($feature in $script:features[$roleName])
        {
            Add-UniqueStringToArray ([REF]$allFeatures) $feature
        }
    }

    if ($ADToolsNeeded)
    {
        Log-Info "Get-AllFeatures: adding featureForADTools: `"$script:featureForADTools`""
        $allFeatures += $script:featureForADTools
    }

    if ($allFeatures -eq $null)
    {
        Log-Info "Get-AllFeatures: the feature array is null."
    }
    else
    {
        Log-Info "Get-AllFeatures: Totally $($allFeatures.Count) feature(s) originally requested: `"$($allFeatures)`""

        if (Is-OSWindows8)
        {
            for ($i = 0; $i -lt $allFeatures.Count; $i++)
            {
                $key = $allFeatures[$i]

                if ($featuresWin8Substitutes.ContainsKey($key))
                {
                    $allFeatures[$i] = $featuresWin8Substitutes[$key]
                }
            }
        }
    }

    return $allFeatures
}

# Get the valid feature names
function Get-ValidFeatureNames($allFeatures)
{
    if ($allFeatures -eq $null)
    {
        Log-Info "Get-ValidFeatureNames: the feature name array is null."
        return $null
    }
    if ($allFeatures.Count -eq 0)
    {
        Log-Info "Get-ValidFeatureNames: the feature name array is empty."
        return $null
    }

    Log-Info "Get-ValidFeatureNames: Totally $($allFeatures.Count) feature(s) finally requested: `"$($allFeatures)`""

    $validFeatures = Get-WindowsFeature $allFeatures

    $validFeatureNames = @()
    foreach ($feature in $validFeatures)
    {
        $validFeatureNames += $feature.Name
    }
    
    return $validFeatureNames
}

# Is the OS version Windows8
function Is-OSWindows8
{
    $version = [System.Environment]::OSVersion.Version

    return ($version.Major -eq 6 -and $version.Minor -eq 2)
}

# Test if Write-ExchangeSetupLog is available
function Test-WriteExchangeSetupLog
{
    $checkCommand = Get-Command Write-ExchangeSetupLog -ErrorAction SilentlyContinue
    return ($checkCommand -ne $null)
}

# Check if there is any warning in the script:warningVariable
function Check-WarningVariable
{
    return ($script:warningVariable -ne $null -or $script:warningVariable.Count -gt 0)
}

# Check if there is any error in the script:errorVariable
function Check-ErrorVariable
{
    return ($script:errorVariable -ne $null -and $script:errorVariable.Count -gt 0)
}

# Backup and set powershell environment variable
function BackupSet-EnvironmentVariable
{
    # Separator for array fields
    $script:backupOFS = $script:OFS
    $script:OFS = ", "
}

# Restore powershell environment variable
function Restore-EnvironmentVariable
{
    $script:OFS = $script:backupOFS
}

########################
## Script starts here ##
########################
# Command Write-ExchangeSetupLog may not available if this is run from regular powershell envrionment.
# So we could log stuff into a separate file.
# Must set this flag before everything else.
$script:writeToExchangeSetupLog = Test-WriteExchangeSetupLog

BackupSet-EnvironmentVariable

# If log folder doesn't exist, create it
if (!(Test-Path $logDir))
{
    New-Item $logDir -type directory	
}

Log-Info "-----------------------------------------------"
Log-Info ("InstallWindowsComponent: {0}" -F $(get-date) )
Log-Info "Start InstallWindowsComponent with options: -ShortNameForRole $ShortNameForRole -ADToolsNeeded $ADToolsNeeded -TestFeature $TestFeature"

# Need to import ServerManager module as it's not loaded by default
Log-Info "Import-Module ServerManager"
Import-Module ServerManager -ErrorVariable script:errorVariable -WarningVariable script:warningVariable -ErrorAction SilentlyContinue
Log-WarningErrorVariable

# Install windows feature only if there is no error so far
$script:ret = -not(Check-ErrorVariable)
if ($script:ret -eq $true)
{
    Log-Info "Call AddWindowsFeature to do the real work"
    $script:installWindowsFeatureResult = Install-WindowsFeaturePerRole
    $script:ret = ($script:installWindowsFeatureResult -eq $null) -or ($script:installWindowsFeatureResult.Success -eq $true)
}

Restore-EnvironmentVariable

if ($script:ret -eq $true)
{
    Log-Info "End InstallWindowsComponent: script completed succesfully."
}
else
{
    # Log as warning on failures.
    # This is to keep the same behavior as it did in the Install-WindowsComponent cmdlet, where failures are ignored.
    # Setup prereq checks will report missing components as errors.
    $moreInfo = ""
    if (Test-Path  $script:roleLogFile)
    {
        $moreInfo = "Check $script:roleLogFile for more info."
    }

    Log-Warning "End InstallWindowsComponent: script completed with one or more errors. $moreInfo"
}

return $script:installWindowsFeatureResult
# SIG # Begin signature block
# MIIdtgYJKoZIhvcNAQcCoIIdpzCCHaMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUwL+WtVAZ9ejpfbWNigNUtWLw
# VougghhkMIIEwzCCA6ugAwIBAgITMwAAAJvgdDfLPU2NLgAAAAAAmzANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBLwwggS4AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB0DAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUh31hAcH80/I/pHw7AXxF3q80w28wcAYKKwYB
# BAGCNwIBDDFiMGCgOIA2AEkAbgBzAHQAYQBsAGwAVwBpAG4AZABvAHcAcwBDAG8A
# bQBwAG8AbgBlAG4AdAAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNv
# bS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAbBjqY0rq8hBbCzFff9uS6W3U
# 2b4TbkYtrDSsdrh6+P1XNlJ5eIWkPeFrupmKgPR3WJ5i/Va26MdRJllLbICvf/GA
# aGzyf2pr4U0cu1Pa4jB9FaYkWKZ0E3rCilW38KU+XuN4SqdmQfzZQghHCtcbgbLr
# 3oVgJ60m0RNzIW3WYHdagqF97uW7DNDrcGPlo5pCqILmk2FxPhSCd3YNkbomcAIk
# qSQr2mUqjcOmVJhZZlgN2xQq5RECe/7SbfUFsimCb8FxGdPJmPvIkE5gIYzgcnbf
# 1RcMEi2ZGF/2mPrjn/cg2YqD/RBhYxoGAeLBnf+b72oGk5lQ6x3rQDj7Kq34VqGC
# AigwggIkBgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1T
# dGFtcCBQQ0ECEzMAAACb4HQ3yz1NjS4AAAAAAJswCQYFKw4DAhoFAKBdMBgGCSqG
# SIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE2MDkwMzE4NDQ1
# NlowIwYJKoZIhvcNAQkEMRYEFApIAwOwVlFz4cDYhS04o+pmQxlsMA0GCSqGSIb3
# DQEBBQUABIIBAG5nR7eLZmyudYeGBMRtAVbU5nHP7wr+kgnH4ohBoFCvabx02dpD
# B60vUH1D+C99ACtuIln3ntN0Rg+GsYEA1D9OqZknu7vr37WGpVvn4JaWTgwi0ZCA
# oFH9oKWEtBQihNQ79UNJrjv+4yqh3d2V3XIQ6+q9gsJLVLcnR1JXsUOgoicbRij4
# r7nLlp52V9i4KWM77eSLLidYAn+F6Ggc90sr7q2sMgddQiAhrvDZ9IuYVOgrPAUn
# ZUGG7EhPejKty9iDpHAOCwND0Ve/TWO2E6ZSRb8HmoSZqtYfYLsJWO7wQBdjZHSm
# q3Br8pa654M4b8QceuGzXrzcbNmLNPLXtGM=
# SIG # End signature block
