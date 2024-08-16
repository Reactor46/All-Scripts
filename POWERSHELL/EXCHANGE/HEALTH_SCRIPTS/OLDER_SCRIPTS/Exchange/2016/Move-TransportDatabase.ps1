<#
.EXTERNALHELP Move-TransportDatabase-help.xml
#>

# Copyright (c) Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

param(
	[Parameter(Mandatory = $false)]
	[String]$queueDatabasePath,
	[Parameter(Mandatory = $false)]
	[String]$queueDatabaseLoggingPath,
	[Parameter(Mandatory = $false)]
	[String]$iPFilterDatabasePath,
	[Parameter(Mandatory = $false)]
	[String]$iPFilterDatabaseLoggingPath,
	[Parameter(Mandatory = $false)]
	[String]$temporaryStoragePath,
	[Parameter(Mandatory = $false)]
	[bool]$setupMode
)

function LogInfo([String] $log)
{
    if ($script:setupMode -ne $true)
    {
        Write-Host $log
    }
    else
    {
        Write-ExchangeSetupLog -Info $log
    }
}

function LogWarning([String] $log)
{
    if ($script:setupMode -ne $true)
    {
        Write-Warning $log
    }
    else
    {
        Write-ExchangeSetupLog -Warning $log
    }
}

function LogError([String] $log)
{
    if ($script:setupMode -ne $true)
    {
        Write-Error $log
    }
    else
    {
        Write-ExchangeSetupLog -Error $log
    }
}

######################################################
# Retrieves the Root setup registry entry.
# returns: return entry value of found else null
######################################################
function GetEdgeInstallPath()
{
    # Get the root setup entires.
    $setupRegistryPath = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup"
    $setupEntries = get-itemproperty $setupRegistryPath
    if($setupEntries -eq $null)
    {
        return $null
    }

    # Try to get the Install Path.
    $installPath = $setupEntries.MsiInstallPath

    return $installPath
}

#############################################################################
# Initialized script variables and load from application config
#############################################################################
function InitializeVariablesAndLoadAppConfig()
{
    $logFiles = @("Trn*.log", "Trnres00001.jrs", "Trnres00002.jrs", "Temp.edb")
    $QD = @{Desc="Queue Database"; Config="QueueDatabasePath"; Input=$script:queueDatabasePath; NewPath=$null; OldPath=$null; Files=@("mail.que", "trn.chk")}
    $QL = @{Desc="Queue Database Logging"; Config="QueueDatabaseLoggingPath"; Input=$script:queueDatabaseLoggingPath; NewPath=$null; OldPath=$null;  Files=$logFiles}
    $ID = @{Desc="IP Filter Database"; Config="IPFilterDatabasePath"; Input=$script:iPFilterDatabasePath; NewPath=$null; OldPath=$null; Files=@("IPFiltering.edb", "trn.chk")}
    $IL = @{Desc="IP Filter Database Logging"; Config="IPFilterDatabaseLoggingPath"; Input=$script:ipFilterDatabaseLoggingPath; NewPath=$null; OldPath=$null; Files=$logFiles}
    $T = @{Desc="Temporary Storage"; Config="TemporaryStoragePath"; Input=$script:temporaryStoragePath; NewPath=$null; OldPath=$null; Files=@()}

    $script:table.Add("QD", $QD)
    $script:table.Add("QL", $QL)
    $script:table.Add("ID", $ID)
    $script:table.Add("IL", $IL)
    $script:table.Add("T", $T)

    # Global variable for storing handle to the clone log file.
    $installPath = GetEdgeInstallPath
    $script:exePath = [System.IO.Path]::Combine($installPath, "bin\EdgeTransport.exe")
    $script:config = [System.Configuration.ConfigurationManager]::OpenExeConfiguration($exePath)

    foreach ($entry in $script:table.Values)
    {
        $desc = $entry["Desc"]
        $oldPath = $script:config.AppSettings.Settings[$entry["Config"]].Value
        if ($oldPath -eq $null)
        {
            LogError ($MoveTransportDatabase_LocalizedStrings.res_0000 -f $desc)
            return $false
        }

        $entry["OldPath"] = [Microsoft.Exchange.Data.LocalLongFullPath]::Parse($oldPath)
    }

    return $true
}

#######################################################################
# Validates the input parameters.
# Returns True if validation succeeds else False.
#######################################################################
function ValidateInput()
{
    #return false if there is any exception in this function

    $newPath = $null
    $hasValidInput = $false
    foreach ($entry in $script:table.Values)
    {
        if ($entry["Input"] -ne "" -and $entry["Input"] -ne $null)
        {
            $desc = $entry["Desc"]
            $success = [Microsoft.Exchange.Data.LongPath]::TryParse($entry["Input"], [ref] $newPath)
            if ($success -eq $false)
            {
                LogError ($MoveTransportDatabase_LocalizedStrings.res_0001 -f $desc)
                return $false
            }

            if ($entry["OldPath"].Equals($newPath))
            {
                LogWarning ($MoveTransportDatabase_LocalizedStrings.res_0002 -f $desc)
            }
            else
            {
                #set NewPath
                $entry["NewPath"] = $newPath
                $oldPath = $entry["OldPath"]
                LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0003 -f $desc,$oldPath,$newPath)
            }
            $hasValidInput = $true
        }
    }

    #make sure we have at least one argument specified
    if ($hasValidInput -ne $true)
    {
        LogError $MoveTransportDatabase_LocalizedStrings.res_0004
        return $false
    }

    #make sure two database paths are not the same, since the log files will have the same file names
    $QDNew = $script:table["QD"]["NewPath"]
    $QDOld = $script:table["QD"]["OldPath"]
    $IDNew = $script:table["ID"]["NewPath"]
    $IDOld = $script:table["ID"]["OldPath"]

    if (($QDNew -ne $null -and ($QDNew.Equals($IDNew) -or $QDNew.Equals($IDOld))) -or
        ($IDNew -ne $null -and ($IDNew.Equals($QDNew) -or $IDNew.Equals($QDOld))))
    {
        LogError $MoveTransportDatabase_LocalizedStrings.res_0005
        return $false
    }

    #make sure two logging paths are not the same, since the log files will have the same file names
    $QLNew = $script:table["QL"]["NewPath"]
    $QLOld = $script:table["QL"]["OldPath"]
    $ILNew = $script:table["IL"]["NewPath"]
    $ILOld = $script:table["IL"]["OldPath"]

    if (($QLNew -ne $null -and ($QLNew.Equals($ILNew) -or $QLNew.Equals($ILOld))) -or
        ($ILNew -ne $null -and ($ILNew.Equals($QLNew) -or $ILNew.Equals($QLOld))))
    {
        LogError $MoveTransportDatabase_LocalizedStrings.res_0006
        return $false
    }

    #Compact the table to remove ignored entries
    $updatedTable = @{}

    foreach ($entry in $script:table.Keys)
    {
        if ($script:table[$entry]["NewPath"] -ne $null)
        {
            $updatedTable.Add($entry, $script:table[$entry])
        }
    }

    $script:table = $updatedTable

    return $true
}

#############################################################################
# Verify that the target disk, if different from the original, has enough disk space (2GB plus size required to move file)
# Returns True if it succeeds else False.
#############################################################################
function CheckDiskSize()
{
    foreach ($entry in $script:table.Values)
    {
        $desc = $entry["Desc"]
        $newDrive = $entry["NewPath"].DriveName
        $oldPath = $entry["OldPath"]

        if ($oldPath.DriveName -eq $newDrive)
        {
            LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0007 -f $desc)
            continue;
        }

        $driveInfo = new-object System.IO.DriveInfo ($newDrive)
        $freeSpace = $driveInfo.TotalFreeSpace

        #2GB, plus the space needed to move the file
        $minSpace = 2*1024*1024*1024
        foreach ($filePattern in $entry["Files"])
        {
            $fullFilePattern = [System.IO.Path]::Combine($oldPath, $filePattern)
            get-item $fullFilePattern -ErrorAction: SilentlyContinue | foreach {$minSpace += $_.Length}
        }

        LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0008 -f $minSpace,$newDrive,$freeSpace)

        if ($freeSpace -lt $minSpace)
        {
            LogError ($MoveTransportDatabase_LocalizedStrings.res_0009 -f $newDrive)
            return $false
        }
    }

    return $true
}

#############################################################################
# Create the folders and grant permissions.
# Returns True if it succeeds else False.
#############################################################################
function CreateFolderIfNecessary()
{
    foreach ($entry in $script:table.Values)
    {
        $newPath = $entry["NewPath"]
        $exists = [System.IO.Directory]::Exists($newPath)

        if (!$exists)
        {
            LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0010 -f $newPath)
            [System.IO.Directory]::CreateDirectory($newPath.ToString())
        }
        else
        {
            LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0011 -f $newPath)
        }

        $directorySecurity = get-acl $newPath

        #Do not inherit permission from parent
        $directorySecurity.SetAccessRuleProtection($true, $false)

        $authorizationRuleCollection = $directorySecurity.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])

        #Add permissions
        $wellKnownSids = @([System.Security.Principal.WellKnownSidType]::NetworkServiceSid,
            [System.Security.Principal.WellKnownSidType]::LocalSystemSid,
            [System.Security.Principal.WellKnownSidType]::BuiltinAdministratorsSid)

        foreach ($wellKnownSid in $wellKnownSids)
        {
            if (HasFullAccess $authorizationRuleCollection $wellKnownSid)
            {
                LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0012 -f $wellKnownSid)
            }
            else
            {
                LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0013 -f $wellKnownSid)
                $sid = new-object System.Security.Principal.SecurityIdentifier ($wellKnownSid, $null)
                $inheritanceFlags = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
                $newRule = new-object System.Security.AccessControl.FileSystemAccessRule ($sid, [System.Security.AccessControl.FileSystemRights]::FullControl, $inheritanceFlags, [System.Security.AccessControl.PropagationFlags]::None, [System.Security.AccessControl.AccessControlType]::Allow)
                $directorySecurity.AddAccessRule($newRule)
            }
        }

        set-acl $newPath $directorySecurity
    }

    return $true
}

#############################################################################
# Check whether the specified SID has full control in the authorization rule collection
# Returns True if it succeeds else False.
#############################################################################
function HasFullAccess([System.Security.AccessControl.AuthorizationRuleCollection] $acl, [System.Security.Principal.WellKnownSidType] $wellKnownSid)
{
    $hasFullControl = $false
    $hasDeny = $false
    #for each FileSystemAccessRule
    foreach ($rule in $acl)
    {
        $inheritanceFlags = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
        if ($rule.IdentityReference.IsWellKnown($wellKnownSid))
        {
            if ($rule.AccessControlType -eq [System.Security.AccessControl.AccessControlType]::Allow -and
                $rule.FileSystemRights -eq [System.Security.AccessControl.FileSystemRights]::FullControl -and
                $rule.InheritanceFlags -eq $inheritanceFlags -and
                !$rule.IsInherited)
            {
                $hasFullControl = $true
            }

            if ($rule.AccessControlType -eq [System.Security.AccessControl.AccessControlType]::Deny)
            {
                $hasDeny = $true
            }
        }
    }

    if ($hasDeny)
    {
        LogWarning ($MoveTransportDatabase_LocalizedStrings.res_0014 -f $wellKnownSid)
    }

    return $hasFullControl
}

#############################################################################
# Backup the EXE configuration file
# Returns True if it succeeds else False.
#############################################################################
function BackupConfigFile()
{
    $timeStr = (get-date).ToString("yyyyMMddHHmmss")
    $newConfigPath = $script:exePath+".config.$timeStr.old"

    LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0015 -f $newConfigPath)
    copy-item $script:config.FilePath $newConfigPath -ErrorAction: Stop

    return $true
}

#############################################################################
# Move the files to the new folders if necessary, and update configuration.
# Returns True if it succeeds else False.
#############################################################################
function MoveFilesAndUpdateConfig()
{
    foreach ($entry in $script:table.Values)
    {
        $desc = $entry["Desc"]
        $oldPath = $entry["oldPath"].ToString()
        $newPath = $entry["newPath"].ToString()

        #Move files if necessary
        foreach ($filePattern in $entry["Files"])
        {
            $fullFilePattern = [System.IO.Path]::Combine($oldPath, $filePattern)
            $found = $false

            get-item $fullFilePattern -ErrorAction: SilentlyContinue | foreach {
                $name = $_.Name;
                $found = $true;
                move-item $_ $newPath -force -ErrorAction: Stop;
                LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0016 -f $name)
            }

            if (!$found)
            {
                LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0017 -f $filePattern)
            }
        }

        $script:config.AppSettings.Settings[$entry["Config"]].Value = $newPath

        LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0018 -f $desc,$newPath)
    }
    $script:config.Save()
    return $true
}

#############################################################################
# Start or stop a service
# Returns True if it succeeds else False.
#############################################################################
function StartStopService([String] $serviceName, [Boolean] $start)
{
    $serviceStatus = get-service $serviceName

    $desiredStatus = [System.ServiceProcess.ServiceControllerStatus]::Stopped
    $desiredStatusString = "stopped"
    $actionString = "stop"

    if ($start)
    {
        $desiredStatus = [System.ServiceProcess.ServiceControllerStatus]::Running
        $desiredStatusString = "started"
        $actionString = "start"
    }

    if ($serviceStatus -eq $null)
    {
        LogError ($MoveTransportDatabase_LocalizedStrings.res_0019 -f $serviceName)
        return $false
    }
    elseif($serviceStatus.Status -eq $desiredStatus)
    {
        LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0020 -f $serviceName,$desiredStatusString)
        return $true
    }

    LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0021 -f $actionString,$serviceName)

    if ($start)
    {
        if ($script:setupMode -ne $true)
        {
            start-service -Name:$serviceName
        }
        else
        {
            Start-SetupService -ServiceName:$serviceName
        }
    }
    else
    {
        if ($script:setupMode -ne $true)
        {
            stop-service -Name:$serviceName
        }
        else
        {
            stop-SetupService -ServiceName:$serviceName
        }
    }

    $serviceStatus = get-service $serviceName

    if(($serviceStatus -eq $null) -or
       ($serviceStatus.Status -ne $desiredStatus))
    {
        LogError ($MoveTransportDatabase_LocalizedStrings.res_0022 -f $serviceName,$desiredStatusString)
        return $false
    }
    else
    {
        LogInfo ($MoveTransportDatabase_LocalizedStrings.res_0023 -f $serviceName,$desiredStatusString)
        return $true
    }
}

#############################################################################
# Global variables
#############################################################################

$table = @{}
$exePath = $null
$config = $null

#######################################################################
# Main Script starts here.
#######################################################################

#Make sure we trap the exception if any, and stop processing
trap
{
    break
}

#load hashtable of localized string
Import-LocalizedData -BindingVariable MoveTransportDatabase_LocalizedStrings -FileName Move-TransportDatabase.strings.psd1

if (!(InitializeVariablesAndLoadAppConfig))
{
    exit
}

if (!(ValidateInput))
{
    exit
}

if (!(CheckDiskSize))
{
    exit
}

if (!(CreateFolderIfNecessary))
{
    exit
}

if (!(StartStopService "MSExchangeTransport" $false))
{
    exit
}

if (!(BackupConfigFile))
{
    exit
}

if (!(MoveFilesAndUpdateConfig))
{
    exit
}

#Start transport
if (!(StartStopService "MSExchangeTransport" $true))
{
    exit
}

LogInfo $MoveTransportDatabase_LocalizedStrings.res_0024


# SIG # Begin signature block
# MIIdtAYJKoZIhvcNAQcCoIIdpTCCHaECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUZGKVisH7PT1synVSwouGqyUZ
# T/mgghhkMIIEwzCCA6ugAwIBAgITMwAAAJzu/hRVqV01UAAAAAAAnDANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBLowggS2AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBzjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUKUDj9nsUvjCx0lcdK/1FoilV2iYwbgYKKwYB
# BAGCNwIBDDFgMF6gNoA0AE0AbwB2AGUALQBUAHIAYQBuAHMAcABvAHIAdABEAGEA
# dABhAGIAYQBzAGUALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20v
# ZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBAAHV3KlI+VI9glLtBralT9angc6o
# otLDjU5TSz85cZIvB+0XuFBk1xvWzOCOdJgJqkLn8qQvR0gOpGMMhMYprq0VOFQW
# X/rG8Rz6IkzK2yNYqU+gI+rFmppOtFmvRh17m0ybjRchRe8td+8TvMcV4DyA2CWj
# ZjWBOm+uIVTM7dQmSOoH7qUBtBL6SEmAxxqDpp64ffgk5o2IWFe7JYZCldyIi4Z2
# BXbjimZ7PShBrV1l8KrbiJyHp9FuhHK2PMkt4V4IVKgrmeeBzDbC6j1RiVgZOONR
# 33LyQ7zpwazal4Qh8aSTRwpMAWW+uuTpCpfPMpQ4OVF1x78HpSFiUWqTjY+hggIo
# MIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBAhMzAAAAnO7+FFWpXTVQAAAAAACcMAkGBSsOAwIaBQCgXTAYBgkqhkiG
# 9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMxODQ1MTha
# MCMGCSqGSIb3DQEJBDEWBBRI0wMAvJzDUGbAxM40ljL2+D1jSjANBgkqhkiG9w0B
# AQUFAASCAQCn5uVkLj56HoeztKq0ePvo8WXla5sn/gYFG3sw4epD3KAW7SgglAVE
# 6GI/HBR8eC76REgfLo8sYSuC1U4zoVP1WbizUEkFX2jrzO5bt5IlWOBJaCWfeZzw
# /RGDjMVYQWfDtVXL6S5K9xelWB+njjW4cVa/XhJN92dn0e7KnpFdGmy3IZ09biF/
# UicDYrJJ2irVZv2uDLzTC4eLXv3t3VsK5Rft50pBD7lLd+YwtmdUqirmkQMf1+sQ
# yP9Ci5sH1udnwsU8QX2H4WgkbXoGtXi/dRJbvZFxYxpOkJX72yMVtLnuz4KhKmD0
# cLiQpWCmQo337r5y5Y9mmC+IVoErOOjA
# SIG # End signature block
