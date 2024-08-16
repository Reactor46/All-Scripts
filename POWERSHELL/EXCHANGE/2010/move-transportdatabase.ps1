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
    $setupRegistryPath = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup"
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
# MIIaegYJKoZIhvcNAQcCoIIaazCCGmcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUaS46rxJtHiNCr3zqnASxPpO8
# 0gigghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# a8kTo/0xggS1MIIEsQIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIHOMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBTeKq6rr8w0IL8TzhfneFvc3p67hzBuBgorBgEEAYI3AgEMMWAw
# XqA2gDQATQBvAHYAZQAtAFQAcgBhAG4AcwBwAG8AcgB0AEQAYQB0AGEAYgBhAHMA
# ZQAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAw
# DQYJKoZIhvcNAQEBBQAEggEAlhz3YUevZm6H6fryU11lWXnLn/9EmaoV1G7RO//8
# gohjwUHhlyvXSAvYdN2NXModkUrNVZZ+8U9FFe75I9rM7t4dS8jy3weiKEtD7XyE
# KOPACfSL43Yauk7tOLDHc/O2AM+b0zkxF3ycEhkU/qLfdfipoevTOV0XLzpO+oE+
# MkHM7+wgPgXgySZ2sv7DCrmZFoOdhAnucxr4oKd/1efkb/B1equdjJI+6AEXXyS4
# ENsEMxIHxl7AEjAgl570pWufbScwSNmFZ9Pl5VGBw9cq1GlzECvnOq+mHJI4YRiX
# BNkr9SsHK4341vf3gt3416iLwnzEmEpNdsdhqaA+nWKHyaGCAigwggIkBgkqhkiG
# 9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMA
# AAArOTJIwbLJSPMAAAAAACswCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkq
# hkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTEzMDIwNTA2MzcyNFowIwYJKoZIhvcN
# AQkEMRYEFP0peQq9AbTnVIsZ07v90FOlXvw4MA0GCSqGSIb3DQEBBQUABIIBAA3U
# ZodGnmWBybqHTMjC4oL7XNN6w9uAgNtdxeQ+FxkcW3j5FVwZZWQ0xvjVqEypmLam
# 5DukU29/uE4mqi9Sx+y8i1xY80GYrXRuUoAN3Zk6B9lOv+X7cMn21SqBw+RJhJcm
# wIe0fQULRv9PajzvJHo7P9XvcODNhmRmb8Td+lZU5ol8x6XrSzMla7xn9DQ2beay
# r0k8tB0XCmUkmTA3TT8lH8RmT7Z0l+rt5XQAU+h/3u76PP3x821zq2dj2tXAKKYp
# UFDFj0f/ro8iDzZPL63bN7QoC4XkyTPE7sR81SQ5go7NA/2YAjWaEL65I6nNcBRi
# JxacGQrcCeWeAzw3Ds0=
# SIG # End signature block
