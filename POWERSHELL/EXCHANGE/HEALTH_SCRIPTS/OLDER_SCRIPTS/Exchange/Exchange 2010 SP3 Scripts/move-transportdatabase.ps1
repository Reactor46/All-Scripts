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
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUaS46rxJtHiNCr3zqnASxPpO8
# 0gigghhqMIIE2jCCA8KgAwIBAgITMwAAAR+XYwozuYPXKwAAAAABHzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzM3
# WhcNMjAwMTEwMjEwNzM3WjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046NDlCQy1FMzdBLTIzM0MxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQCppklVnT29zi13dODY0ejMsdoe7n2iCvC6QdH5FJkRYfy+
# cXoHBmpDgDF/65Kt9GMmu/K8HKAzjKHeG18rgRXQagLwIIH5yCRbXGwOfuHIu1dC
# 26o/CT22+YlRvBJwH36WVjML8BLNDT3Fr+yhU4ZM7Hbegql4r5kSgsrrjyx5bJY5
# r2N0G7RDnbhRd79iqXbvDnvkatjB5xgluzfQEAPbJjXjmRb5685DEEZg1qFsQJer
# XuBA+ZVevuCX0DuDj8UmhHGC5Y32sulFTn283R6LU+8+AALtbHOOIHV7QHNYV8mN
# jxHuKLvE9tNEGIpbG2WF2yQkSGe3sRbGQmaILWeHAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUuPNVyPmK8/JJioMtQFlTUeF3IOgwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAmAYfr1fEosYv9VTf
# 0Msya6aFm0Id6Zq1O5jNy74ByTh7EEac/l/4e3DOyrczHS6zwvMKYzLtmifeGZvD
# 70qbbUfF+yjpzpyu00uuzZ1HNOpktp5/dJXkzz0NyVnEeFGOXLpNyZNIA9dKGDwN
# XbsEUukTX9lJFx5RcBhE8AOl22IHSgJ6NYf4DpATCjSJbC9IrKYGBchHobCLZHEt
# cLBjxXiWJRG2YY+LBAVW95gwNdPmLCKrob7SdNLK1VnM35Q2VgNF7YfDc5nw4E7C
# 4VaZvlyuDET6fYycIVPx5GsLhx3it4a+WKcBwarK7inH9skUArxMZrpWmjuQ/o4b
# GprEnjCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU3iquq6/MNCC/E84X53hb3N6eu4cw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAB9QA6RSwvQnKE63puFcemA10HreJ7CfHgzR6eusq3q/
# QlWVCuXuxDPuvkH7m3DPhgeghCAMWXiK5jDJMcrOQN0QZ7AKFXoq525TFnfznpPN
# y/dXR8uxA6mNpqieGbChwBHXV2bFT9GJJllrnmuu9i9wRsHRhKNhHYuN10wujTi0
# yQWacnPMUNTMFsVHBNxGvC3oqdKru6gx8qVz/ynHQmmTrTn/sY3vZRfBrMClhezs
# 5BPhjIj3fKjUXiJDnQePla7lbMq55NWS2JoRn3icTjA49oF3/w5zEMPm2+ws17EW
# Lo+QGJI20dnuZHolMkdErUZXRASQjftw2jUgiAvFiwOhggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# H5djCjO5g9crAAAAAAEfMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI2MDZaMCMGCSqGSIb3DQEJ
# BDEWBBQiuthjzGq0SdfItN/WPiap4LFm7TANBgkqhkiG9w0BAQUFAASCAQAHXj8+
# vECU9PPiFoYYN85cjtZGlczoHt4Ym062QvtZPQb7EWY+Lrjmt7bwlURB5rZqR25I
# zhcvyAHraS2sx9gcF/ctg4Elmc0UmLU9+rNdht+5M4SE7ZVb2dXuL1cPCt/N4pTY
# GHXAwOtR5WRwwiEQS4E3weMt4Bfp/HTasWeQ9DBucE1yYvEli3Dr34GV4gGHb+1Z
# rULlTAh3wEcNBQjEUu70zg4D1xIMTOyJ2gSg7rNfWYIoUxGUbContA6ggEhK8mZc
# ivAzqXd5jMuNf9NWncvAD2kUzEVUKN9rbGCJ2gm0Ivvs1gvfeZmBNyTq+ppSI/bB
# 45MSa2lQDe2RnKDY
# SIG # End signature block
