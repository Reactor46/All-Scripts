#############################################################################################################
#.SYNOPSIS
#   Updates client permissions of several users to a public folder
#
#.DESCRIPTION
#   Updates the client permissions of a public folder (and its children if -recurse
#	is provided) clearing the permissions a set of users have on the folder and setting
#	the provided access rights
#
#	Copyright (c) 2014 Microsoft Corporation. All rights reserved.
#
#	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
#	OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
#.PARAMETER  IncludeFolders
#   Identities of the Public Folders that will be updated
#
#.PARAMETER  Users
#   List of users whose current access rights to the folder will be overriten
#
#.PARAMETER  AccessRights
#   List of permissions to assign to the users
#
#.PARAMETER  Recurse
#   If provided the permission changes will also be applied to the children of the folders.
#
#.PARAMETER  ExcludeFolderEntryIds
#   List of EntryIds of the folders that should be ignored from the update. Notice however
#   that if you use the Recurse option the children of these folders won't be ignored unless
#   their EntryIds are also provided in this list.
#
#.PARAMETER  SkipCurrentAccessCheck
#   If provided the right access updates will be performed in the folder regardless of whether
#   the current folder has the same permissions already applied.
#
#.PARAMETER  Confirm
#   If this switch parameter is set to $false all operations on the public folder will be 
#   performed without requesting confirmation from the user.
#
#.PARAMETER  WhatIf
#   If this switch parameter is present the operations on the public folder will not be 
#   performed but information on what task would be performed are printed to the console.
#
#.PARAMETER  ProgressLogFile
#   File to log EntryIds of folders that were successfully updated. The content of this file may
#   become handy to save time if the previous execution of the script was aborted and you want to restart
#   from the point the script stopped. To do this simply get the contents of the file (get-content) and 
#   provide the data to the ExcludeFolderEntryIds parameter.
#
#   The default path is UpdatePublicFolderPermission.[yyyyMMdd_HHmm].log where the portion in square brackets
#   gets replaced with the current date and time at the moment of execution of the script.
#
#.EXAMPLE
#    .\Update-PublicFolderPermissions.ps1 -IncludeFolders "\MyFolder" -AccessRights "Owner" -Users "John", "Administrator" -Recurse -Confirm:$false
#
#	This command replaces the current client permissions for users "John" and "Administrator" on the "\MyFolder"
#   Public Folder and all its children. The users will be granted "Owner" access rights. These actions will be 
#	performed without requesting confirmation to the user.
#
#.EXAMPLE
#    $foldersProcessed = get-content .\UpdatePublicFolderPermission.20141031_1820.log
#    .\Update-PublicFolderPermissions.ps1 -IncludeFolders "\MyFolder" -AccessRights "Owner" -Users "John", "Administrator" -Recurse -ExcludeFolderEntryIds $foldersProcessed -Confirm:$false
#
#	These commands replace the current client permissions for users "John" and "Administrator" on the "\MyFolder"
#   Public Folder and all its children but skips those folders that were completd in the execution of Oct 30th 2014 at 6:20 pm.
#   The users will be granted "Owner" access rights. These actions will be performed without requesting confirmation to the user.
#############################################################################################################

param (
    [Parameter(Mandatory=$True)]
    [string[]]$IncludeFolders,
    [Parameter(Mandatory=$True)]
    [string[]]$Users,
    [Parameter(Mandatory=$True)]
    [string[]]$AccessRights,
    [switch]$Recurse,
    [string[]]$ExcludeFolderEntryIds = @(),
    [switch]$SkipCurrentAccessCheck,
    [string]$ProgressLogFile = ".\UpdatePublicFolderPermission.$((Get-Date).ToString('yyyyMMdd_HHmm')).log",
    [switch]$confirm,
    [switch]$whatIf
)

#############################################################################################################
#   Returns the list of public folders to process ignoring duplicates and folders in the exclude list
#############################################################################################################
function FindFoldersToUpdate([string[]]$includeFolders, [bool]$recurseOnFolders, [string[]]$excludeFolderEntryIds)
{
    $folderToSkip = new-object 'System.Collections.Generic.HashSet[string]' -ArgumentList @(,$excludeFolderEntryIds)
    $currentIncludeFolder=0;
    foreach($includeFolder in $includeFolders)
    {
        $progress = 100 * $currentIncludeFolder / $includeFolders.Count;
        Write-Progress -Activity "Retrieving folders to update" -Status $includeFolder -PercentComplete $progress

        $foldersFound = Get-PublicFolder -Recurse:$recurseOnFolders $includeFolder

        if ($foldersFound -eq $null)
        {
            continue;
        }

        foreach($foundFolder in $foldersFound)
        {
            if ($foundFolder -eq $null)
            {
                continue;
            }

            if ($folderToSkip -notContains $foundFolder.EntryId)
            {
                #Return found folder
                $foundFolder;
            }

            $folderToSkip.Add($foundFolder.EntryId) > $null;
        }

        $currentIncludeFolder++;
    }
}

#############################################################################################################
#   Returns the Identity of the users that need processing.
#############################################################################################################
function GetUserIdentities([string[]]$Users)
{
    $userIdentities = new-object 'System.Collections.Generic.HashSet[object]' 
    $currentUserNumber=0;
    foreach($user in $Users)
    {
        $progress = 100 * $currentUserNumber / $Users.Count;
        Write-Progress -Activity "Retrieving users" -Status $user -PercentComplete $progress
        $id = (Get-Recipient $user).Identity

        if ($id -ne $null)
        {
            $userIdentities.Add($id) > $null
        }

        $currentUserNumber++;
    }

    $userIdentities
}

#############################################################################################################
#   Returns whether all the elements of a collection are present in a reference collection.
#############################################################################################################
function CollectionContains($referenceCollection, $otherCollection)
{
    foreach($item in $otherCollection)
    {
        if ($referenceCollection -notcontains $item)
        {
            return $false
        }
    }

    return $true
}

#############################################################################################################
#   Verifies whether there is a mismatch between the desired and found permissions.
#############################################################################################################
function IsUpdateRequired ($currentAccessRights, $desiredAccessRights)
{
    $allDesiredPermissionsWhereFound = CollectionContains $currentAccessRights $desiredAccessRights
    $allFoundPermissionsAreDesired = CollectionContains $desiredAccessRights $currentAccessRights

    return -not ($allDesiredPermissionsWhereFound -and $allFoundPermissionsAreDesired)
}

#############################################################################################################
#   Gets the list of users whose access right to a folder don't match the desired ones.
#############################################################################################################
function GetUsersToUpdate($currentFolder, [Array]$usersToUpdate, [string[]]$accessRights)
{
    Write-Progress -Id 1 -Activity "Querying current permissions" -Status "Processing";

    $existingPermissions = [Array](Get-PublicFolderClientPermission $currentFolder.Identity);
    $existingPermissionsPerUser = @{}

    $permissionCount = 0;
    foreach($permission in $existingPermissions)
    {
        $progress = 100 * $permissionCount / $existingPermissions.Count;
        Write-Progress -Id 1 -Activity "Processing current permissions" -PercentComplete $progress -Status "Processing";

        $adIdentity = $permission.User.ADRecipient.Identity;

        if ($adIdentity -ne $null)
        {
            $existingPermissionsPerUser[$adIdentity] = $permission;
        }
    }

    $permissionCount = 0;
    foreach($user in $usersToUpdate)
    {
        $progress = 100 * $permissionCount / $usersToUpdate.Count;
        Write-Progress -Id 1 -Activity "Comparing permissions" -PercentComplete $progress -Status "Processing";

        if (-not $existingPermissionsPerUser.ContainsKey($user))
        {
            $user;
        }
        else
        {
            if (IsUpdateRequired $existingPermissionsPerUser[$user].AccessRights $AccessRights)
            {
                $user;
            }
        }

        $permissionCount++;
    }
}

#############################################################################################################
#   Script logic.
#############################################################################################################

$foldersToUpdate=[Array](FindFoldersToUpdate $IncludeFolders $Recurse $ExcludeFolderEntryIds);
$usersToUpdate=[Array](GetUserIdentities $Users)

$foldersProcessed=0;
foreach($currentFolder in $foldersToUpdate)
{
    $percentFoldersProcessed = 100 * $foldersProcessed/($foldersToUpdate.Count);
    Write-Progress -Activity "Processing folders" -Status $currentFolder.Identity -PercentComplete $percentFoldersProcessed

    $usersToUpdateForFolder = @()
    if (-not $SkipCurrentAccessCheck)
    {
        $usersToUpdateForFolder =  [Array](GetUsersToUpdate $currentFolder $usersToUpdate $AccessRights)
    }
    else
    {
        $usersToUpdateForFolder = $usersToUpdate;
    }

    $folderOperationFailed=$false;
    $usersProcessed=0;

    if (($usersToUpdateForFolder -eq $null) -or ($usersToUpdateForFolder.Count -eq 0))
    {
        Write-Warning "Couldn't find any changes to perform for folder $($currentFolder.Identity)"
        continue;
    }

    foreach($user in $usersToUpdateForFolder)
    {
        $percentUsersProcessed = 100 * $usersProcessed/($usersToUpdateForFolder.Count)

        Write-Progress -Id 1 -Activity "Processing User" -Status $user -CurrentOperation "Removing exisitng permission" -PercentComplete $percentUsersProcessed
        Remove-PublicFolderClientPermission -User $user $currentFolder.Identity -ErrorAction SilentlyContinue -Confirm:$confirm -WhatIf:$whatIf

        Write-Progress -Id 1 -Activity "Processing User" -Status $user -CurrentOperation "Adding permission" -PercentComplete $percentUsersProcessed

        try
        {
            Add-PublicFolderClientPermission -AccessRights $accessRights -User $user $currentFolder.Identity -ErrorAction Stop -Confirm:$confirm -WhatIf:$whatIf
        }
        catch
        {
            Write-Error $_
            $folderOperationFailed=$true;
        }

        $usersProcessed++;
    }

    if (-not $folderOperationFailed)
    {
        Add-Content $ProgressLogFile "$($currentFolder.EntryId)`n" -Confirm:$confirm -WhatIf:$whatIf
    }

    $foldersProcessed++;
}


# SIG # Begin signature block
# MIIdxAYJKoZIhvcNAQcCoIIdtTCCHbECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUq74XSfqJ3hL3JRUKjy+fin2Y
# s2+gghhkMIIEwzCCA6ugAwIBAgITMwAAAJmqxYGfjKJ9igAAAAAAmTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI4
# WhcNMTcwNjMwMTkyMTI4WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# Ojk4RkQtQzYxRS1FNjQxMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAipCth86FRu1y
# rtsPu2NLSV7nv6A/oVAcvGrG7VQwRn+wGlXrBf4nyiybHGn9MxuB9u4EMvH8s75d
# kt73WT7lGIT1yCIh9VC9ds1iWfmxHZtYutUOM92+a22ukQW00T8U2yowZ6Gav4Q7
# +9M1UrPniZXDwM3Wqm0wkklmwfgEEm+yyCbMkNRFSCG9PIzZqm6CuBvdji9nMvfu
# TlqxaWbaFgVRaglhz+/eLJT1e45AsGni9XkjKL6VJrabxRAYzEMw4qSWshoHsEh2
# PD1iuKjLvYspWv4EBCQPPIOpGYOxpMWRq0t/gqC+oJnXgHw6D5fZ2Ccqmu4/u3cN
# /aAt+9uw4wIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFHbWEvi6BVbwsceywvljICto
# twQRMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBABbNYMMt3JjfMAntjQhHrOz4aUk970f/hJw1jLfYspFpq+Gk
# W3jMkUu3Gev/PjRlr/rDseFIMXEq2tEf/yp72el6cglFB1/cbfDcdimLQD6WPZQy
# AfrpEccCLaouf7mz9DGQ0b9C+ha93XZonTwPqWmp5dc+YiTpeAKc1vao0+ru/fuZ
# ROex8Zd99r6eoZx0tUIxaA5sTWMW6Y+05vZN3Ok8/+hwqMlwgNR/NnVAOg2isk9w
# ox9S1oyY9aRza1jI46fbmC88z944ECfLr9gja3UKRMkB3P246ltsiH1fz0kFAq/l
# 2eurmfoEnhg8n3OHY5a/Zzo0+W9s1ylfUecoZ4UwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBMowggTGAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB3jAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUkvJPfjobLXnJTV4KFSsbfQBtjY0wfgYKKwYB
# BAGCNwIBDDFwMG6gRoBEAFUAcABkAGEAdABlAC0AUAB1AGIAbABpAGMARgBvAGwA
# ZABlAHIAUABlAHIAbQBpAHMAcwBpAG8AbgBzAC4AcABzADGhJIAiaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQBa9EA+
# eliZ/d6sr3+F0Vuhj+FH95jRZ7zdT5HZFLtx44axKiT7yXFgYG95sFKoRXvQ/Xnw
# GBFoJH9PLhtlKRt11Ilgm2X9EsnzdB2B47hPkjApypG3MsrVxRYFB/wayqj9NLdF
# W5tzRJTgc6nrICFDFP1p1tNMnKumwm5cNZTFpflvaJ/dUEimWs3qyXMwbSG3Zjye
# ndGEeho+/WW2prxNVtgTNbCFsdBHdEM1Cw6cWUjEkmuJCqQOQFzgtVvV1OV8+9VL
# mBMpm1xFzesZps9ww1qzhr8yQ6BGnRfoDDiG7Wf21tjtvC1S/fzPJn5qMVBcTBrg
# IDbj1it2rCUdJFBPoYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAJmqxYGfjKJ9igAAAAAAmTAJBgUr
# DgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMTYwOTAzMTg0NDUyWjAjBgkqhkiG9w0BCQQxFgQUlYID+tUk3qckVxH5vKrY
# qob6gh8wDQYJKoZIhvcNAQEFBQAEggEAKfna/6BlxVjkGafyFCxEES2PjvXM2Lvf
# PH3S4UaH9fb/30aIOX4VUNsMEWSCUmra2Lizxo98xXCdfzf+E4r+mwaE6D4WQ7ai
# WVQjjv+Uljfr20EraUJkJDWVt76ARk21hpgHurCuD/whGBFn9+ov4tNvwFaXjEm1
# WbGnUvq0bW+PSvhpjLzYyJVgwwEY2/1Jg/afW3Zlwpuctp7Ekv0D+wWBxRI0CxWv
# Bpu+V0cLWUns5iosVwVwu6FLz9NwN2AibfFV0RHCr2e4qDW3NP6bEwIFZRt4yiRB
# k4nKnXYs7qYde8SdIinJLCWDp2IT0Co71W68c0pMsm3NtTzdzcVGvQ==
# SIG # End signature block
