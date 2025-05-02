# .SYNOPSIS
# Merge-PublicFolderMailbox.ps1
#    Merges the contents of the given public folder mailbox with the target public folder mailbox
#
# .DESCRIPTION
#
# Copyright (c) 2011 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

Param(
    # Source Mailbox
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Please specify the source public folder mailbox to merge from")]
    [ValidateNotNull()]
    [string] $SourcePublicFolderMailbox,
    
    # Target Mailbox
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Please specify the target public folder mailbox where the merge contents need to go")]
    [ValidateNotNull()]
    [string] $TargetPublicFolderMailbox,
    
    # Name of the organization
    [Parameter(Mandatory=$false)]
    [ValidateNotNull()]
    [string] $OrganizationName,
    
    # Minimum percentage to Split
    [Parameter(Mandatory=$false)]
    [ValidateNotNull()]
    [ValidateRange(1,100)]
    [int] $MinimumPercentageToSplit,
    
    [Parameter(Mandatory=$false)]
    [switch] $Whatif
    )

################ START OF DEFAULTS ################
$script:minPercentToSplit = 60;
$script:targetPercentToFillFromSplitPoint = 60;
$script:FoldersToMerge = @();
$script:sourceMailboxInformation  = $null;
$script:targetMailboxInformation  = $null;
$script:IPM_SUBTREE = "IPM_SUBTREE";
################ END OF DEFAULTS ################

############## Function to get information about quota ##########
function GetMailboxQuota()
{
    param ($mailboxType)
    
    if ($mailboxType -match "Source")
    {
        $useDatabaseQuotaDefault = $script:sourceMailboxInformation.UseDatabaseQuotaDefaults;
        $databaseName = $script:sourceMailboxInformation.Database.ToString();
        $prohibitSendReceiveQuota = $script:sourceMailboxInformation.ProhibitSendReceiveQuota.Value;
    }
    else
    {
        $useDatabaseQuotaDefault = $script:targetMailboxInformation.UseDatabaseQuotaDefaults;
        $databaseName = $script:targetMailboxInformation.Database.ToString();
        $prohibitSendReceiveQuota = $script:targetMailboxInformation.ProhibitSendReceiveQuota.Value;
    }
     
    if ($useDatabaseQuotaDefault -eq $true)
    {
        if ($databaseName -ne $null)
        {
            $prohibitSendReceiveQuota = (Get-MailboxDatabase -Identity:$databaseName).ProhibitSendReceiveQuota.Value; 
        }
        else
        {
            return $null;
        }
    }
   
    return $prohibitSendReceiveQuota;    
}
############## End Function #####################################

############## Function to ensure if mailbox can be split #######
function EnsureMergeMailboxPossible()
{
    # Retrieving statistics information on the source mailbox
    if ($OrganizationName -ne "")
    {
        $sourceMailboxOccupied = (Get-MailboxStatistics -Identity:($OrganizationName + "\" + $SourcePublicFolderMailbox)).TotalItemSize.Value.ToBytes();
    }
    else
    {
        $sourceMailboxOccupied = (Get-MailboxStatistics -Identity:$SourcePublicFolderMailbox).TotalItemSize.Value.ToBytes();
    }

    if ($sourceMailboxOccupied -eq $null)
    {
        exit;
    }
    
    # Determining if the minimum percentage is passed as an argument
    if ($MinimumPercentageToSplit -ne "")
    {
        $script:minPercentToSplit = [int]$MinimumPercentageToSplit;
    }
    
    # Determining ProhibitSendReceiveQuota value
    $targetMailboxTotalSize = GetMailboxQuota("Target");
    
    if ($targetMailboxTotalSize -eq $null)
    {
        return $true;
    }
    
    # Retrieving statistics information on the target mailbox
    if ($OrganizationName -ne "")
    {
        $targetMailboxOccupied = (Get-MailboxStatistics -Identity:($OrganizationName + "\" + $TargetPublicFolderMailbox)).TotalItemSize.Value.ToBytes();
    }
    else
    {
        $targetMailboxOccupied = (Get-MailboxStatistics -Identity:$TargetPublicFolderMailbox).TotalItemSize.Value.ToBytes();
    }

    if ($targetMailboxOccupied -eq $null)
    {
        exit;
    }
    
    $minTargetMailboxSizeToSplit = [Math]::Round(($script:minPercentToSplit / 100 * $targetMailboxTotalSize.ToBytes()), 0);
    $maxTargetMailboxSizeAfterSplit = [Math]::Round($script:targetPercentToFillFromSplitPoint / 100 * $minTargetMailboxSizeToSplit, 0);
    
    # Check to see if the given target mailbox is the right candidate
    if (($sourceMailboxOccupied + $targetMailboxOccupied) -le $maxTargetMailboxSizeAfterSplit)
    {
        return $true;
    }
    else
    {
        Write-Host;
        Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.ImpossibleToMove -f $TargetPublicFolderMailbox);
        return $false;
    }
}
############## End Function #####################################

#load hashtable of localized string
Import-LocalizedData -BindingVariable PublicFolderManagement_LocalizedStrings -FileName PublicFolderStorageManagement.strings.psd1

if ($SourcePublicFolderMailbox -eq $TargetPublicFolderMailbox)
{
    Write-Host "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.SameMailbox;
    exit; 
}

# Retrieve the information about the given source public folder mailbox
if ($OrganizationName -ne "")
{
    $script:sourceMailboxInformation = Get-Mailbox -PublicFolder -Identity:$SourcePublicFolderMailbox -Organization:$OrganizationName;
    $script:targetMailboxInformation = Get-Mailbox -PublicFolder -Identity:$TargetPublicFolderMailbox -Organization:$OrganizationName;
}
else
{
    $script:sourceMailboxInformation = Get-Mailbox -PublicFolder -Identity:$SourcePublicFolderMailbox;
    $script:targetMailboxInformation = Get-Mailbox -PublicFolder -Identity:$TargetPublicFolderMailbox;
}

if ($script:sourceMailboxInformation -eq $null -or $script:targetMailboxInformation -eq $null)
{
    exit;
}

Write-Host;
Write-Host "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.FeasibilityToMerge;
$retVal = EnsureMergeMailboxPossible;
if ($retVal -ne $true)
{
    Write-Host;
    Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.ImpossibleToMerge -f $SourcePublicFolderMailbox);
    exit;
}

Write-Host;
Write-Host "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.RetrieveFoldersFromSourceMailbox;

if ($OrganizationName -ne "")
{ 
    $script:publicFoldersToProcess = Get-PublicFolder -Organization:$OrganizationName -ResidentFolders -Mailbox $script:sourceMailboxInformation.ExchangeGuid.ToString() -Recurse -WarningAction SilentlyContinue | `
                                     Where-Object -FilterScript { $_.Name -ne $script:IPM_SUBTREE };
}
else
{
    $script:publicFoldersToProcess = Get-PublicFolder -ResidentFolders -Mailbox $script:sourceMailboxInformation.ExchangeGuid.ToString() -Recurse -WarningAction SilentlyContinue | `
                                     Where-Object -FilterScript { $_.Name -ne $script:IPM_SUBTREE };
}

# Check if there is atleast one folder belonging to the given source public folder mailbox
if ($script:publicFoldersToProcess.Count -lt 1)
{
    Write-Host;
    Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.FolderUnavailableToMerge -f $SourcePublicFolderMailbox);
    exit;
}

# Populating the folders that will be part of the merge request
$publicFoldersToProcessCount = $script:publicFoldersToProcess.Count;
for ($index = 0; $index -lt $publicFoldersToProcessCount; $index++)
{
    $script:FoldersToMerge += $script:publicFoldersToProcess[$index].Identity;
}

Write-Host;
Write-Host "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.MoveFolders;
$script:FoldersToMerge | Format-Wide -Property MapiFolderPath -Column 1;

if ($Whatif)
{
    exit;
}

# Checking if there are any pending request in this organization
if ($OrganizationName -ne "")
{
    $anyPendingRequest = Get-PublicFolderMoveRequest -Organization:$OrganizationName;
}
else
{
    $anyPendingRequest = Get-PublicFolderMoveRequest;
}

if ($anyPendingRequest -ne $null)
{
    Write-Host;
    Write-Host "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.RemoveExistingRequest;
    exit;
}

# Initiating the request
Write-Host;
Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.IssueMergeRequest -f $SourcePublicFolderMailbox);
if ($OrganizationName -ne "")
{
    $request = New-PublicFolderMoveRequest -Folders:$script:FoldersToMerge -TargetMailbox:$TargetPublicFolderMailbox -Organization:$OrganizationName;
}
else
{
    $request = New-PublicFolderMoveRequest -Folders:$script:FoldersToMerge -TargetMailbox:$TargetPublicFolderMailbox;
}

if ($request -ne $null)
{
   Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.RequestName -f $($request));
   Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.SourceMailbox -f $($request.SourceMailbox));
   Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.TargetMailbox -f $($request.TargetMailbox));
   Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderManagement_LocalizedStrings.RequestStatus -f $($request.Status));
   Write-Host;
   Write-Host "[$($(Get-Date).ToString())]" $PublicFolderManagement_LocalizedStrings.JobStatus;
}
# SIG # Begin signature block
# MIIdugYJKoZIhvcNAQcCoIIdqzCCHacCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUHLY2U5ztU6xQpb6wRfxcxOO6
# MxKgghhkMIIEwzCCA6ugAwIBAgITMwAAAJ1CaO4xHNdWvQAAAAAAnTANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBMAwggS8AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB1DAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQU3KyXgw+fVhoV6xqCMnf/QQPfAUwwdAYKKwYB
# BAGCNwIBDDFmMGSgPIA6AE0AZQByAGcAZQAtAFAAdQBiAGwAaQBjAEYAbwBsAGQA
# ZQByAE0AYQBpAGwAYgBvAHgALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBACRJFVOQfqNBlgztdeJ/
# z+cKdD3ATQLDRXyD76vedOknbf39m152jwbTSFExRDAtWjKq/YnlImGkKPMcJAtN
# fel8Pkdf/mc8yyQP7VdOdHkVll67lFx+rCwQTbCmS8XxWD6MlKrdAyVjAekqivVZ
# CQ+XIUfeMbBUIew4l1KA5D6eQ10Bn9P4P1f4s5anw6SRWNDpc9z2Y0k7FOsHQqum
# VxtlQI7SVS6HlXrpPBgVw8E4uIqi5MEE8ZGtM6mffZbSrOcNNL5ovCj7VrXpJfCZ
# IsWKdLNrdoRYdbe/zfJ5j6GaYO0QQhPk8Oa5kDCc8cwDb8KIgKpu/cRpV4dw60qd
# YfShggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBAhMzAAAAnUJo7jEc11a9AAAAAACdMAkGBSsOAwIaBQCgXTAY
# BgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMx
# ODQ1MDZaMCMGCSqGSIb3DQEJBDEWBBTmRbQDlsVJFGK5zVusvFu7OVeBfzANBgkq
# hkiG9w0BAQUFAASCAQCSzLyn37OAhvdF3aCDXLppHEDNut3k8177hokF97Ett80q
# DSqMvnS3SzZ6l44w9pRzW0XqQBNoQnwq10wYBelKqL7GJJUZ9fpHrjFj49PhUlkU
# B7sLPPs15OAyo1H07i+uPnVSV/5Zs77TW//S4EokfYWiDgduXkP5zSI93j1WaMco
# EuvE9m8Qbbh9iM6wuD3NXPfWjVs6PVFckrIdcTp7sN5G0HcPOcQn9BTgqv4pe2bx
# tjEs9Y3Dssf+yBKwu8FOjhInFnEh/Z11+7uAakFcJYGMTvShWjzRrryCgM51zYlI
# Eze0GHnd22IE6sHoTvjv5h7DHwx9DUjYKDc1fjjH
# SIG # End signature block
