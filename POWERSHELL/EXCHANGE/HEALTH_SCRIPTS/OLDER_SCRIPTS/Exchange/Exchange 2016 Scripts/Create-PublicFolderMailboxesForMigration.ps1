﻿# .SYNOPSIS
#    Creates public folder mailboxes, based on the output of PublicFolderToMailboxMapGenerator.ps1 script, in preparation for Public Folder migration.
#
# .DESCRIPTION
#    The script must be executed on target side of migration. For migration from Exchange 2007 or 2010 to Exchange Online, the script must be executed on a
#    remote PowerShell session created against the Exchange Online organization. For migration from Exchange 2007 or 2010 to Exchange 2013 or 2016 on the
#    same forest, the script must executed on a PowerShell session opened against Exchange 2013 or 2016 servers.
#
#    Copyright (c) 2015 Microsoft Corporation. All rights reserved.
#
#    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
#    OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
# .PARAMETER FolderMappingCsv
#    Mapping file generated by the PublicFoldertoMailboxMapGenerator.ps1 script.
#
# .PARAMETER EstimatedNumberOfConcurrentUsers
#    Estimated number of simultaneous user connections browsing the Public Folder Hierarchy. This is usually less than the total users in the organization.
#
# .PARAMETER Confirm
#    The Confirm switch causes the script to pause processing and requires you to acknowledge what the script will do before processing continues. You don't have to specify
#    a value with the Confirm switch.
#
# .PARAMETER WhatIf
#    The WhatIf switch instructs the script to simulate the actions that it would take on the object. By using the WhatIf switch, you can view what changes would occur
#    without having to apply any of those changes. You don't have to specify a value with the WhatIf switch.
#
# .EXAMPLE
#    .\Create-PublicFolderMailboxesForMigration.ps1 -FolderMappingCsv mapping.csv -EstimatedNumberOfConcurrentUsers:2000
#
#    This example shows how to invoke the script to create the target public folder mailboxes for migration and also to support 2,000 users browsing the hierarchy at the same time.
#
param(
    [Parameter(Mandatory=$true, HelpMessage = "Mapping file generated by the PublicFoldertoMailboxMapGenerator.ps1 script")]
    [ValidateNotNull()]
    [string] $FolderMappingCsv,

    [Parameter(Mandatory=$true, HelpMessage = "Estimated number of simultaneous user connections browsing the Public Folder Hierarchy")]
    [int] $EstimatedNumberOfConcurrentUsers,

    [Parameter(Mandatory=$false)]
    [bool] $Confirm = $true,

    [Parameter(Mandatory=$false)]
    [switch] $WhatIf = $false
)

function PromptForConfirmation()
{
    param ($title, $message)

    $yes = New-Object System.Management.Automation.Host.ChoiceDescription $LocalizedStrings.ConfirmationYesOption;
    $no = New-Object System.Management.Automation.Host.ChoiceDescription $LocalizedStrings.ConfirmationNoOption;
    [System.Management.Automation.Host.ChoiceDescription[]]$options = $yes, $no;

    $confirmation = $host.ui.PromptForChoice($title, $message, $options, 0);
    if ($confirmation -ne 0)
    {
        Exit;
    }
}

function CreateMailbox()
{
    param($name, $isExcludedFromServingHierarchy, $isMigrationTarget, $isPrimary)

    $mbx = New-Mailbox -PublicFolder $name -IsExcludedFromServingHierarchy:$isExcludedFromServingHierarchy -HoldForMigration:$isPrimary -WhatIf:$WhatIf;
    $outputObj = New-Object psobject;
    $outputObj | add-member noteproperty "Name" ($name);
    $outputObj | add-member noteproperty "IsServingHierarchy" (-not $isExcludedFromServingHierarchy);
    $outputObj | add-member noteproperty "IsMigrationTarget" ($isMigrationTarget);
    $script:mailboxesCreated += $outputObj;

    if (-not $isExcludedFromServingHierarchy)
    {
        $script:totalMailboxesServingHierarchy++;
    }
}

function CreatePrimaryMailbox()
{
    param($name, $isExcludedFromServingHierarchy, $isMigrationTarget)
    CreateMailbox $name $isExcludedFromServingHierarchy $isMigrationTarget $true;
}

function CreateSecondaryMailbox()
{
    param($name, $isExcludedFromServingHierarchy, $isMigrationTarget)
    CreateMailbox $name $isExcludedFromServingHierarchy $isMigrationTarget $false;
}

$MaxConnectionsPerMailbox = 2000;
$MaxMailboxesServingHierarchy = 100;
$script:totalMailboxesServingHierarchy = 0;
$script:mailboxesCreated = @();
$script:mailboxUpdates = @{};

Import-LocalizedData -BindingVariable LocalizedStrings -FileName CreatePublicFolderMailboxesForMigration.strings.psd1

if ($EstimatedNumberOfConcurrentUsers -lt 1)
{
    Write-Error ($LocalizedStrings.InvalidNumberOfConcurrentUsers -f $EstimatedNumberOfConcurrentUsers);
    return;
}

$FolderMappings = Import-Csv $FolderMappingCsv -ErrorAction Stop;
if ($FolderMappings.Count -eq 0)
{
    Write-Error $LocalizedStrings.InvalidCsvEmptyMapping;
    return;
}

# Validate the input CSV and collect the unique mailbox names (multiple folders can map to the same mailbox).
$primaryMailboxName = $null;
$secondaryMailboxesToCreateFromCsv = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::CurrentCultureIgnoreCase);
$folderPaths = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::CurrentCultureIgnoreCase);
foreach ($folderMapping in $FolderMappings)
{
    if (-not $folderPaths.Add($folderMapping.FolderPath))
    {
        Write-Error ($LocalizedStrings.InvalidCsvDuplicateMapping -f $folderMapping.FolderPath);
        Exit;
    }

    if ($folderMapping.FolderPath -eq "\")
    {
        $primaryMailboxName = $folderMapping.TargetMailbox;
    }
    else
    {
        [void]$secondaryMailboxesToCreateFromCsv.Add($folderMapping.TargetMailbox);
    }
}

if ($primaryMailboxName -eq $null)
{
    Write-Error $LocalizedStrings.InvalidCsvMissingRootFolder;
    Exit;
}

# The order in which primary mailbox appears on the CSV is unpredictable. It may also appear several times mapped
# to different folders. Because of that, it may get added to the secondaryMailboxesToCreateFromCsv collection.
[void]$secondaryMailboxesToCreateFromCsv.Remove($primaryMailboxName);

# Determine the total count of mailboxes to create serving hierarchy
# If primary is the single target for migration, it must serve the hierarchy.
$excludePrimaryFromServingHierarchy = ($secondaryMailboxesToCreateFromCsv.Count -gt 0);
$totalMailboxesServingHierarchyQuota = [Math]::Ceiling($EstimatedNumberOfConcurrentUsers / $MaxConnectionsPerMailbox);
if ($totalMailboxesServingHierarchyQuota -le 1)
{
    # There should be at least one mailbox serving hierarchy.
    $totalMailboxesServingHierarchyQuota = 1;
}
elseif ($totalMailboxesServingHierarchyQuota -gt $MaxMailboxesServingHierarchy)
{
    if ($Confirm)
    {
        PromptForConfirmation $LocalizedStrings.ConfirmTooManyUsersTitle ($LocalizedStrings.ConfirmTooManyUsersMessage -f $MaxMailboxesServingHierarchy);
    }

    $totalMailboxesServingHierarchyQuota = $MaxMailboxesServingHierarchy;
}

$hierarchyMailboxInfo = $(Get-OrganizationConfig).RootPublicFolderMailbox;
if ($hierarchyMailboxInfo.HierarchyMailboxGuid -ne [Guid]::Empty)
{
    if (-not $hierarchyMailboxInfo.LockedForMigration)
    {
        Write-Error $LocalizedStrings.DeploymentNotLockedForMigration;
        Exit;
    }

    $primaryMailbox = Get-Mailbox -PublicFolder $hierarchyMailboxInfo.HierarchyMailboxGuid.ToString() -ErrorAction:Stop;
    if ($primaryMailboxName -ne $primaryMailbox.Name)
    {
        Write-Error ($LocalizedStrings.PrimaryMailboxNameNotMatching -f $primaryMailboxName, $primaryMailbox.Name);
        Exit;
    }

    if ($Confirm)
    {
        PromptForConfirmation $LocalizedStrings.PublicFolderMailboxesAlreadyExistTitle $LocalizedStrings.PublicFolderMailboxesAlreadyExistMessage;
    }

    $secondaryMailboxesServingHierarchyQuota = $totalMailboxesServingHierarchyQuota;
    if ($excludePrimaryFromServingHierarchy)
    {
        # Exclude primary from serving hierarchy when there is at least one other secondary mailbox in the migration plan
        if (-not $primaryMailbox.IsExcludedFromServingHierarchy)
        {
            $script:mailboxUpdates.Add($primaryMailbox.Identity, $true);
        }
    }
    else
    {
        # If primary is the only mailbox in the CSV, it must serve the hierarchy.
        if ($primaryMailbox.IsExcludedFromServingHierarchy)
        {
            $script:mailboxUpdates.Add($primaryMailbox.Identity, $false);
        }

        # Since primary will be serving, reduce the quota for secondary by 1
        $secondaryMailboxesServingHierarchyQuota--;
        $script:totalMailboxesServingHierarchy++;
    }

    $existingMailboxes = @(Get-Mailbox -PublicFolder -ResultSize:Unlimited -ErrorAction:Stop);
    $existingSecondaryMailboxesInfo = @();
    $existingSecondaryInCsvServingHierarchyCount = 0;
    foreach ($existingMailbox in $existingMailboxes)
    {
        if ($primaryMailbox.ExchangeGuid -eq $existingMailbox.ExchangeGuid)
        {
            # Skip primary
            continue;
        }

        $migrationTarget = $secondaryMailboxesToCreateFromCsv.Remove($existingMailbox.Name);
        $mailboxInfo = New-Object psobject -Property @{
            Name = $existingMailbox.Name;
            ServingHierarchy = (-not $existingMailbox.IsExcludedFromServingHierarchy);
            MigrationTarget = $migrationTarget;
        };

        if ($migrationTarget -and `
            $mailboxInfo.ServingHierarchy -and `
            $secondaryMailboxesServingHierarchyQuota -gt 0)
        {
            $secondaryMailboxesServingHierarchyQuota--;
            $script:totalMailboxesServingHierarchy++;
        }
        else
        {
            $existingSecondaryMailboxesInfo += $mailboxInfo;
        }
    }

    if ($secondaryMailboxesServingHierarchyQuota -le $secondaryMailboxesToCreateFromCsv.Count)
    {
        $secondaryMailboxesServingHierarchyQuota = 0;
    }
    else
    {
        $secondaryMailboxesServingHierarchyQuota -= $secondaryMailboxesToCreateFromCsv.Count;
    }

    # Sort the mailboxes in the priority order to serve hierarchy. We should favor mailboxes that are included on the input CSV
    # (i.e. mailboxes that are part of the migration plan). Secondly, we should minimize the number of updates needed to include
    # or exclude existing mailboxes in the hierarchy-serving collection.
    $existingSecondaryMailboxesInfo = @($existingSecondaryMailboxesInfo | sort MigrationTarget,ServingHierarchy -Descending);
    for ($i = 0; $i -lt $existingSecondaryMailboxesInfo.Length; $i++)
    {
        $mailboxInfo = $existingSecondaryMailboxesInfo[$i];
        if ($i -lt $secondaryMailboxesServingHierarchyQuota -and (-not $mailboxInfo.ServingHierarchy))
        {
            # Update to serve hierarchy
            $script:mailboxUpdates.Add($mailboxInfo.Name, $false);
            $script:totalMailboxesServingHierarchy++;
        }
        elseif ($mailboxInfo.ServingHierarchy)
        {
            if ($i -ge $secondaryMailboxesServingHierarchyQuota)
            {
                # Exclude from serving hierarchy
                $script:mailboxUpdates.Add($mailboxInfo.Name, $true);
            }
            else
            {
                $script:totalMailboxesServingHierarchy++;
            }
        }
    }
}

$totalMailboxesToCreate = $secondaryMailboxesToCreateFromCsv.Count;
if ($hierarchyMailboxInfo.HierarchyMailboxGuid -eq [Guid]::Empty)
{
    $totalMailboxesToCreate++;
}

if ($Confirm)
{
    $confirmationMessage = ($LocalizedStrings.ConfirmMailboxOperationsMessage -f ($totalMailboxesToCreate + $additionalCount), $script:mailboxUpdates.Count, $totalMailboxesServingHierarchyQuota);
    PromptForConfirmation $LocalizedStrings.ConfirmMailboxOperationsTitle $confirmationMessage;
}

# Execute operations in the following order:
# 1. Exclude existing mailboxes from serving hierarchy
# 2. Mark existing mailboxes to serve hierarchy (1 and 2 are the reason why we sort the updates by "Value")
# 3. Create the mailboxes that are part of the migration plan
# 4. Create additional mailboxes to serve hierarchy
$updatesExecuted = 0;
foreach ($update in ($script:mailboxUpdates.GetEnumerator() | sort Value -Descending))
{
    Write-Progress -Activity $LocalizedStrings.UpdatingMailboxesActivity `
            -Status ($LocalizedStrings.UpdatingMailboxesProgressStatus -f $updatesExecuted, $script:mailboxUpdates.Count, $update.Key) `
            -PercentComplete (100 * $updatesExecuted / $script:mailboxUpdates.Count);

    Set-Mailbox -PublicFolder -Identity:$update.Key -IsExcludedFromServingHierarchy:$update.Value -ErrorAction:Stop;
    $updatesExecuted++;
}

# Create primary if it doesn't exist yet.
if ($hierarchyMailboxInfo.HierarchyMailboxGuid -eq [Guid]::Empty)
{
    Write-Progress -Activity $LocalizedStrings.CreatingMailboxesForMigrationActivity `
            -Status ($LocalizedStrings.CreatingMailboxesProgressStatus -f 0, $totalMailboxesToCreate, $primaryMailboxName) `
            -PercentComplete 0;

    CreatePrimaryMailbox $primaryMailboxName $excludePrimaryFromServingHierarchy $true;
}
else
{
    Write-Host $LocalizedStrings.SkipPrimaryCreation;
}

$isMigrationTarget = $true;
foreach($secondaryMailboxName in $secondaryMailboxesToCreateFromCsv)
{
    Write-Progress -Activity $LocalizedStrings.CreatingMailboxesForMigrationActivity `
        -Status ($LocalizedStrings.CreatingMailboxesProgressStatus -f $script:mailboxesCreated.Length, $totalMailboxesToCreate, $secondaryMailboxName) `
        -PercentComplete (100 * $script:mailboxesCreated.Length / $totalMailboxesToCreate);

    $excludeFromServingHierarchy = ($script:totalMailboxesServingHierarchy -ge $totalMailboxesServingHierarchyQuota);
    CreateSecondaryMailbox $secondaryMailboxName $excludeFromServingHierarchy $isMigrationTarget;
}

$isMigrationTarget = $false;
$excludeFromServingHierarchy = $false;
$additionalCount = $totalMailboxesServingHierarchyQuota - $script:totalMailboxesServingHierarchy;
for ($i = 0; $i -lt $additionalCount; $i++)
{
    $name = "AutoSplit_" + [Guid]::NewGuid().ToString("n");
    Write-Progress -Activity $LocalizedStrings.CreatingMailboxesToServeHierarchyActivity `
        -Status ($LocalizedStrings.CreatingMailboxesProgressStatus -f $i, $additionalCount, $name) `
        -PercentComplete (100 * $i / $additionalCount);

    CreateSecondaryMailbox $name $excludeFromServingHierarchy $isMigrationTarget;
}

Write-Host ($LocalizedStrings.FinalSummary -f $script:mailboxesCreated.Length, $script:mailboxUpdates.Count, $script:totalMailboxesServingHierarchy);

if ($script:mailboxesCreated.Length -gt 0)
{
    Write-Host $LocalizedStrings.MailboxesCreatedSummary;
    $script:mailboxesCreated | ft -a;
}
# SIG # Begin signature block
# MIId2wYJKoZIhvcNAQcCoIIdzDCCHcgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUbw5cw8VzzwoPeFHhnALpjtBt
# YM+gghhkMIIEwzCCA6ugAwIBAgITMwAAAJmqxYGfjKJ9igAAAAAAmTANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBOEwggTdAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB9TAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUE7xZz8aTRLCn97syAOKlOLAqZXIwgZQGCisG
# AQQBgjcCAQwxgYUwgYKgWoBYAEMAcgBlAGEAdABlAC0AUAB1AGIAbABpAGMARgBv
# AGwAZABlAHIATQBhAGkAbABiAG8AeABlAHMARgBvAHIATQBpAGcAcgBhAHQAaQBv
# AG4ALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2Ug
# MA0GCSqGSIb3DQEBAQUABIIBAB/QG/k4C3RTiMefVmgIjdkxFnCBTdT7gd4rGtfc
# CnkgSj9cSIt3S33nzObFKTZ32RQ1barmgECHUfG9J3odtOhB9O+tlCW7IXsmmKiq
# 7/zT+8njx0LFABHvRABWL/wSNPIA1fDWznO1Y+e4DYxtynkMsM0OGaEuSG43Ef3m
# MyPVMsjWsauqRk/LjbVwyEsXM5h1m4qGRGO2m23egOAdhERyCHwqigIabRuvW2XR
# vF0Ep62BbNqMK1rplk84Sd7lLWSuBdV2A1bkqVEkIoHNTTLSQOttNi8fwuypPsV7
# PKx37FPKOZgos+ahghCWrVl5wU8Lb2VNUYIU/xYjRbrTS9ehggIoMIICJAYJKoZI
# hvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMz
# AAAAmarFgZ+Mon2KAAAAAACZMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJ
# KoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMxODQ0NTRaMCMGCSqGSIb3
# DQEJBDEWBBRyiOQ82jueePZF7bo5oRG9dAb/CDANBgkqhkiG9w0BAQUFAASCAQBp
# O2d0/Zq+YJK4CGh96pqDDVJmHyTYs6ZJG0kMBsBQx4FhD4Gfwom2WkaOK2hdySdZ
# dcIcwwX78f9P6QdxkADaIztdAkewF5UOJcFfpvr2GK59mgqNYa67zmg0P1Eg/OaE
# XtvCPamkK021xk8xKLadKdVoYZzYrz5WlIMJGIB2t72Z3aYeJzH+6tzEx4DCc1C2
# 6sA5R37juKAF04E7erDi9de3uhUPYeGdU7bma3uqcM3HAgJweY8PworFC7HpHDWG
# aPXTSQrlzpyjLD2bdTaYMTbOtqBl4DS2hXg2xNDjh4JMva4k8CRE+9ZiH6Gy5sf6
# 7hUVLh1j16OM4sIxwzCf
# SIG # End signature block