# .SYNOPSIS
# Export-PublicFolderStatistics.ps1
#    Generates a CSV file that contains the list of public folders and their individual sizes
#
# .DESCRIPTION
#
# Copyright (c) 2011 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

Param(
    # File to export to
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Full path of the output file to be generated. If only filename is specified, then the output file will be generated in the current directory.")]
    [ValidateNotNull()]
    [string] $ExportFile,
    
    # Server to connect to for generating statistics
    [Parameter(
        Mandatory=$true,
        HelpMessage = "Public folder server to enumerate the folder hierarchy.")]
    [ValidateNotNull()]
    [string] $PublicFolderServer
    )

#load hashtable of localized string
Import-LocalizedData -BindingVariable PublicFolderStatistics_LocalizedStrings -FileName Export-PublicFolderStatistics.strings.psd1
    
################ START OF DEFAULTS ################

$WarningPreference = 'SilentlyContinue';
$script:Exchange14MajorVersion = 14;
$script:Exchange12MajorVersion = 8;

################ END OF DEFAULTS #################

# Function that determines if to skip the given folder
function IsSkippableFolder()
{
    param($publicFolder);
    
    $publicFolderIdentity = $publicFolder.Identity.ToString();

    for ($index = 0; $index -lt $script:SkippedSubtree.length; $index++)
    {
        if ($publicFolderIdentity.StartsWith($script:SkippedSubtree[$index]))
        {
            return $true;
        }
    }
    
    return $false;
}

# Function that gathers information about different public folders
function GetPublicFolderDatabases()
{
    $script:ServerInfo = Get-ExchangeServer -Identity:$PublicFolderServer;
    $script:PublicFolderDatabasesInOrg = @();
    if ($script:ServerInfo.AdminDisplayVersion.Major -eq $script:Exchange14MajorVersion)
    {
        $script:PublicFolderDatabasesInOrg = @(Get-PublicFolderDatabase -IncludePreExchange2010);
    }
    elseif ($script:ServerInfo.AdminDisplayVersion.Major -eq $script:Exchange12MajorVersion)
    {
        $script:PublicFolderDatabasesInOrg = @(Get-PublicFolderDatabase -IncludePreExchange2007);
    }
    else
    {
        $script:PublicFolderDatabasesInOrg = @(Get-PublicFolderDatabase);
    }
}

# Function that executes statistics cmdlet on different public folder databases
function GatherStatistics()
{   
    # Running Get-PublicFolderStatistics against each server identified via Get-PublicFolderDatabase cmdlet
    $databaseCount = $($script:PublicFolderDatabasesInOrg.Count);
    $index = 0;
    
    if ($script:ServerInfo.AdminDisplayVersion.Major -eq $script:Exchange12MajorVersion)
    {
        $getPublicFolderStatistics = "@(Get-PublicFolderStatistics ";
    }
    else
    {
        $getPublicFolderStatistics = "@(Get-PublicFolderStatistics -ResultSize:Unlimited ";
    }

    While ($index -lt $databaseCount)
    {
        $serverName = $($script:PublicFolderDatabasesInOrg[$index]).Server.Name;
        $getPublicFolderStatisticsCommand = $getPublicFolderStatistics + "-Server $serverName)";
        Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.RetrievingStatistics -f $serverName);
        $publicFolderStatistics = Invoke-Expression $getPublicFolderStatisticsCommand;
        Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.RetrievingStatisticsComplete -f $serverName,$($publicFolderStatistics.Count));
        RemoveDuplicatesFromFolderStatistics $publicFolderStatistics;
        Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.UniqueFoldersFound -f $($script:FolderStatistics.Count));
        $index++;
    }
}

# Function that removed redundant entries from output of Get-PublicFolderStatistics
function RemoveDuplicatesFromFolderStatistics()
{
    param($publicFolders);
    
    $index = 0;
    While ($index -lt $publicFolders.Count)
    {
        $publicFolderEntryId = $($publicFolders[$index].EntryId);
        $folderSizeFromStats = $($publicFolders[$index].TotalItemSize.Value.ToBytes());
        $folderPath = $script:IdToNameMap[$publicFolderEntryId];
        $existingFolder = $script:FolderStatistics[$publicFolderEntryId];
        if (($existingFolder -eq $null) -or ($folderSizeFromStats -gt $existingFolder[0]))
        {
            $newFolder = @();
            $newFolder += $folderSizeFromStats;
            $newFolder += $folderPath;
            $script:FolderStatistics[$publicFolderEntryId] = $newFolder;
        }
       
        $index++;
    }    
}

# Function that creates folder objects in right way for exporting
function CreateFolderObjects()
{   
    $index = 1;
    foreach ($publicFolderEntryId in $script:FolderStatistics.Keys)
    {
        $existingFolder = $script:NonIpmSubtreeFolders[$publicFolderEntryId];
        $publicFolderIdentity = "";
        if ($existingFolder -ne $null)
        {
            $result = IsSkippableFolder($existingFolder);
            if (!$result)
            {
                $publicFolderIdentity = "\NON_IPM_SUBTREE\" + $script:FolderStatistics[$publicFolderEntryId][1];
                $folderSize = $script:FolderStatistics[$publicFolderEntryId][0];
            }
        }  
        else
        {
            $publicFolderIdentity = "\IPM_SUBTREE" + $script:FolderStatistics[$publicFolderEntryId][1];
            $folderSize = $script:FolderStatistics[$publicFolderEntryId][0];
        }  
        
        if ($publicFolderIdentity -ne "")
        {
            if(($index%10000) -eq 0)
            {
                Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.ProcessedFolders -f $index);
            }
            
            # Create a folder object to be exported to a CSV
            $newFolderObject = New-Object PSObject -Property @{FolderName = $publicFolderIdentity; FolderSize = $folderSize}
            $retValue = $script:ExportFolders.Add($newFolderObject);
            $index++;
        }
    }   
}

####################################################################################################
# Script starts here
####################################################################################################

# Array of folder objects for exporting
$script:ExportFolders = $null;

# Hash table that contains the folder list (IPM_SUBTREE via Get-PublicFolderStatistics)
$script:FolderStatistics = @{};

# Hash table that contains the folder list (NON_IPM_SUBTREE via Get-PublicFolder)
$script:NonIpmSubtreeFolders = @{};

# Hash table that contains the folder list (IPM_SUBTREE via Get-PublicFolder)
$script:IpmSubtreeFolders = @{};

# Hash table EntryId to Name to map FolderPath
$script:IdToNameMap = @{};

# Recurse through IPM_SUBTREE to get the folder path foreach Public Folder
# Remarks:
# This is done so we can overcome a limitation of Get-PublicFolderStatistics
# where it fails to display Unicode chars in the FolderPath value, but 
# Get-PublicFolder properly renders these characters (as MapiIdentity)
Write-Host "[$($(Get-Date).ToString())]" $PublicFolderStatistics_LocalizedStrings.ProcessingIpmSubtree;
$ipmSubtreeFolderList = Get-PublicFolder "\" -Server $PublicFolderServer -Recurse -ResultSize:Unlimited;
$ipmSubtreeFolderList | %{ $script:IdToNameMap.Add($_.EntryId, $_.MapiIdentity.ToString()) };
Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.ProcessingIpmSubtreeComplete -f $($ipmSubtreeFolderList.Count));

# Folders that are skipped while computing statistics
$script:SkippedSubtree = @("\NON_IPM_SUBTREE\OFFLINE ADDRESS BOOK", "\NON_IPM_SUBTREE\SCHEDULE+ FREE BUSY",
                           "\NON_IPM_SUBTREE\schema-root", "\NON_IPM_SUBTREE\OWAScratchPad",
                           "\NON_IPM_SUBTREE\StoreEvents", "\NON_IPM_SUBTREE\Events Root");

Write-Host "[$($(Get-Date).ToString())]" $PublicFolderStatistics_LocalizedStrings.ProcessingNonIpmSubtree;
$nonIpmSubtreeFolderList = Get-PublicFolder "\NON_IPM_SUBTREE" -Server $PublicFolderServer -Recurse -ResultSize:Unlimited;
Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.ProcessingNonIpmSubtreeComplete -f $($nonIpmSubtreeFolderList.Count));
foreach ($nonIpmSubtreeFolder in $nonIpmSubtreeFolderList)
{
    $script:NonIpmSubtreeFolders.Add($nonIpmSubtreeFolder.EntryId, $nonIpmSubtreeFolder); 
}

# Determining the public folder database deployment in the organization
GetPublicFolderDatabases;

# Gathering statistics from each server
GatherStatistics;

# Allocating space here
$script:ExportFolders = New-Object System.Collections.ArrayList -ArgumentList ($script:FolderStatistics.Count + 3);

# Creating folder objects for exporting to a CSV
Write-Host "[$($(Get-Date).ToString())]" ($PublicFolderStatistics_LocalizedStrings.ExportStatistics -f $($script:FolderStatistics.Count));
CreateFolderObjects;

# Creating folder objects for all the skipped root folders
$newFolderObject = New-Object PSObject -Property @{FolderName = "\IPM_SUBTREE"; FolderSize = 0};
# Ignore the return value
$retValue = $script:ExportFolders.Add($newFolderObject);
$newFolderObject = New-Object PSObject -Property @{FolderName = "\NON_IPM_SUBTREE"; FolderSize = 0};
$retValue = $script:ExportFolders.Add($newFolderObject);
$newFolderObject = New-Object PSObject -Property @{FolderName = "\NON_IPM_SUBTREE\EFORMS REGISTRY"; FolderSize = 0};
$retValue = $script:ExportFolders.Add($newFolderObject);

# Export the folders to CSV file
Write-Host "[$($(Get-Date).ToString())]" $PublicFolderStatistics_LocalizedStrings.ExportToCSV;
$script:ExportFolders | Sort-Object -Property FolderName | Export-CSV -Path $ExportFile -Force -NoTypeInformation -Encoding "Unicode";

# SIG # Begin signature block
# MIIdwgYJKoZIhvcNAQcCoIIdszCCHa8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUT2/iL2TUamli0lXpQvDA1zNc
# d1OgghhkMIIEwzCCA6ugAwIBAgITMwAAAJgEWMt/IwmwngAAAAAAmDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI3
# WhcNMTcwNjMwMTkyMTI3WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjdBRkEtRTQxQy1FMTQyMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA1jclqAQB7jVZ
# CvOuH5jFixrRTGFtwMHws1sEZaA3ciobVIdWIejc5fBu3XdwRLfxjsmyou3JeTaa
# 8lqA929q2AyZ9A3ZBfxf8VqTxbu06wBj4o4g5YCsz0C/81N2ESsQZbjDxbW5sKzD
# hhT0nTzr82aepe1drjT5dvyU/AvEkCzaEDU0dZTq2Bm6NIif11GzA+OkD0bdZG+u
# 4EJRylQ4ijStGgXUpAapb0y2RtlAYvZSpLYzeFFcA/yRXacCnoD++h9r66he/Scv
# Gfd/J/5hPRCtgsbNr3vFJzBWgV9zVqmWOvZBPGpLhCLglTh0stPa/ZxZjTS/nKJL
# a7MZId131QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFPPCI5/SvSWNvaj1nBvoSHO7
# 6ZPBMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAD+xPVIhFl30XEe39rlgUqCCr2fXR9o0aL0Oioap6LAUMXLK
# 4B+/L2c+BgV32joU6vMChTaA+7XEw7pXCRN+uD8ul4ifHrdAOEEqOTBD7N5203u2
# LN667/WY71purP2ezNB1y+YAgjawEt6VjjQcSGZ9bTPRtS2JPS5BS868paym355u
# 16HMxwmhlv1klX6nfVOs1DYK5cZUrPAblCZEWzGab8j9d2ZIGLQmTEmStdslOq79
# vujEI0nqDnJBusUGi28Kh3Hz1QIHM5UZg/F5sWgt0EobFGHmk4KH2vreGZArtCIB
# amDc5cIJ48na9GfA2jqJLWsbvNcwC486g5cauwkwggYHMIID76ADAgECAgphFmg0
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUbyqwSKeN+rpNLtDvyKVgkhPrAGgwfAYKKwYB
# BAGCNwIBDDFuMGygRIBCAEUAeABwAG8AcgB0AC0AUAB1AGIAbABpAGMARgBvAGwA
# ZABlAHIAUwB0AGEAdABpAHMAdABpAGMAcwAuAHAAcwAxoSSAImh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEACD/e421p
# pxf94SfCkEfChCyifa/vekEgVGRPnlRE3T4mJqGwgIDuNgJIyz2bH774zClzICEg
# CDGJ7llGMbsbwZhXRQ9g0IZCYUqt5auGF9pl5xq80UhLNV0728Yk75SvgUvYq7dv
# qk/hz4NlAXxevB9vhmg6xFYVxjSGPeieNtPjX//vs0e6iFdaRm0Caov7OVXslGiM
# bWT/gXZx2zTT5uqYnKuFUjs4HTAjBZUkE/8TZIFIggCG763Pf1UFTDgJpY1CQzBv
# GbaoE4Ouzg3FRUOVTxF5RdbOFdT6GAgHtBioeV3YbP68YGH3j6DRqJMgLbpRa93P
# nhZXJa0pi81+p6GCAigwggIkBgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAACYBFjLfyMJsJ4AAAAAAJgwCQYFKw4D
# AhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8X
# DTE2MDkwMzE4NDMxN1owIwYJKoZIhvcNAQkEMRYEFKIot9xBtbMj8+GhDgYbVsoP
# Q6LZMA0GCSqGSIb3DQEBBQUABIIBAIXiYAlyl/sc0K4QdygswdQVECkXbS37YLbj
# r2jXFLSDwKSRzFm2EZQn06j27xs2Xau+f8DqnTiY2lkcSQc6FJgdKt37ZsKPHbTk
# 13LQXP2VuOgrWoNmwrY8Hw5bimSVycdT63+2onQFbSvPCuTrJv/kIiCS0WSxm52g
# iGGM1vRPWjN4jcb+HYa7qVM66KUJabCNHAL3fnzSxL3YsEQZkXxwMyG8RLsnIIC+
# 8A7jsxxIph8WvPD3AesFR61PhjrjK3BoeKaH2/D+hLepTmoARYbRYG7ml1/WvShh
# 8Rpzy07XomTy+JbOy9B9NP8r2M7XYAa+7XPoRsNxzHBF2cEh65s=
# SIG # End signature block
