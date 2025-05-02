# .SYNOPSIS
# Import-MailPublicFolders.ps1
#    Create dummy mail public folder objects in target forest
#
# The script needs to be run from Onprem. 
#
# One of the forest involved should always be cloud.
#
# If the input do not contain the switch parameter, then the target forest will be assumed as Onprem.
#
# Default URI to connect to cloud is https://outlook.office365.com/powerShell-liveID. This can be changed by passing the appropriate URI to ConnectionUri parameter
#
# Example input to the script:
#
# Import-MailPublicFolders.ps1 -Credential <credential> -ToCloud
#
# .DESCRIPTION
#
# Copyright (c) 2012 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
param (
    [Parameter(Mandatory=$true)]
    [PSCredential] $Credential,
    
    [Parameter(Mandatory=$false)]
    [Switch] $ToCloud,    
    
    [Parameter(Mandatory=$false)]
    [string] $ConnectionUri = "https://outlook.office365.com/powerShell-liveID"
    )
    
################ DECLARING GLOBAL VARIABLES ################

$script:session = $null;

################ END OF DECLARATION #################    

## Create a tenant PSSession.
function CreateTenantSession()
{
    param ($cred, $ConnUri)
    
    $sessionOption = (New-PSSessionOption -SkipCACheck);
    $script:session = New-PSSession -ConnectionURI:$ConnUri `
                             -ConfigurationName:Microsoft.Exchange `
                             -AllowRedirection `
                             -Authentication:"Basic" `
                             -SessionOption:$sessionOption `
                             -Credential:$cred `
                             -ErrorAction:SilentlyContinue;
    if ($script:session -eq $null)
    {
        Write-Host "Input credentials are incorrect";
        exit;
    }
    else
    {
        $params = @{
                    "Session" = $script:session;
                    "Prefix" = "EXO";
                    "AllowClobber" = $true
                  }
        if ( -not (&"Import-PSSession" @params))
        {
            WriteHost "Exchange online session cannot be imported to current PowerShell window.";
            Remove-PSSession $script:session;
            exit;                
        }
    }
}

## Get organization guid
function GetOrganizationGuid()
{
    param ($targetForest)
    
    $organizationGuid = "";
    if ($targetForest)
    {
        $orgConfig = Get-OrganizationConfig -ErrorAction:SilentlyContinue;        
    }
    else
    {
        $orgConfig = Get-EXOOrganizationConfig -ErrorAction:SilentlyContinue;        
    }
    
    # Return the results     
    if ($orgConfig -ne $null)
    {
        $organizationGuid = $($orgConfig.Guid.ToString());
    }
    
    return $organizationGuid;
}

## Retrieve mail public folders
function GetMailPublicFolders()
{
    param ($fromSource, $targetForest)
    
    $mailPublicFolders = @();
    if (($fromSource -and $targetForest) -or (!$fromSource -and !$targetForest))
    {
        $mailPublicFolders = Get-MailPublicFolder -ResultSize:Unlimited -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue;
    }
    else
    {
        $mailPublicFolders = Get-EXOMailPublicFolder -ResultSize:Unlimited -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue;
    }
                 
    return $mailPublicFolders;
}

## MailPublicFolders whose external email address do not point to an
## existing object need to be removed
function RemoveOrphanMailPublicFolders()
{
    param ($srcFolderHashtable, $tgtMailPublicFolders, $targetForest, $orgGuid)
    
    foreach ($mailPublicFolder in $tgtMailPublicFolders)
    {
        if ($mailPublicFolder.ExternalEmailAddress -ne $null)
        {
            if ($srcFolderHashtable.ContainsKey($mailPublicFolder.ExternalEmailAddress.ToString()))
            {
                $srcFolderHashtable.Remove($mailPublicFolder.ExternalEmailAddress.ToString());
                continue;
            }
            elseif ($srcFolderHashtable.ContainsKey($mailPublicFolder.ExternalEmailAddress.ToString().ToUpper().Replace("SMTP:","")))
            {
                $srcFolderHashtable.Remove($mailPublicFolder.ExternalEmailAddress.ToString().ToUpper().Replace("SMTP:",""));
                continue;
            }
        }
        
        if ($mailPublicFolder.LegacyExchangeDN -ne $null -and ($mailPublicFolder.LegacyExchangeDN.Contains($orgGuid)))
        {
            if ($targetForest)
            {
                Disable-EXOMailPublicFolder $mailPublicFolder.Alias -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue -Confirm:$false;
            }
            else
            {
                Disable-MailPublicFolder $mailPublicFolder.Alias -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue -Confirm:$false;
            }
        }
    }    
    
    return $srcFolderHashtable;   
}

## Retrieve accepted domains
function GetAcceptedDomains()
{
    param ($targetForest, $session)
    
    $acceptedDomains = @();
    if ($targetForest)
    {
        $acceptedDomains = Get-EXOAcceptedDomain -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue;
    }
    else
    {
        $acceptedDomains = Get-AcceptedDomain -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue;
    }
    
    return $acceptedDomains;
}

## Import mail public folders.
function ImportMailPublicFolders()
{
    param ($targetForest, $acceptedDomains, $srcFolderHashtable, $orgGuid)

    if ($targetForest)
    {
        $cmdletToExecute = "New-EXOSyncMailPublicFolder";
    }
    else
    {
        $cmdletToExecute = "New-SyncMailPublicFolder";
        $session = $null;
    }
    
    $acceptedDomainCount = $acceptedDomains.Count;
    $inputParameters = @{};
    foreach ($mailPublicFolder in $($srcFolderHashtable.Values))
    {
        # Collect the properties of mail enabled public folder
        $alias = $mailPublicFolder.Alias.Trim();
        $name = $mailPublicFolder.Name.Trim();
        $entryId = $mailPublicFolder.EntryId.ToString();
        $windowsEmailAddress = $mailPublicFolder.WindowsEmailAddress.ToString();
        $externalEmailAddress = $mailPublicFolder.PrimarySmtpAddress.ToString();
        
        if ($alias.Length -lt 1 -or 
            $entryId.Length -lt 1 -or
            $externalEmailAddress.Length -lt 1)
        {
            continue;
        }
        
        $entryId = $orgGuid + $entryId;

        $emailAddressesArray = @($mailPublicFolder.EmailAddresses);
        
        if ($windowsEmailAddress -ne "")
        {
            $localPart = @($windowsEmailAddress.Split('@'))[0];
            for ($index = 0; $index -lt $acceptedDomainCount; $index++)
            {
                $emailAddressesArray += $localPart + "@" + $acceptedDomains[$index].DomainName.ToString();
            }
        }

        for ($index = 0; $index -lt $emailAddressesArray.Count; $index++)
        {
            $emailAddressesArray[$index] = $emailAddressesArray[$index].ToString().Replace("SMTP:","");
            $emailAddressesArray[$index] = $emailAddressesArray[$index].ToString().Replace("smtp:","");
        }
        
        # Remove duplicate email addresses if any 
        $emailAddressesArray = $emailAddressesArray | Sort-Object -Unique;
        
        $inputParameters.Clear();
        $inputParameters.Add("Name", $name);
        $inputParameters.Add("Alias", $alias);
        
        if ($($mailPublicFolder.HiddenFromAddressListsEnabled) -eq "True")
        {
            $inputParameters.Add("HiddenFromAddressListsEnabled",$true);
        }
        
        $inputParameters.Add("EmailAddresses",$emailAddressesArray);
        $inputParameters.Add("EntryId",$entryId);
        
        if ($windowsEmailAddress -ne "")
        {
            $inputParameters.Add("WindowsEmailAddress",$windowsEmailAddress);                                                          
        }
    
        if ($externalEmailAddress -ne "")
        {
            $inputParameters.Add("ExternalEmailAddress",$externalEmailAddress);                                                                                                                    
        }
        
        $inputParameters.Add("ErrorAction","Continue");
        $inputParameters.Add("WarningAction","Continue");
        
        # Execute the command
        &$cmdletToExecute @inputParameters;                     
    }
}

################################ BEGINNING OF SCRIPT ################################

# Create a PSSession for this organization
CreateTenantSession $Credential $ConnectionUri;

# Determine the guid of the organization from where to export objects
$organizationGuid = GetOrganizationGuid $ToCloud;

# Get mail enabled public folders from source forest
$sourceMailPublicFolders = @(GetMailPublicFolders $true $ToCloud);

$sourceFoldersHashTable = @{};
foreach ($mailPublicFolder in $sourceMailPublicFolders)
{
    $sourceFoldersHashTable.Add($mailPublicFolder.PrimarySmtpAddress.ToString(), $mailPublicFolder);
}

# Get mail enabled public folders from target forest
$targetMailPublicFolders = @(GetMailPublicFolders $false $ToCloud);

if ($targetMailPublicFolders.Count -gt 0)
{
    $sourceFoldersHashTable = RemoveOrphanMailPublicFolders $sourceFoldersHashTable $targetMailPublicFolders $ToCloud $organizationGuid;    
}

Write-Host "Successfully removed any existing orphan mail public folders already created by the script";

if ($sourceMailPublicFolders.Count -lt 1)
{
    if ($ToCloud)
    {
        Write-Host "Couldn't find any mail enabled public folder objects in Onprem environment";
    }
    else
    {
        Write-Host "Couldn't find any mail enabled public folder objects in Cloud environment";
    }

    Remove-PSSession $script:session;
    exit;
}

# Retrieve the accepted domains for this organization
$acceptedDomains = @(GetAcceptedDomains $ToCloud);

# Import the mail enabled public folders to other forest
ImportMailPublicFolders $ToCloud $acceptedDomains $sourceFoldersHashTable $organizationGuid;

Write-Host "Completed importing of mail enabled public folders";

# Terminate the PSSession
Remove-PSSession $script:session;

################################ END OF SCRIPT ################################

# SIG # Begin signature block
# MIIduAYJKoZIhvcNAQcCoIIdqTCCHaUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU54ngRQdKvgFQERrbm7wdx0JG
# Ca6gghhkMIIEwzCCA6ugAwIBAgITMwAAAJvgdDfLPU2NLgAAAAAAmzANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBL4wggS6AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB0jAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUiTjX4YgQQy6928bqlB2QE3G5DJswcgYKKwYB
# BAGCNwIBDDFkMGKgOoA4AEkAbQBwAG8AcgB0AC0ATQBhAGkAbABQAHUAYgBsAGkA
# YwBGAG8AbABkAGUAcgBzAC4AcABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQCPZVOnuR1mmzJ94kh8FLeq
# eXVw9f1QSCwNmWxrKJ7oLw9MXNUCqMfdRNaPA09muNxrDDGwA2XrkYYJR2VbQJwk
# iMJwv/HnHkJ+KtT0PYntbA/SolK/AGYgB153scDhVdm2nAYboOXzNW2FGCRkAbLe
# uKx5RG03AWzZiWoXgzHJdj1v9f89tsYITrXRmX9FuDjdfHIn09w3XQpsMdyQUhvP
# 30/N7dvAijoN3oL56uey7eAJ/HAB0zcDPNoQPoP0NpVrx+vVgaHf7vL32pKLyfNp
# XcJ+/zUpIvGfN11lQiVdFWmDlkYxFgl8BgOTzbE6gxXw5gfOT06QZ1gzsn5x5PCy
# oYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVT
# MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQK
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1l
# LVN0YW1wIFBDQQITMwAAAJvgdDfLPU2NLgAAAAAAmzAJBgUrDgMCGgUAoF0wGAYJ
# KoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTYwOTAzMTg0
# NDIzWjAjBgkqhkiG9w0BCQQxFgQUlBur4kbpmaEhktX642BwqUrhyekwDQYJKoZI
# hvcNAQEFBQAEggEAd9lD7bObpOknyi0vqvMD5v2d4MaC+F4gl3IcloTS4nWcwEma
# IE3+8jAoM/h3PdNjsrpvhA7qdaBy1XpT/RrbVazsqgIho4Q/P53+RYVeGUOGp9W6
# aqNR2Vodw8J+xPZxtb1oF0ChKQPx0id8amzlO/ddePP+9j34Z17OhqS8330L+ZGD
# 2ZxQphpfanYTNrPZ8+AdD6s63yFrfoGM3dxNAMST05k+H0UhhrnVhmuW66ur+xfy
# db2Gpk8DJ+o+6XQaIkDY/VTUDCjhxTxqYpjOktC6IYrHL83st13exitCERellQT6
# rHdTiqDTMyl75PmEp7QIR/2L3XBdCayS8BFa8w==
# SIG # End signature block
