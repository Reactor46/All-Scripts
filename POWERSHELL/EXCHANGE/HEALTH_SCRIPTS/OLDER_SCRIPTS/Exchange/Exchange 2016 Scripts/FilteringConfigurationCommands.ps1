# Copyright (c) Microsoft Corporation. All rights reserved.

# Finds and backs up the FIPS configuration file
# Returns the full path to the FIPS configuration file
# Returns null on failure
function FindAndBackupFipsConfig
{
    # find the configuration file and make sure it is there
    if ($exchangeInstalled)
    {    
        $dbpath = (Get-Item -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ExchangeServer\v15\FIP-FS).GetValue("DatabasePath")
    }
    else
    {
        $dbpath = (Get-Item -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\FIP-FS).GetValue("DatabasePath")
    }
    
    $configFilePath = Join-Path $dbpath "Configuration.xml"
    if (!(Test-Path $configFilePath))
    {
        return null
    }

    # backup the configuration file and return the path
    Copy-Item $configFilePath ($configFilePath + ".bak")
    return $configFilePath
}

# Finds and returns the "Module" nodes from an xml object containing FIPS configuration
function GetModuleNodesFromFipsConfig
{
    param([xml] $ConfigXml)

    $configXmlNamespaces =
    @{ 
        root = "http://schemas.microsoft.com/forefront/2010/1/fs-configuration";
        sys = "http://schemas.microsoft.com/forefront/2010/1/fs-systemconfiguration" 
    }
    $modules = $ConfigXml | Select-Xml -Namespace $configXmlNamespaces -XPath "/root:Configuration/sys:System/sys:TextExtractionSettings/sys:ModuleLists/sys:ModuleList/sys:Module"
    return $modules.Node
}

# Transforms $originalNodes so that OutsideIn is disabled
$outsideInModule = "OutsideInModule.dll"

# Modifies $OriginalNodes by changing the value of OutsideIn's "Enabled" attribute 
# to $NewValue
function SetOutsideInEnabledAttribute
{
    param($OriginalNodes, $NewValue)

    foreach ($moduleNode in $OriginalNodes)
    {
        # set the "Enabled" attribute for all OutsideIn modules
        if ($moduleNode.InnerText -eq $outsideInModule)
        {
            $moduleNode.Attributes["Enabled"].Value = $NewValue;
        }
    }
}

# Writes out the inner text of a set of elements
function WriteInnerText
{
    param($Nodes)

    foreach ($moduleNode in $Nodes)
    {
        $prefix = "F"
        if ("true" -eq $moduleNode.Attributes["Enabled"].Value)
        {
            $prefix = "T"
        }

        $message = $prefix + ": " + $moduleNode.InnerText;
        Write-Output $message
    }
}

# list services that need to be restarted as part of this change
# check if MSExchangeTransport exists (if exchange is installed)
$exchangeInstalled = !((Get-Service "MSExchangeTransport" -ErrorAction SilentlyContinue) -eq $null)
if ($exchangeInstalled)
{
    $serviceList = @("MSExchangeTransport", "FMS")
}
else
{
    $serviceList = @("FMS")
}

# Stops all impacted services
function StopServices
{
    foreach ($serviceName in $serviceList)
    {
        Stop-Service -Name $serviceName -WarningAction SilentlyContinue
    }
}

# Starts all impacted services
function StartServices
{
    for ($i = $serviceList.Length - 1; $i -ge 0; --$i)
    {
        $serviceName = $serviceList[$i]
        Start-Service -Name $serviceName -WarningAction SilentlyContinue
    }
}

# import localized data
Import-LocalizedData -BindingVariable FilteringConfigurationCommands_LocalizedStrings -FileName FilteringConfigurationCommands.strings.psd1

# Transforms text extraction module configuration
function TransformModuleConfiguration
{
    param([bool] $EnableOutsideIn, [string] $OutputFile = "", [bool] $Verbose = $False)

    # stop the transport service
    StopServices

    # find and backup the config file
    $configFilePath = FindAndBackupFipsConfig
    if (!$configFilePath)
    {
        Write-Output $FilteringConfigurationCommands_LocalizedStrings.res_info_text_extraction_configuration_could_not_be_modified
        return
    }

    # load the configuration into an xml object amd find the modules node
    [xml] $configXml = Get-Content $configFilePath
    $modules = GetModuleNodesFromFipsConfig -ConfigXml $configXml

    # if we're in verbose mode, output the original modules
    $separator = "--------------------------------------------------";
    if ($Verbose)
    {
        Write-Output $separator
        WriteInnerText -Nodes $modules
        Write-Output $separator
    }

    # perform the transform (FIPS config booleans are case-sensitive, so we need to 
    # convert the boolean value to a string that config can understand)
    $newEnabledValue = "false"
    if ($EnableOutsideIn)
    {
        $newEnabledValue = "true"
    }

    SetOutsideInEnabledAttribute -OriginalNodes $modules -NewValue $newEnabledValue

    # if we're in verbose mode, output the changed modules
    if ($Verbose)
    {
        WriteInnerText -Nodes $modules
        Write-Output $separator
    }
    
    # finally, save the configuration file
    if ($OutputFile -eq "")
    {
        $OutputFile = $configFilePath
    }

    $configXml.Save($OutputFile)

    # restart the transport service
    StartServices

    Write-Output $FilteringConfigurationCommands_LocalizedStrings.res_info_text_extraction_configuration_was_modified
}

# SIG # Begin signature block
# MIIdxAYJKoZIhvcNAQcCoIIdtTCCHbECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUowJTcpSX6EyAf+1fEV/wRRcx
# +s6gghhkMIIEwzCCA6ugAwIBAgITMwAAAJzu/hRVqV01UAAAAAAAnDANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBMowggTGAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB3jAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUDgDnOCsJYcYxvZANWnuIvOqns7QwfgYKKwYB
# BAGCNwIBDDFwMG6gRoBEAEYAaQBsAHQAZQByAGkAbgBnAEMAbwBuAGYAaQBnAHUA
# cgBhAHQAaQBvAG4AQwBvAG0AbQBhAG4AZABzAC4AcABzADGhJIAiaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQAv3MZS
# BbCbSX1g2QMKAKF1GoqhvduBhaxmKrYZO5mFIvAIgdc3kJHljNWMPUlND6MH183B
# m11SkniKK1KXkzVul5cL4TSwZ74iKB5o473uRK9ToivJ8JPknxbiezeH40XTcCeA
# bInvKutFuhwehps15B3e4YkwqGEIUVd4H3m9pblX1zgQ3QFANziP9lyNxzEgib3I
# 4WTzPnpU1jt0xudrdnY0FaCG/TnYiFS9hL6nWS60t253LFy//4scMNX2nV9tuHHo
# SJFNpEiAqKdlrZG051mLscn92GnkKaKeoVdrcq5lsTU/H7CpW8zagd4l+CUAkAp8
# Dj4v2Jli9F2LMqgloYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAJzu/hRVqV01UAAAAAAAnDAJBgUr
# DgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMTYwOTAzMTg0MzIyWjAjBgkqhkiG9w0BCQQxFgQUBcnUd6wEyUEdlGRnOsIK
# j4j2XuUwDQYJKoZIhvcNAQEFBQAEggEAviHh99pAqlYxY32SLgWWN1ntYp1dqXB6
# h8qhilbctLpBcZ1mhQxbzaXLNZfNHcTnSbus/hVtfBYMr3VRjdT15Tp5LDVXkLrn
# 7eYSo8MCInm8qFCi4LVGCpXWtORup/M747zBWP7GuwyLnyjEYsP4wuGp0rA9HQsd
# O7//3impox12X4y7wWnYvehgKxde+5AxcUEnsazDiVePxlp8j/G7z1rP4bdk7sT8
# HXf43EdwBUmWO9cl4ctKi9uZI8AgUg8mwqoOjZ6GXV0VE86WFA13XDESK7aqSmbY
# Sue/+xSYY/AMkuZoI/sIsbtqk3YnL6GHhovFDUL8MZAxN3dg+Ghh8w==
# SIG # End signature block
