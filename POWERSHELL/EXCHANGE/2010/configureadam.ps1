<#
.EXTERNALHELP ConfigureAdam-help.xml
#>

# Copyright (c) Microsoft Corporation. All rights reserved.  
# 
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

param($ldapPort, $sslPort, $logPath, $dataPath)

#load hashtable of localized string
Import-LocalizedData -BindingVariable ConfigureAdam_LocalizedStrings -FileName ConfigureAdam.strings.psd1

# Adam Service Name
$AdamServiceName = "ADAM_MSExchange"

# Transport Service Name
$EdgeTransportServiceName = "MSExchangeTransport"

# Adam Instance Name
$AdamInstanceName = "MSExchange"

# Activate Instance command string
$ActivateInstanceCommand = "'Activate Instance " + $AdamInstanceName + "'"

# Get the windows directry.
$sysDir = [System.Environment]::SystemDirectory
$windowsDir = [System.IO.Path]::GetDirectoryName($sysDir)

# Path to the dsdbutil
$dsdbutil = $windowsDir + "\Adam\dsdbutil.exe"

# Registry location of the ADAM Config
$adamConfigRegistryPath = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\EdgeTransportRole\AdamSettings\MSExchange"

# LDAP Registry Entry Name
$ldapPortRegistryEntry = "LdapPort"

# SSL Registry Entry Name
$sslPortRegistryEntry = "SslPort"

# ADAM Data Log Path Registry Entry
$logFilesPathRegistryEntry = "LogFilesPath"

# ADAM Data File Path Registry Entry Name
$dataFilesPathRegistryEntry = "DataFilesPath"

######################################################
# Retrieves the ADAM Regsitry value of the given entry.
# adamProperty: Adam Setting to the retrieved
# returns: return entry value of found else null
######################################################

function GetAdamRegistryEntry([String] $adamProperty)
{
    $adamEntries = get-itemproperty $adamConfigRegistryPath
    if($adamEntries)
    {
        return $adamEntries.($adamProperty)
    }
    return $null
}

######################################################
# Verifies whether the sequence of commands failed or succeeded.
# statusLog: Contains the status of the sequence of commands executed
# returns: True if successs else false.
######################################################
function VerifyStatusLog([Object] $statusLog)
{
    $status = $false
    foreach($log in $statusLog)
    {
        # Very Nasty way of doing this check.
        # Other way of doing is pass the entire success message and compare it.
        # either way it still ends with comparing with some string.
        # "SUCCESSFULLY" is seen only if the entire command is succeeded.
        # Opened a Bug to track this(Bug 59676)
        if($log.ToString().ToUpper().Contains("SUCCESSFULLY"))
        {
            $status = $true
            break 
        }
    }

    if($status -eq $false)
    {
        write-host $statusLog
    }
    return $status
}
######################################################
# Set the ADAM Regsitry value of the given entry.
# adamProperty: Adam Setting to the changed
# adamPropertyValue: Value for the Entry.
######################################################

function SetAdamRegistryEntry([String] $adamProperty, [String] $adamPropertyValue)
{
    set-itemproperty $adamConfigRegistryPath -name $adamProperty -value $adamPropertyValue
}

######################################################
# Changes the LDAP Port of the ADAM. if the new value
# is same as the existing value do nothing. Also updates
# the registry Entry of the ADAM Section in Exchange
# ldapport: LDAP Port Value
######################################################

function ChangeLDAPPort([int] $ldapPort)
{
    $currentValue = GetAdamRegistryEntry($ldapPortRegistryEntry)
    if($currentValue -eq $null)
    {
        write-host $ConfigureAdam_LocalizedStrings.res_0000
        return
    }
    elseif($currentValue -eq $ldapPort)
    {
        write-host ($ConfigureAdam_LocalizedStrings.res_0001 -f $ldapPort)
        return
    }

    $ldapPortCommand = "'LDAP Port " + $ldapPort + "'"
    $quit = "'quit'"
    $command = $dsdbutil + 
                        " " +
                        $ActivateInstanceCommand + 
                        " " + $ldapPortCommand + 
                        " " + $quit

    $statusLog = invoke-expression($command)
    $isSuccess = VerifyStatusLog($statusLog)
    if($isSuccess -eq $true)
    {
        SetAdamRegistryEntry $ldapPortRegistryEntry $ldapPort
        write-host ($ConfigureAdam_LocalizedStrings.res_0002 -f $ldapPort)
    }
    else
    {
        write-host ($ConfigureAdam_LocalizedStrings.res_0003 -f $ldapPort)
    }
}

######################################################
# Changes the SSL Port of the ADAM. if the new value
# is same as the existing value do nothing. Also updates
# the registry Entry of the ADAM Section in Exchange
# sslport: SSL Port Value
######################################################

function ChangeSSLPort([int] $sslPort)
{
    $currentValue = GetAdamRegistryEntry($sslPortRegistryEntry)
    if($currentValue -eq $null)
    {
        write-host $ConfigureAdam_LocalizedStrings.res_0004
        return
    }
    elseif($currentValue -eq $sslPort)
    {
        write-host ($ConfigureAdam_LocalizedStrings.res_0005 -f $sslport)
        return
    }
    $sslPortCommand = "'SSL Port " + $sslPort + "'"
    $quit = "'quit'"
    $statusLog = invoke-expression($dsdbutil + 
                        " " + 
                        $ActivateInstanceCommand + 
                        " " + $sslPortCommand + 
                        " " + $quit)
    $isSuccess = VerifyStatusLog($statusLog)
    if($isSuccess -eq $true)
    {
        SetAdamRegistryEntry $sslPortRegistryEntry $sslPort
        write-host ($ConfigureAdam_LocalizedStrings.res_0006 -f $sslPort)
    }
    else
    {
        write-host ($ConfigureAdam_LocalizedStrings.res_0007 -f $sslPort)
    }
}

######################################################
# Changes the Log Path of the ADAM. if the new value
# is same as the existing value do nothing. Also updates
# the registry Entry of the ADAM Section in Exchange
# dataLogPath: Log files path
######################################################

function ChangeLogPath([String] $logPath)
{
    $currentValue = GetAdamRegistryEntry($logFilesPathRegistryEntry)
    if($currentValue -eq $null)
    {
        write-host $ConfigureAdam_LocalizedStrings.res_0008
    return
    }
    elseif($currentValue -eq $logPath)
    {
        write-host ($ConfigureAdam_LocalizedStrings.res_0009 -f $logPath)
        return
    }
    $filesCommand = "'Files'"
    $quit = "'quit'"
    $setlogpathCommand = "'set path logs `\`"" + $logPath + "`\`"'"
    $command = $dsdbutil + " " + 
               $ActivateInstanceCommand + " " + 
               $filesCommand + " " + $setlogpathCommand + " " +
               $quit + " " + $quit
    $statusLog = invoke-expression($command)
    $isSuccess = VerifyStatusLog($statusLog)
    if($isSuccess -eq $true)
    {
        SetAdamRegistryEntry $logFilesPathRegistryEntry $logPath
        write-host ($ConfigureAdam_LocalizedStrings.res_0010 -f $logPath)
    }
    else
    {
        write-host ($ConfigureAdam_LocalizedStrings.res_0011 -f $logPath)
    }
}

######################################################
# Changes the DATA File Path of the ADAM. if the new value
# is same as the existing value do nothing. Also updates
# the registry Entry of the ADAM Section in Exchange
# dataLogPath: Data file path
######################################################

function ChangeDataPath([String] $dataPath)
{
    $currentValue = GetAdamRegistryEntry($dataFilesPathRegistryEntry)
    if($currentValue -eq $null)
    {
        write-host $ConfigureAdam_LocalizedStrings.res_0012
    return
    }
    elseif($currentValue -eq $dataPath)
    {
        write-host ($ConfigureAdam_LocalizedStrings.res_0013 -f $dataPath)
        return
    }
    $filesCommand = "'Files'"
    $quit = "'quit'"
    $setdatafilepathCommand = "'Move DB to `\`"" + $dataPath + "`\`"'"
    $command = $dsdbutil + " " + 
               $ActivateInstanceCommand + " " + 
               $filesCommand + " " + 
               $setdatafilepathCommand + " " + 
               $quit + " " + $quit
    $statusLog = invoke-expression($command)
    $isSuccess = verifyStatusLog($statusLog)
    if($isSuccess -eq $true)
    {
        setAdamRegistryEntry $dataFilesPathRegistryEntry $dataPath
        write-host ($ConfigureAdam_LocalizedStrings.res_0014 -f $dataPath)
    }
    else
    {
        write-host ($ConfigureAdam_LocalizedStrings.res_0015 -f $dataPath)
    }
}

######################################################
# Validate the input data.
# return: return false if no input values are passed
#         else true.
######################################################

function ValidateInput()
{
    if(($ldapport -eq $null) -and
       ($sslport -eq $null) -and
       ($logPath -eq $null) -and
       ($dataPath -eq $null))
    {
        write-host $ConfigureAdam_LocalizedStrings.res_0016
        write-host $ConfigureAdam_LocalizedStrings.res_0017

        return $false
    }
    return $true
}


#############################################################
# Script starts here. Validate the input.
# Calls the methods to set the each config property
# to be set.
#############################################################

$isValidInput = ValidateInput
if($isValidInput -eq $false)
{
    return
}

stop-service -Name:$AdamServiceName -Force:$true

$adamServiceStatus = get-service $AdamServiceName

if(($adamServiceStatus -eq $null) -or 
   ($adamServiceStatus.Status -ne [System.ServiceProcess.ServiceControllerStatus]::Stopped))
{
    write-host $ConfigureAdam_LocalizedStrings.res_0018
    return
}

if($ldapport)
{
    $status = ChangeLDAPPort $ldapPort
}

if($sslport)
{
    $status = ChangeSSLPort $sslPort
}

if($logPath)
{
    $status = ChangeLogPath $logPath
}

if($dataPath)
{
    $status = ChangeDataPath $dataPath
}

start-service $AdamServiceName

start-service $EdgeTransportServiceName

# SIG # Begin signature block
# MIIaVgYJKoZIhvcNAQcCoIIaRzCCGkMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUAt5R4BPaoehRWzyXOM80ApR/
# GdigghUmMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggS6MIIDoqADAgEC
# AgphAo5CAAAAAAAfMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQTAeFw0xMjAxMDkyMjI1NThaFw0xMzA0MDkyMjI1NThaMIGzMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMScwJQYD
# VQQLEx5uQ2lwaGVyIERTRSBFU046RjUyOC0zNzc3LThBNzYxJTAjBgNVBAMTHE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQCW7I5HTVTCXJWA104LPb+XQ8NL42BnES8BTQzY0UYvEEDeC6RQ
# UhKIC0N6LT/uSG5mx5HmA8pu7HmpaiObzWKezWqkP+ejQ/9iR6G0ukT630DBhVR+
# 6KCnLEMjm1IfMjX0/ppWn41jd3swngozhXIbykrIzCXN210RLsewjPGPQ0hHBbV6
# IAvl8+/BuvSz2M04j/shqj0KbYUX0MrnhgPAM4O1JcTMWpzEw9piJU1TJRRhj/sb
# 4Oz3R8aAReY1UyM2d8qw3ZgrOcB1NQ/dgUwhPXYwxbKwZXMpSCfYwtKwhEe7eLrV
# dAPe10sZ91PeeNqG92GIJjO0R8agVIgVKyx1AgMBAAGjggEJMIIBBTAdBgNVHQ4E
# FgQUL+hGyGjTbk+yINDeiU7xR+5IwfIwHwYDVR0jBBgwFoAUIzT42VJGcArtQPt2
# +7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3NvZnQu
# Y29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNybDBY
# BggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNVHSUE
# DDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAc/99Lp3NjYgrfH3jXhVx
# 6Whi8Ai2Q1bEXEotaNj5SBGR8xGchewS1FSgdak4oVl/de7G9TTYVKTi0Mx8l6uT
# dTCXBx0EUyw2f3/xQB4Mm4DiEgogOjHAB3Vn4Po0nOyI+1cc5VhiIJBFL11FqciO
# s3xybRAnxUvYb6KoErNtNSNn+izbJS25XbEeBedDKD6cBXZ38SXeBUcZbd5JhaHa
# SksIRiE1qHU2TLezCKrftyvZvipq/d81F8w/DMfdBs9OlCRjIAsuJK5fQ0QSelzd
# N9ukRbOROhJXfeNHxmbTz5xGVvRMB7HgDKrV9tU8ouC11PgcfgRVEGsY9JHNUaeV
# ZTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEFBQAwXzETMBEG
# CgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsG
# A1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5MB4XDTEw
# MDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENvZGUgU2lnbmlu
# ZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCycllcGTBkvx2a
# YCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPVcgDbNVcKicqu
# IEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlcRdyvrT3gKGiX
# GqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZC/6SdCnidi9U
# 3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgGhVxOVoIoKgUy
# t0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdcpReejcsRj1Y8
# wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTL
# EejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYBBAGCNxUBBAUC
# AwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy82C0wGQYJKwYB
# BAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBWJ5flJRP8KuEK
# U5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNyb3NvZnQuY29t
# L3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3JsMFQGCCsGAQUF
# BwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcNAQEFBQADggIB
# AFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGjI8x8UJiAIV2s
# PS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbNLeNK0rxw56gN
# ogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y4k74jKHK6BOl
# kU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnpo1hW3ZsCRUQv
# X/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6H0q70eFW6NB4
# lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20OE049fClInHLR
# 82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8Z4L5UrKNMxZl
# Hg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9GuwdgR2VgQE6w
# QuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXrilUEnacOTj5XJ
# jdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEvmtzjcT3XAH5i
# R9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCCA++gAwIBAgIK
# YRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPyLGQBGRYDY29t
# MRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRNaWNyb3NvZnQg
# Um9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1MzA5WhcNMjEw
# NDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZUSNQrc7dGE4k
# D+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cOBJjwicwfyzMk
# h53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn1yjcRlOwhtDl
# KEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3U21StEWQn0gA
# SkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG7bfeI0a7xC1U
# n68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMBAAGjggGrMIIB
# pzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A+3b7syuwwzWz
# DzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1UdIwSBkDCBjYAU
# DqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/IsZAEZFgNjb20x
# GTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBS
# b290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0BxMuZTBQBgNV
# HR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9w
# cm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEESDBGMEQG
# CCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01p
# Y3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
# 9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwTq86+e4+4LtQS
# ooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+jwoFyI1I4vBT
# Fd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwYTp2Oawpylbih
# OZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPfwgphjvDXuBfr
# Tot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5ZlizLS/n+YWG
# zFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8csu89Ds+X57H21
# 46SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUwZuhCEl4ayJ4i
# IdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHudiG/m4LBJ1S2
# sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9La9Zj7jkIeW1
# sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g74TKIdbrHk/J
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0xggSa
# MIIElgIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQITMwAAAJ0ejSeu
# uPPYOAABAAAAnTAJBgUrDgMCGgUAoIG8MBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBQQCnsCkwWH4HyTmR6NEwzzEU/D+DBcBgorBgEEAYI3AgEMMU4wTKAkgCIAQwBv
# AG4AZgBpAGcAdQByAGUAQQBkAGEAbQAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAjQe3mzGYAHkj
# JHx2RzLh0K1SWfC1E+WM+R30Uh2+VV+Mdt8fR060oM5Y/3dUWwQoi+FvU4xPt7Nv
# beVgqszIeM2hbzstWg5i19XEGpiPRvtpiCcW7hedZv33DmQOsj/Tznb7lTYucseE
# 4yP9FxbagDRSsmiYLKihM74IsrIIldPtGcpZb8OUrysoLpvhDdajwXBmIW4pHQie
# JyR54iihVxAt6m2AItrv+chC2mvexVgcKERIZrMkJAHdXCTAHqyOz6sBe0aZVPHa
# G1yuG8MC+vXs1rNNfdkZ+yDHvtn/AuNU7a6qAErtVH+oIwCUuU5lDM36dITmsw1Z
# +c5gRbsX5KGCAh8wggIbBgkqhkiG9w0BCQYxggIMMIICCAIBATCBhTB3MQswCQYD
# VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEe
# MBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3Nv
# ZnQgVGltZS1TdGFtcCBQQ0ECCmECjkIAAAAAAB8wCQYFKw4DAhoFAKBdMBgGCSqG
# SIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTEzMDIwNTA2Mzcy
# MVowIwYJKoZIhvcNAQkEMRYEFFGeIIejd2LfMfZRdfIDrTqEQpVEMA0GCSqGSIb3
# DQEBBQUABIIBAD7IdBdZNHGviBU8/p5kdOmjxj5FvidRWlr5hL/bZ7zOQKeoN7cy
# swQr9TswvFW+TnovSEX+u6LYKCDWry+HA4iFcncZIX1tC4qHdUitWHxjp4CYNx31
# vQ5v15eLjoJg5znl0B80Fp+9RlIch5qpcbjlNOmzN3rTCLauLbTW42VTQ3lm8glx
# Jbu0HbNeoS41FlZjorG6zfS8I3Won4dfIlttA8ss7eUx6dxzmQVyqxLoDgFdxgbX
# qykbq8Fjy9faHhoSexznBrwahmPH60dsKuKtmChl5RjM2FITOIecUjR/JFFvBvTt
# IXQa2KQ2k3aQ0iL0+72OLWbmOioIWt4I87I=
# SIG # End signature block
