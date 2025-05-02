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
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUAt5R4BPaoehRWzyXOM80ApR/
# GdigghhqMIIE2jCCA8KgAwIBAgITMwAAAR+XYwozuYPXKwAAAAABHzANBgkqhkiG
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUEAp7ApMFh+B8k5kejRMM8xFPw/gw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBADAlDfJX4XBlxWWnUjDKHm+U90BI7b8dSEksSxKvkxZc
# 1c7SjaaYG5iWhd6gE71oX8l4UcTOmkQpRItLdv9uHpzlcNvk+6+gqEHlzJenhcJz
# ywduFIF4+LwQDYw7ryGhHp/0vau782TGjYVCy94d19qrtmJsIpdvUPHnEdkcQc4D
# v7fuIbv/wLMt0HTs6ncueaZUD8rgUTne0HAxxLQi2s2ylxQcLxozSOcsIUXCJdvg
# PavDuWOztdELopbX3+2wR63r05J1fzQC+oxatxdCCwNQ/cew2PEL3Lkcb7DBoEN+
# /07+pZhoN5iDYpXhjkr3U5t6mIRRNhdaoc15IieEafihggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# H5djCjO5g9crAAAAAAEfMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI1NThaMCMGCSqGSIb3DQEJ
# BDEWBBQnyoFO9NGc5oO2HC2bjZqQD301YzANBgkqhkiG9w0BAQUFAASCAQB22dcL
# 2tl18fAqoVRE59DhJnXuw2HtG0k0sv3euR5Ar6+3e2/QxldijCMfsXASZdcPSXew
# CMEJeH5ZRh5fJTntrlg93FkTBKXLVFb2AEQiEzptIT6+hUStuTOgEsnKJQDsr6oP
# R50FHrMB1vwLZbdSWB7rwaAcS+ZHW8EDOY09oJZ73CrX7J/Qokum7o+eaxx21R1n
# MDeE+u2mb6k458aa9kVKuBllcnUBVOf4ByN4Bf32bTLEMH9OZNTnsCRA+bKXxRHM
# 9PVu4+9K2UytNC4AIERGHcDXxEO6Ywu6DvmHuQKGcht+ML1JPD0G9wlhnnp39PmT
# XVfIX08whK2MU1Dt
# SIG # End signature block
