<#
.EXTERNALHELP LargeToken-IIS_EWS-help.xml
#>

# Copyright (c) 2011 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.


# This script requires Powershell 2
# It uses Remote Registry Service & WS-Management on Exchange 2010 SP1 CAS servers
#
#
# It will do the following:
# 1.On all CAS Servers set:
#    HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\HTTP\Parameters
#      MaxFieldLength DWORD 65534
#      MaxRequestBytes DWORD 16777216
# Refer: http://support.microsoft.com/kb/920862
#
# 2.On Exchange 2010 SP1 CAS Servers set:
#    Under the custom bindings in the EWS web.config file set maxReceivedMessageSize to 512000000 and maxBufferSize to 163840
#    for these bindings:
#        <binding name="EWSAnonymousHttpsBinding">
#        <binding name="EWSAnonymousHttpBinding">
#        <binding name="EWSBasicHttpsBinding">
#        <binding name="EWSBasicHttpBinding"> 
#        <binding name="EWSNegotiateHttpsBinding">
#        <binding name="EWSNegotiateHttpBinding">
#    Note:  These bindings do not exist on Exchange 2010 RTM
#


$CASRole = 4

# Min VersionNumber in AD for the Exchange 2010 RTM server (it is translated from the ServerVersion object)
# It is copied from \\exsrc\sources\latest\e14sp1\sources\dev\data\src\directory\systemconfiguration\server.cs
$E14RTMVersionNumber = 0x73800000 #RTM
$E14SP1VersionNumber = 0x73810000 #SP1   -   The bindings dont exist on Exchange 2010 RTM
# $E12VersionNumber = 0x72000000 #RTM

#Values for settings
$httpMaxFieldLength = 65534
$httpMaxRequestBytes = 16777216


Function Get-ExchangeServerInSite
{

    $ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]
    $siteDN = $ADSite::GetComputerSite().GetDirectoryEntry().distinguishedName
    $configNC=([ADSI]"LDAP://RootDse").configurationNamingContext
    $search = new-object DirectoryServices.DirectorySearcher([ADSI]"LDAP://$configNC")
    $objectClass = "objectClass=msExchExchangeServer"
    $site = "msExchServerSite=$siteDN"
    $search.Filter = "(&($objectClass)($site))"

    $search.PageSize=1000
     [void] $search.PropertiesToLoad.Add("name")
     [void] $search.PropertiesToLoad.Add("msexchcurrentserverroles")
     [void] $search.PropertiesToLoad.Add("networkaddress")
     [void] $search.PropertiesToLoad.Add("versionnumber")
     [void] $search.PropertiesToLoad.Add("msexchinstallpath")

    $search.FindAll()
}


Function Set-CASServerKeysAndEWS
{

    #add all servers in the local site to an array
    $servers = New-Object System.Collections.ArrayList
    Get-ExchangeServerInSite | %{ [void]$servers.Add(($_)) }


    foreach($server in $servers)
    {
        foreach($role in $server.Properties.msexchcurrentserverroles)
        {    
            # Verify its a CAS server
            if( ($role -band $CASRole) -gt 0 )
            {
                $serverName = $server.Properties.name
                $installPath = $server.Properties.msexchinstallpath
                
                try{
                    echo "Setting Http keys on $serverName"
               
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $serverName)

                    $regKey= $reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Services\\HTTP\\Parameters" , $true)  #open for Write

                    $regKey.SetValue("MaxFieldLength" , $httpMaxFieldLength, [Microsoft.Win32.RegistryValueKind]::DWord)
                    $regKey.SetValue("MaxRequestBytes" , $httpMaxRequestBytes, [Microsoft.Win32.RegistryValueKind]::DWord)
                }
                catch
                {
                    Write-Warning "Unable to set Http keys on $serverName"
                }


                # Pick Exchange 2010  CAS servers
                if($server.Properties.versionnumber -ge $E14SP1VersionNumber)
                {
                    # Modify EWS web.config
                    Set-EWSWebConfig $serverName $installPath
                }
                elseif($server.Properties.versionnumber -ge $E14RTMVersionNumber)
                {
                    Write-Warning "Exchange 2010 RTM : Does not have the required bindings in EWS web.config - they were added in SP1"
                }
            }
        }
     }

    echo "DONE"
}



Function Set-EWSWebConfig($serverName, [string]$installPath)
{

    #block of code to be run on remote machine
    $block =
    {
        param([string]$installPath)
        $ewsMaxReceivedMessageSize = 512000000 
        $ewsmaxBufferSize = 163840
        $bindingNameList = "EWSAnonymousHttpsBinding", "EWSAnonymousHttpBinding", "EWSBasicHttpsBinding", "EWSBasicHttpBinding", "EWSNegotiateHttpsBinding", "EWSNegotiateHttpBinding"
        [int]$cnt = 0

        #$vdir = get-webservicesvirtualdirectory
        [string]$vdir = $installPath + "\ClientAccess\exchweb\EWS"
        [string]$webConfigPath = $vdir + "\web.config"
        echo "Editing $webConfigPath"

        #$webConfigPath = "D:\Program Files\Microsoft\Exchange Server\V14\ClientAccess\exchweb\EWS\web.config"

        try
        {
            #Create a backup file
            $currentDate = (get-date).tostring("mm_dd_yyyy-hh_mm_s") # month_day_year - hours_mins_seconds   7: $backup = $webConfigPath + "_$currentDate"
            $backup = $webConfigPath + "_$currentDate"

            $xml = [xml](Get-Content $webConfigPath)
            $xml.Save($backup)

            #Do the edit
            $root = $xml.get_DocumentElement()
            $elemList = $root.GetElementsByTagName("customBinding")
            foreach($element in $elemList)
            {
                foreach($binding in $element.ChildNodes)
                {
                   foreach($bindingName in $bindingNameList)
                   {
                       if($binding.Name -eq $bindingName)
                       {
                           #edit the fields
                           foreach($child in $binding.ChildNodes)
                           {
                               if($child.Name -eq "httpTransport")
                               {
                                   #echo "$bindingName"
                                   #echo $child
                                   $child.maxReceivedMessageSize = "$ewsMaxReceivedMessageSize"
                                   $child.maxBufferSize = "$ewsMaxBufferSize"
                                   $cnt++
                                   break
                               }
                               elseif($child.Name -eq "httpsTransport")
                               {
                                   #echo "$bindingName "
                                   #echo $child
                                   $child.maxReceivedMessageSize = "$ewsMaxReceivedMessageSize"
                                   $child.maxBufferSize = "$ewsMaxBufferSize"
                                   $cnt++
                                   break
                               }
                           }
                       }
                   }
                }
            }
            
            if($cnt -ne $bindingNameList.Count)
            {
                write-warning "Did not modify all of the bindings specified"
            }
            $xml.Save($webConfigPath)

            # Delete $backup
            del $backup
        }
        catch
        {
            Write-Warning "Edit of $webConfigPath failed"
        }
    }#block



    try
    {
        #Execute $block on $serverName 
        echo "Setting EWS web.config on $serverName"
        #WinRM 1 opens port 80 on quick config
        #WinRM 2 opens port 5985
        $session = new-pssession -ComputerName $serverName -port 80  -ErrorAction SilentlyContinue
        if($session -eq $null)
        {
            $session = new-pssession -ComputerName $serverName -port 5985 -ErrorAction SilentlyContinue
        }
        if($session -eq $null)
        {
            throw
        }

        #Pass $installPath as argument
        invoke-command -session $session -ArgumentList $installPath -ScriptBlock $block
        Remove-PSSession -session $session
     }
     catch
     {
        Write-Warning "Setting EWS web.config on $serverName failed. Make sure winRM service is started (winRM quickconfig)"
        throw $_
     }

    echo "DONE"
}




################################# Main ################################# 

Set-CASServerKeysAndEWS

# SIG # Begin signature block
# MIIacgYJKoZIhvcNAQcCoIIaYzCCGl8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUXwNZIVksg0cKOWFfIc3V59rj
# WRqgghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# a8kTo/0xggStMIIEqQIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIHGMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBTdqgFMy8MXpUpeZNMlFmW4ehKlYzBmBgorBgEEAYI3AgEMMVgw
# VqAugCwATABhAHIAZwBlAFQAbwBrAGUAbgAtAEkASQBTAF8ARQBXAFMALgBwAHMA
# MaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3
# DQEBAQUABIIBAB5pgMLInqsy48JY2DPibIaUcpkqaFAqi+EbDwAN/uyEnwSMKHGz
# IFha6WEOqaDXOK33cMoRmPCFtLBsWuCOtwhz6vMg9XWRhFroESUredBXSexunv1B
# QWT9tQsNaeC5Yti0cCJJmOpKoBmpbSsrhEWPAkxwtoCv1oUHvNWgNfXQmshBJmJE
# VWkZlWmh+1rXu4duJT5RHyjCFRqTk+LB55pWX6Ux5aK+jdo8BQgqhdYn/lwPzKLp
# THPLrmchn5D0M+1z9tvMQ3XGlZE6akU6MU7zmN+8vKr4M9LOjS8OuJH1xUOnrZGA
# e7o+uvX4PEAX96QWu2DZTnHljfos5yUIjiehggIoMIICJAYJKoZIhvcNAQkGMYIC
# FTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAKzkySMGy
# yUjzAAAAAAArMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0xMzAyMDUwNjM3MjRaMCMGCSqGSIb3DQEJBDEWBBTU
# +C13Jd/f7tW88317NMxhxJXb9jANBgkqhkiG9w0BAQUFAASCAQBcYmvtMWtcUycy
# 0J+WSSbQU89dDpq6PU3CpxB1//T7xD/Qoq3A/6G3aULN9l6BAyiONgVfIGrfLswg
# XxvmUSIg/N5TMo01+Xw3JfivQEXEPtM2rkuRkWt9xb6XbdSsoZX862yf6dGQ9pCd
# sVR5/6+oHB+f4FpuYitubpLOcxMcAbMXqnaZEki3p5sepHaWcTkRWzOYkbYgp7PK
# SUqy4Chqg8OgUW/n3MfHJj5YQPykH5NGwHJc0Hyla5w7+PPcstf95sYZPtfg31UZ
# /yWmXi2xZQN+6kCxt8NICK0fEoQIX7O8E4c9OUWAS11gWo24AJL8BM+q+kNFVGBR
# 3qJqbRPi
# SIG # End signature block
