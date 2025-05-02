# Globals
$ComputerName = [string]$Env:computername
$setupRegistryPath = Get-ItemProperty -path 'HKLM:SOFTWARE\Microsoft\ExchangeServer\v15\Setup'
$exchangeInstallPath = $setupRegistryPath.MsiInstallPath


$AutoDiscoverPath =  "Default Web Site/Autodiscover"
$EwsPath = "Default Web Site/EWS"
$EcpPath = "Default Web Site/ECP"
$OabPath = "Default Web Site/OAB"
$MapiPath = "Default Web Site/Mapi"

# Initialize IIS metabase management object
$InitWebAdmin = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Web.Administration") 
$Iis = new-object Microsoft.Web.Administration.ServerManager 

# Adds a registry key under HKLM\Software\Microsoft\Rpc\RpcProxy to signal the RpcHttp servicelet
# that it needs to create the RpcWithCert Vdir
function ConfigureRpcWithCert
{
    $registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', '.')		
    $RpcProxyKey = $registry.OpenSubKey("Software\Microsoft\Rpc\RpcProxy", $True)
    $RpcProxyKey.SetValue("EnableRpcWithCert", 1)
    Write-Output "Registry updated, servicelet should configure RpcWithCert"
}

# Updates path SSL Flags and enables client-cert AD mapping 
function EnableClientCertAuthForPath ([string]$IisPath)
{
    Write-Output "Enabling Request-Client-Certs + AD Cert Mapping for: $IisPath."
    $config = $Iis.GetApplicationHostConfiguration();
    
    # Set SslFlags to require SSL and allow - not require client certificate auth 
    $accessSection = $config.GetSection("system.webServer/security/access", $IisPath)
    $accessSection["sslFlags"] = "Ssl, SslNegotiateCert"
    
    # Enable certificate-to-AD object mapping
    $clientCertificateMappingAuthenticationSection = $config.GetSection("system.webServer/security/authentication/clientCertificateMappingAuthentication", $IisPath)
    $clientCertificateMappingAuthenticationSection["enabled"] = $true
    $Iis.CommitChanges()
}

# Updates path to enable client-cert AD mapping 
function EnableAdClientCertAuthForPath([string]$IisPath)
{
    $config = $Iis.GetApplicationHostConfiguration();
    if ($IisPath -eq "")
    {
        Write-Output "Enabling AD Cert Mapping feature in IIS."
        $clientCertificateMappingAuthenticationSection = $config.GetSection("system.webServer/security/authentication/clientCertificateMappingAuthentication")
    }
    else
    {
        Write-Output "Enabling AD Cert Mapping for: $IisPath."
        $clientCertificateMappingAuthenticationSection = $config.GetSection("system.webServer/security/authentication/clientCertificateMappingAuthentication", $IisPath)
    }

    $clientCertificateMappingAuthenticationSection["enabled"] = $true
    $Iis.CommitChanges()
}

# Loads OAB auth module by updating web.config for OAB vdir
function UpdateOabWebConfig()
{
    if (Get-WebManagedModule -PSPath "iis:\sites\Default Web Site\OAB" -Name Microsoft.Exchange.OABAuth)
    {
		Write-Output "OABAuthModule already present in OAB's web.config."
    }
    else
    {
		Begin-WebCommitDelay
		New-WebManagedModule -PSPath "iis:\sites\Default Web Site\OAB" -Name "Microsoft.Exchange.OABAuth" -Type "Microsoft.Exchange.OABAuth.OABAuthModule"
		Write-Output "Added OABAuthModule in OAB's web.config."
		End-WebCommitDelay
    }
}

# Look for SslBinding's DefaultFlags and update as necessary
function FixSslDefaultFlags
{
    [int]$HTTP_SERVICE_CONFIG_SSL_FLAG_USE_DS_MAPPER = 0x1	# Tells Schannel to map client certs to the AD    
    $registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', '.')		
    $defaultSslBinding = $registry.OpenSubKey("SYSTEM\\CurrentControlSet\\services\\HTTP\\Parameters\\SslBindingInfo\\0.0.0.0:443")
    if ($defaultSslBinding -eq $Null)
    {
        # SSL is not configured on 0.0.0.0
        Write-Output "Cannot find a wildcard HTTP site using SSL (0.0.0.0:443), skipping HTTP_SERVICE_CONFIG_SSL_FLAG_USE_DS_MAPPER check"
    }
    else
    {
        [int]$defaultFlags = $defaultSslBinding.GetValue("defaultflags")     
        if (($defaultFlags -band $HTTP_SERVICE_CONFIG_SSL_FLAG_USE_DS_MAPPER) -ne $HTTP_SERVICE_CONFIG_SSL_FLAG_USE_DS_MAPPER)
        {
            # Bit isn't set - so do it.
            $defaultFlags = $defaultFlags -bor $HTTP_SERVICE_CONFIG_SSL_FLAG_USE_DS_MAPPER 
            
            # Need to set value to $HTTP_SERVICE_CONFIG_SSL_FLAG_USE_DS_MAPPER and then restart IIS and HTTP.SYS
            Write-Output "SChannel AD certificate mapping registry setting needs to be updated. Shutting down IIS and HTTP.SYS."
            iisreset /stop
            net stop http -force
            $defaultSslBinding = $registry.OpenSubKey("SYSTEM\\CurrentControlSet\\services\\HTTP\\Parameters\\SslBindingInfo\\0.0.0.0:443", $True)
            $defaultSslBinding.SetValue("defaultflags", $defaultFlags)
            Write-Output "Registry updated, Restarting IIS and HTTP.SYS."
            iisreset /start	
        }
    }
}

# Look for SslBinding's DefaultFlags and update as necessary
function FixValidPorts
{    
    $registry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', '.')		
    $rpcProxyKey = $registry.OpenSubKey("SOFTWARE\\Microsoft\\Rpc\\RpcProxy")
    
    if ($rpcProxyKey -eq $Null)
    {
        # RPC/HTTP component is not correctly installed
        Write-Warning "RPC over HTTP Proxy feature is not correctly installed.  Please use Server Manager to reinstall this Windows Feature."
        break
    }

    $validPorts = $rpcProxyKey.GetValue("Validports_Autoconfig_Exchange") 
    
    if ($validPorts -eq $null)
    {
        # enable-OutlookAnywhere was likely only recently enabled, add stub value and restart dependent services
        $rpcProxyKey = $registry.OpenSubKey("SOFTWARE\\Microsoft\\Rpc\\RpcProxy", $true)
        $rpcProxyKey.SetValue("Validports_Autoconfig_Exchange", "")  # set a stub value
        restart-service MSExchangeServiceHost
        restart-service MSExchangeFBA
    }
}

function TestOabVdirExchangeConfig
{
    Write-Output "" "Checking Exchange OAB virtual directory objects for $ComputerName..."
    [bool]$ShouldBreak = $False
    $OabVdir = Get-OabVirtualDirectory -Server $ComputerName
    [System.Uri]$internalUrl = $OabVdir.InternalUrl
    [System.Uri]$externalUrl = $OabVdir.ExternalUrl
    if (($internalUrl -ne $Null) -and ($internalUrl.scheme -ne "https"))
    {
        Write-Output "$ComputerName OAB internal URL does not use 'https' scheme and must be configured for SSL."
        $ShouldBreak = $True
    }
    else
    {
        Write-Output "$ComputerName OAB internal URL using SSL."
    }
    Write-Output "Current Internal Url: $internalUrl"
    if (($externalUrl -ne $Null) -and ($externalUrl.scheme -ne "https"))
    {
        Write-Output "$ComputerName OAB exernal URL does not use 'https' scheme and must be configured for SSL."
        $ShouldBreak = $True
    }
    else
    {
        Write-Output "$ComputerName OAB external URL using SSL."
    }
    Write-Output "Current External Url: $externalUrl"
    if ($ShouldBreak) 
    {
        Write-Warning "SSL is not configured for OAB access."
        Write-Output "Use the Set-OabVirtualDirectory cmdlet to configure 'https' for InternalUrl and ExternalUrl endpoints."
        Write-Output "Then, re-run Enable-OutlookCertificateAuthentication.ps1 to configure certificate authentication."
        break;
    }
}

# Main
Write-Output "Configuring client certificate authentication for OutlookAnywhere on $ComputerName..." ""

# Test for OutlookAnywhere on current machine
if (Get-OutlookAnywhere -Server $ComputerName)
{
    Write-Output "OutlookAnywhere is configured on current machine."
    FixValidPorts 
}
else
{
    Write-Warning "Enable-OutlookAnywhere must be run before configuring client certificate authentication.  Exiting."
    break
}

# Check on OAB URL configuration first
TestOabVdirExchangeConfig

# Is the OS version Windows8 or higher
function IsOSWin8OrHigher
{
    $version = [System.Environment]::OSVersion.Version
    $windows8Version = [version]"6.2.0.0"
    return $version -ge $windows8Version
}

# Install IIS Client certificate-to-AD authentication mapping if necessary
if (IsOSWin8OrHigher)
{
    Powershell -Command {
        Import-Module ServerManager
        if (-not $(Get-WindowsFeature Web-Client-Auth).Installed)
        {
            Write-Output "Enabling IIS AD Certificate mapping module"
            Add-WindowsFeature -Name Web-Client-Auth -IncludeAllSubFeature
        }
        else
        {
            Write-Output "IIS AD Certificate mapping module already installed."
        }
    
    }
}
else
{
    $winFeatures = ServerManagerCmd.exe -query
    foreach($winFeature in $winFeatures)
    {
        if ($winFeature.Contains("[ ] Client Certificate Mapping Authentication"))
        {
            Write-Output "Enabling IIS AD Certificate mapping module"
            ServerManagerCmd.exe -install Web-Client-Auth
            break
        }
    }
}

# IIS: Enable server-wide Client certificate-to-AD authentication mapping
EnableAdClientCertAuthForPath ("") 					# Global
EnableClientCertAuthForPath($AutoDiscoverPath) 		# AutoDiscover
EnableClientCertAuthForPath($EwsPath) 				# EWS
EnableClientCertAuthForPath($EcpPath) 				# ECP
EnableClientCertAuthForPath($OabPath) 				# OAB
EnableClientCertAuthForPath($MapiPath)				# Mapi

# IIS: OutlookAnywhere: Enable Client certificate-to-AD authentication mapping (client cert auth already *required* on this vdir)
ConfigureRpcWithCert

# Check on Schannel settings to ensure the 
FixSslDefaultFlags

# Update OAB add web.config to add OABAuth module
UpdateOabWebConfig

Write-Output "Done!  $ComputerName configured for OutlookAnywhere with client certificate authentication."
$a=$Iis.Dispose()

# SIG # Begin signature block
# MIId2QYJKoZIhvcNAQcCoIIdyjCCHcYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU2XTd1yVLK6UxVuuqZymKs/4T
# ikOgghhkMIIEwzCCA6ugAwIBAgITMwAAAJmqxYGfjKJ9igAAAAAAmTANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBN8wggTbAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB8zAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUCpo+ksUqTUjOUU+7WlNWtkj1ThcwgZIGCisG
# AQQBgjcCAQwxgYMwgYCgWIBWAGUAbgBhAGIAbABlAC0ATwB1AHQAbABvAG8AawBD
# AGUAcgB0AGkAZgBpAGMAYQB0AGUAQQB1AHQAaABlAG4AdABpAGMAYQB0AGkAbwBu
# AC4AcABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDAN
# BgkqhkiG9w0BAQEFAASCAQCEBUDs2K11DPXeWzq96Z31STGNz8yzglSZbcVqqelL
# rSv1KK6a2DwtNyc48GIpWtMvAblErghuqsLFI0/ERmLJ8Ovjii/VbNJBIYl7AjQt
# gz3je/KTmX5HYlIB8M9dUAq0nbosW1Bb84KLlf48ylzReLzhKbmR1jflAoZDRW6s
# VI6kcJ5rhQfqS89v5Ym45wlqX1eF4JUKUs2FsF2kX/HJ1cMukrWTpucbZ0lUbjLU
# 2bYp+SQi85VzkYq1Bjtoi8Eo9MF1rU57b9kwovBbU4aHz8Jc51llavo5R7CddALf
# zJsimDo60gXekBWtG1EqB/k9GlZ2M4veJ/cRtQVNlVN5oYICKDCCAiQGCSqGSIb3
# DQEJBjGCAhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAA
# AJmqxYGfjKJ9igAAAAAAmTAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqG
# SIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTYwOTAzMTg0MzEyWjAjBgkqhkiG9w0B
# CQQxFgQUVryhvXBvDC48f++XltvL5q6wjBswDQYJKoZIhvcNAQEFBQAEggEAPJ/3
# 7uCg+QjDMUL8ZY6uRavX/qKw5LBzXEJfbqkSarO1pZd3+Mqi2l65Ed2RBqI3Rh31
# 8Vn94bkX75B4j3wVNcY7J0bR6VlUW4Kr7DxDYVFtQEF09PxLHTMtNg01QgpReVBO
# vfq6TMM1bHJfx7raCZVk6e7UfHUQaDy84nK59fu03Im6gPWCldFPPo3YMQkHT4TV
# J6NHnFNR3wj6pq7lSsK7+lRChr1ZP/lptfdpmUR+PqMJ5xKlkKScASskHcKmEQtS
# fPVtTS7HiksKnEzY95JtK/attSK/Rla5geBF/HRHAnuRxejXDACVOcIbciKNq+qx
# HGtOoWNLj8Hm8rZ/2w==
# SIG # End signature block
