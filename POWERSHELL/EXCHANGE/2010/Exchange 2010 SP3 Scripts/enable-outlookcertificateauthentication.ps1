# Globals
$ComputerName = [string]$Env:computername
$setupRegistryPath = Get-ItemProperty -path 'HKLM:SOFTWARE\Microsoft\ExchangeServer\v14\Setup'
$exchangeInstallPath = $setupRegistryPath.MsiInstallPath


$AutoDiscoverPath =  "Default Web Site/Autodiscover"
$EwsPath = "Default Web Site/EWS"
$EcpPath = "Default Web Site/ECP"
$OabPath = "Default Web Site/OAB"
$RpcHttpWithCertPath = "Default Web Site/RpcWithCert"

# Initialize IIS metabase management object
$InitWebAdmin = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Web.Administration") 
$Iis = new-object Microsoft.Web.Administration.ServerManager 

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

# Creates OAB app pool based on DefaultAppPool, running as LocalSystem
function CreateOabAppPool
{
    # Get existing OAB authentication values to set later
    $config = $Iis.GetApplicationHostConfiguration();
    $basicAuthenticationSectionEnabled = $config.GetSection("system.webServer/security/authentication/basicAuthentication", "Default Web Site/OAB")["enabled"];
    $windowsAuthenticationSectionEnabled = $config.GetSection("system.webServer/security/authentication/windowsAuthentication", "Default Web Site/OAB")["enabled"];

    $apppool = $Iis.ApplicationPools["MSExchangeOabAppPool"]
    if ($apppool)
    {
        # Delete existing app pool
        $apppool.Delete()       
        # Flush
        $Iis.CommitChanges()
    }

    # Create new app pool, then bind to it
    $a=$Iis.applicationPools.Add("MSExchangeOabAppPool")
    $apppool = $Iis.ApplicationPools["MSExchangeOabAppPool"]
    
    # Now make sure it runs as LocalSystem, and prevent unnecessary app pool restarts
    $apppool.ProcessModel.IdentityType = [Microsoft.Web.Administration.ProcessModelIdentityType]"LocalSystem"
    $apppool.ProcessModel.idleTimeout = "0.00:00:00"
    $apppool.Recycling.PeriodicRestart.time = "0.00:00:00"

    # Create /OAB application
    $OabApplication = $Iis.Sites["Default Web Site"].Applications["/OAB"]
    if ($OabApplication)
    {
        # Delete it
        $OabApplication.Delete()
        # Flush
        $Iis.CommitChanges()
    }

    $oabvdir=$Iis.Sites["Default Web Site"].Applications["/"].VirtualDirectories["/OAB"]
    if ($oabvdir)
    {
        # Clean up vdir
        $oabvdir.Delete()
        $Iis.CommitChanges()
    }
        
    $addSite=$Iis.Sites["Default Web Site"].Applications.Add("/OAB", $ExchangeInstallPath + "ClientAccess\OAB")
    $OabApplication = $Iis.Sites["Default Web Site"].Applications["/OAB"]
    if ($OabApplication -eq $Null)
    {
        # Error creating OAB vdir.  Need to fix existing one and rest
        Write-Warning "Error updating Default Web Site/OAB to support the OABAuth component."
        Write-Output "Please use IIS Manager to remove the Default Web Site/OAB virtual directory, then the following commands to recreate the OAB virtual directory:"
        Write-Output "Get-OabVirtualDirectory -server $ComputerName | Remove-OabVirtualDirectory"
        Write-Output "New-OabVirtualDirectory -server $ComputerName"
        break
    }

    #Set app pool
    $OabApplication.ApplicationPoolName = "MSExchangeOabAppPool"

    #Restore previous auth settings and enabled anonymous
    # Reload applicationHost.config
    $config = $Iis.GetApplicationHostConfiguration();
    
    # Check null (inherited from root of web server), otherwise set to previous value
    if ($basicAuthenticationSectionEnabled)
    {
        $basicAuthenticationSection = $config.GetSection("system.webServer/security/authentication/basicAuthentication", "Default Web Site/OAB")
        $basicAuthenticationSection["enabled"] = $basicAuthenticationSectionEnabled
    }
    
    # Check null (inherited from root of web server), otherwise set to previous value
    if ($windowsAuthenticationSectionEnabled)
    {
        $windowsAuthenticationSection = $config.GetSection("system.webServer/security/authentication/windowsAuthentication", "Default Web Site/OAB")
        $windowsAuthenticationSection["enabled"] = $windowsAuthenticationSectionEnabled
    }

    # Enables HTTP anonymous access; access without valid credentials is still denied as IUSR is prevented from reading OAB files
    $anonymousAuthenticationSection = $config.GetSection("system.webServer/security/authentication/anonymousAuthentication", "Default Web Site/OAB")
    $anonymousAuthenticationSection["enabled"] = "true" 

    $Iis.CommitChanges()
}

# Loads OAB auth module by creating or overwriting web.config for OAB vdir
function UpdateOabWebConfig()
{
    $webConfigPath = $ExchangeInstallPath + "ClientAccess\OAB\web.config"
    $webConfigOriginal = @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <system.webServer>
    <modules>
      <add name="Microsoft.Exchange.OABAuth" type="Microsoft.Exchange.OABAuth.OABAuthModule" />
    </modules>
  </system.webServer>
  <system.web>
    <compilation defaultLanguage="c#" debug="false">
      <assemblies>
        <add assembly="Microsoft.Exchange.Net, Version=14.0.0.0, Culture=neutral, publicKeyToken=31bf3856ad364e35"/>
        <add assembly="Microsoft.Exchange.Diagnostics, Version=14.0.0.0, Culture=neutral, publicKeyToken=31bf3856ad364e35"/>
        <add assembly="Microsoft.Exchange.OabAuthModule, Version=14.0.0.0, Culture=neutral, publicKeyToken=31bf3856ad364e35"/>
      </assemblies>
    </compilation>
  </system.web>
  <runtime>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.Exchange.OABAuthModule" publicKeyToken="31bf3856ad364e35" culture="neutral" />
        <codeBase version="14.0.0.0" href="file:///{0}bin\Microsoft.Exchange.OABAuthModule.dll"/>
      </dependentAssembly>
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.Exchange.Net" publicKeyToken="31bf3856ad364e35" culture="neutral" />
        <codeBase version="14.0.0.0" href="file:///{0}bin\Microsoft.Exchange.Net.dll"/>
      </dependentAssembly>
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.Exchange.Rpc" publicKeyToken="31bf3856ad364e35" culture="neutral" />
        <codeBase version="14.0.0.0" href="file:///{0}bin\Microsoft.Exchange.Rpc.dll"/>
      </dependentAssembly>
      <dependentAssembly>
        <assemblyIdentity name="Microsoft.Exchange.Diagnostics" publicKeyToken="31bf3856ad364e35" culture="neutral" />
        <codeBase version="14.0.0.0" href="file:///{0}bin\Microsoft.Exchange.Diagnostics.dll"/>
      </dependentAssembly>
    </assemblyBinding>
  </runtime>
</configuration>
"@
    # Swap in Exchange installation path
    $webConfigData = [string]::Format($webConfigOriginal, $ExchangeInstallPath)

    # Check for existing web.config
    if (Test-Path $webConfigPath)
    {
        # Make a backup copy of current web.config
        $backupPath = $webConfigPath + " Backup " + [string](get-date -Format "yyyy-MM-dd HHmmss")
        Write-Output "Backing up existing web.config to ""$backupPath"""
        Copy-Item $webConfigPath $backupPath
    }
    Out-File -FilePath $webConfigPath -InputObject $webConfigData -Encoding "UTF8"
    Write-Output "Created $webConfigPath."
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
            stop-service HTTP -force
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

# IIS: OutlookAnywhere: Enable Client certificate-to-AD authentication mapping (client cert auth already *required* on this vdir)
EnableAdClientCertAuthForPath ($RpcHttpWithCertPath)    #RpcWithCert

# Check on Schannel settings to ensure the 
FixSslDefaultFlags

# E14:475426 - Apparently we need to restart this service later in the script; earlier restart can encounter race conditions
restart-service MSExchangeFBA

# Restart of FDS will update default ACLs, and include a Deny read for the Anonymous account
Write-Output "Restarting Microsoft Exchange File Distribution Service"
restart-service MSExchangeFDS

# Get IIS config to create OAB-specific app pool, then add web.config for OABAuth module
UpdateOabWebConfig
CreateOabAppPool

Write-Output "Done!  $ComputerName configured for OutlookAnywhere with client certificate authentication."
$a=$Iis.Dispose()

# SIG # Begin signature block
# MIIanwYJKoZIhvcNAQcCoIIakDCCGowCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUbTlkfCCFJzrGyJWyikfAd+nY
# HHCgghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# a8kTo/0xggTaMIIE1gIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIHzMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBQfak9tlyDCEUs+cUI8JlBNiLshFTCBkgYKKwYBBAGCNwIBDDGB
# gzCBgKBYgFYAZQBuAGEAYgBsAGUALQBPAHUAdABsAG8AbwBrAEMAZQByAHQAaQBm
# AGkAYwBhAHQAZQBBAHUAdABoAGUAbgB0AGkAYwBhAHQAaQBvAG4ALgBwAHMAMaEk
# gCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEB
# AQUABIIBACNAHEYFWkuZIhK4UeOJUQ4Dz7q909i/jNiYcXvOmCFdGpr0GmW7GcDe
# ilrgXbcOnv0TkO8kYidJqHsQcc/3jqAhDASkv9HrTjhKDKwFL/EycCzSo2dnzeiF
# BdCjC4a1Pc5eq+T3B3oxbu1xIrQKucW95Quu1LtGYlABo3ekRI2PtJmCh7ErejzU
# iLnHInFpe64Oxf8aYNIcizUe7qoydi/Y5HDMVAyzZvKLOzSEx44GRsSnvG6/Eu0x
# sUboS+SZrKBNxErisfbO4ogU/KL35EIC4gJKzNcwukuNEeiRKALi6uGvDzfJMzaD
# VwtMdgghedy148Q/+oZXBcrlFLwAj0ShggIoMIICJAYJKoZIhvcNAQkGMYICFTCC
# AhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
# BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEh
# MB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAKzkySMGyyUjz
# AAAAAAArMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwG
# CSqGSIb3DQEJBTEPFw0xMzAyMDUwNjM3MjFaMCMGCSqGSIb3DQEJBDEWBBSAEPGM
# B0RrtHzeL64BM3ljTOl8PzANBgkqhkiG9w0BAQUFAASCAQA9wRwiiPZcHXhgu1mx
# 4KTCpQ3hdwSIyZssgbOccF1BYkJ01leM/NhVS5Rju9vfanLoSYvlyhHiHZ8sQjkk
# a9rcTDFN2A2QEgtwrlLvx7LUsDiRZZZpVOr8JNKU7fGe8ldMFr8X7h81iEMLV7V5
# ajMwdRRDp2kgRrT/lUmjLf42bb/+tZ3FaBJ2ls3WqLPRe8Xb/K4d1UIIR83PF1Yk
# iLT2LJGAuBhg1OScvrFbtaOI7AdC+6hiNO1T0stukxVVMkS5i39WQxJgPCCD5Bjo
# 0h8iHI/fhd51Ml8Qfjh7vs0keLD/XHdEfbnV5M1uDU0jryi3pe07mYZKawsCoMsp
# fQRv
# SIG # End signature block
