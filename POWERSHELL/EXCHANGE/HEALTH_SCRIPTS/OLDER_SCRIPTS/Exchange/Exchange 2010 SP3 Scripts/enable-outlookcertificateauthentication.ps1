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
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUbTlkfCCFJzrGyJWyikfAd+nY
# HHCgghhqMIIE2jCCA8KgAwIBAgITMwAAASIn72vt4vugowAAAAABIjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTgxMDI0MjEwNzQw
# WhcNMjAwMTEwMjEwNzQwWjCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046RUFDRS1FMzE2LUM5MUQxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQDRU0uHr30XLzP0oZTW7fCdslb6OXTQeoCa/IpqzTgDcXyf
# EqI0fdvhtqQ84neZE4vwJUAbJ2S+ajJirzzZIEU/JiTZpJgeeAMtN+MuAbzXrdyS
# ohyUDuGkuN+vVSeCnEZkeGcFf/zrNWWXmS7JsVK2BJR8YvXk0sBUbWVpdj0uvz68
# Y+HUyx8AKKE2nHRu54f6fC4eiwP/hs+L7NejJm+sNo7HXV4Y6edQI36FdY0Sotq8
# 7Lh3U96U4O6X9cD0iqKxr4lxYYkh98AzVUjiiSdWUt65DAMbdjBV6cepatwVVoET
# EtNK/f83bMS3sOL00QMWoyQM1F7+fLoz1TF7qlozAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUtlFVlkkUKuXnuF3JZxfDlHs2paYwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAZldsd7vjji5U30Zj
# cKiJvhDtcmx0b4s4s0E7gd8Lp4VnvAQAnpc3SkknUslyknvHGE77OSdxKdrO8qnn
# T0Tymqvf7/Re2xJcRVcM4f8TeE5hCaffCkB7Gtu90R+6+Eb1BnBDYMbj3b42Jq8K
# 42hnDG0ntrgv4/TmyJWIvmGQORWMCWyM/NraY3Ldi7pDpTfx9Z9s4eNE/cxipoST
# XHMIgPgDgbZcuFBANnWwF+/swj69cv87x+Jv/8HM/Naoawrr8+0yDjiJ90OzLGI5
# RScuGfQUlH0ESbzevO/9PFpoUywmNYhHoEPngLJVT2W6y13jFUx3IS9lnR0r1dCh
# mynB8jCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
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
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUH2pPbZcgwhFLPnFCPCZQTYi7IRUw
# XAYKKwYBBAGCNwIBDDFOMEygJIAiAEMASQBUAFMAQwBvAG4AcwB0AGEAbgB0AHMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAF6UJx5dzRelMjBygQ1+AZOE7i3RDYIiy8apngfPPuxV
# nn8qseA2vpEs9vxnwQzcSVcnhyLJ+tOZUJCLrOf5HLlYdHasjd4P4pQTJLaSFVLT
# 45+M+MldsFCBA1o7MeiNUJGQOS2KDcEPkdNmQSLkkNjYLYBVCPq//rjSnOn1gdn9
# X5cJX/ZKBi79bB+k0lrZcylyZqsJX8vRs354lPQoo6hpgEqnl6vfl1CTd+DVrFde
# 863O5wI22j100unMyxOPpfBUoMy8OiOvOhk+vTlmSBPSlC7QT6LoInYcK1Cdguq7
# j+IlL8uXmU5jMswoJOEyDZY0Xs63O47o8y46EkSU4UShggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAB
# Iifva+3i+6CjAAAAAAEiMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xOTA2MjYxNTI1NTlaMCMGCSqGSIb3DQEJ
# BDEWBBRicpkYl64T/zJH2fPlzwj8LF9eVjANBgkqhkiG9w0BAQUFAASCAQBWnHrX
# 2mQE8uhZCfSTsklWAfhCqyKkF0C+cqwI1sNCEyiF8/3UCRXSj471aGuV5fbWF+aL
# 0mskikEwcPtzeIY6txvVzRP9WArU6PWGcBYypvs6B1tUPlkntXcwNAxOv8yUIlx5
# 2KGrBGVXx9HFBlTDqlOzkpI8fOX4F/pFwOd7M3G8hRjFydIfCifozyg6nxQQmITP
# wQhnj3MwpI4VRuSw8eiInkgyJXWc0HF9FYoGeBoLm2afUp10wR/fIeIc4J94AuX5
# 9ghN+XuTehMfpDJx8b/LEx2Tp0oWY5K0FFyZl+11xNYqZtYV8rZVrV1nVtY3DrXx
# OXTR7xZVk0qjcKRi
# SIG # End signature block
