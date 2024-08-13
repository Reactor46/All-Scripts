##############################################################################
# Configure_IIS_WebDav_ExternalRepo_Sample.ps1
# - Configure IIS WebDav server to support OneView 3.10 External Repository.
#   Windows Server 2012 R2 or Windows Server 2016
#
#   VERSION 1.0
#
# (C) Copyright 2013-2018 Hewlett Packard Enterprise Development LP 
##############################################################################
<#
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
#>
##############################################################################

[CmdletBinding()]
param
(

    [Parameter (Mandatory = $false, HelpMessage = "The IIS website name.  Defaults to 'Default Web Site'.")]
    [ValidateNotNullorEmpty()]
    [String]$WebsiteName = 'Default Web Site',

    [Parameter (Mandatory, HelpMessage = "Specify the phyiscal path of the virtual directory.")]
    [ValidateNotNullorEmpty()]
    [String]$Path,

    [Parameter (Mandatory = $false, HelpMessage = "Specify the Virtual Directory Name.")]
    [ValidateNotNullorEmpty()]
    [String]$VirtualDirectoryName = "HPOneViewRemoteRepository",

    [Parameter (Mandatory, HelpMessage = "Specify the max size in GB for the repository.")]
    [ValidateNotNullorEmpty()]
    [Int]$Size,

    [Parameter (Mandatory = $false, HelpMessage = "Specify the max size in GB for the repository.")]
    [Switch]$RequireSSL

)

function Test-IsAdmin 
{

    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

}

if (-not(Test-IsAdmin))
{

    Write-Error -Message "Please run this script within an elevated PowerShell console." -Category AuthenticationError -ErrorAction Stop

}

$ErrorActionPreference = "Stop"
$FeatureName = 'Web-DAV-Publishing'

if (-not(Get-WindowsFeature -Name Web-Server).Installed)
{

    Write-Error -Message 'IIS is required to be installed.  Please install the Web-Server feature on this host.' -Category NotInstalled -TargetObject 'WindowsFeature:Web-Server'

}


if (-not(Get-WindowsFeature -Name $FeatureName).Installed)
{

    Stop-Service w3svc

    Write-Host 'Installing WebDAV' -ForegroundColor Cyan

    Try
    {
    
        $resp = Install-WindowsFeature -Name $FeatureName -IncludeManagementTools

        if ($resp.RestartNeeded -eq 'Yes')
        {

            Write-Warning "A reboot is needed to complete installation.  Please reboot the server, and re-run this script.  It will continue the configuration of WebDAV."

        }
    
        Write-Host 'Done.' -ForegroundColor Green

    }
    
    Catch
    {
    
        $PSCmdlet.ThrowTerminatingError($_)
    
    }

    Start-Service w3svc

}

if ((Get-WindowsFeature -Name $FeatureName).Installed -and $resp.RestartNeeded -ne 'Yes')
{

    if (-not(Get-WindowsFeature -Name Web-Dir-Browsing).Installed)
    {

        Write-Host 'Installing IIS Directory Browsing' -ForegroundColor Cyan

        Try
        {
        
            $null = Install-WindowsFeature -Name Web-Dir-Browsing -IncludeManagementTools
        
            Write-Host 'Done.' -ForegroundColor Green

        }
        
        Catch
        {
        
            $PSCmdlet.ThrowTerminatingError($_)
        
        }

    }

    Import-Module WebAdministration

    #Add Virtual Directory
    Try
    {

        if (-not(Test-Path IIS:\Sites\$WebsiteName\$VirtualDirectoryName))
        {

            $null = New-WebVirtualDirectory -Site $WebsiteName -Name $VirtualDirectoryName -PhysicalPath $Path

        }    

    }

    Catch
    {

        $PSCmdlet.ThrowTerminatingError($_)

    }

    #Check and enable Directory Browsing on the Virtual Directory
    if (-not(Get-WebConfigurationProperty -Filter /system.webServer/directoryBrowse -Location "$WebsiteName/$VirtualDirectoryName" -Name enabled).Value)
    {

        $null = Set-WebConfigurationProperty -Filter /system.webServer/directoryBrowse -Location "$WebsiteName/$VirtualDirectoryName" -Name enabled -Value $true

    }

    #Add custom HTTP Header for reposize
    Try
    {

        if (-not(Get-WebConfigurationProperty -Filter /system.webServer/httpProtocol/customHeaders -Location $WebsiteName -Name collection[name="MaxRepoSize"]))
        {

            $null = Add-WebConfigurationProperty -PSPath ('MACHINE/WEBROOT/APPHOST/{0}' -f $WebsiteName) -Filter 'system.WebServer/httpProtocol/customHeaders' -Name . -Value @{name='MaxRepoSize'; value=('{0}G' -f $Size.ToString())} -ErrorAction Stop

        }

        elseif ((Get-WebConfigurationProperty -Filter /system.webServer/httpProtocol/customHeaders -Location $WebsiteName -Name collection[name="MaxRepoSize"]).Value -ne $Size.ToString())
        {

            $null = Set-WebConfigurationProperty -PSPath ('MACHINE/WEBROOT/APPHOST/{0}' -f $WebsiteName) -Filter '/system.WebServer/httpProtocol/customHeaders' -Name . -Value @{name='MaxRepoSize'; value=('{0}G' -f $Size.ToString())} -ErrorAction Stop

        }

    }

    Catch
    {

        $PSCmdlet.ThrowTerminatingError($_)

    }

    #Add required MIME types
    Try
    {

        if (-not(Get-WebConfigurationProperty -Filter //staticContent -Location $WebsiteName -Name collection[fileExtension=".iso"]))
        {

            Add-webconfigurationproperty -Filter "//staticContent" -PSPath ("IIS:\Sites\{0}" -f $WebsiteName) -name collection -value @{fileExtension='.iso'; mimeType='application/octet-stream'} 
        
        }

        if (-not(Get-WebConfigurationProperty -Filter //staticContent -Location $WebsiteName -Name collection[fileExtension=".scexe"]))
        {

            Add-webconfigurationproperty -Filter "//staticContent" -PSPath ("IIS:\Sites\{0}" -f $WebsiteName) -name collection -value @{fileExtension='.scexe'; mimeType='application/octet-stream'} 
        
        }
        
        if ((Get-WebConfigurationProperty -Filter //staticContent -Location $WebsiteName -Name collection[fileExtension=".rpm"]).mimeType -ne "application/octet-stream")
        {

            Set-WebConfigurationProperty -Filter "//staticContent/mimeMap[@fileExtension='.rpm']" -PSPath ("IIS:\Sites\{0}" -f $WebsiteName) -Name mimeType -Value "application/octet-stream"

        }
        

    }

    Catch
    {

        $PSCmdlet.ThrowTerminatingError($_)

    }

    #Set WebDAV Access Rules
    Try
    {
        
        $NewRule = @{

            users  = "*";
            path   = "*";
            access = "Read"

        }

        if (-not(Get-WebConfigurationProperty -Filter system.webServer/webdav/authoringRules -Location $WebsiteName -Name collection[users="*"]))
        {

            $null = Add-WebConfiguration -Filter system.webServer/webdav/authoringRules -PSPath "MACHINE/WEBROOT/APPHOST" -Location $WebsiteName -Value $NewRule

        }   

        if (-not(Get-WebConfigurationProperty -filter 'system.webServer/webdav/authoring' -Location $WebsiteName -Name Enabled).Value)
        {

            [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Web.Administration")

            $IIS = new-object Microsoft.Web.Administration.ServerManager
            $WebSite = $IIS.Sites["Default Web Site"]

            $GlobalConfig = $IIS.GetApplicationHostConfiguration()
            $Config = $GlobalConfig.GetSection("system.webServer/webdav/authoring", "Default Web Site")

            if ($Config.OverrideMode -ne 'Allow')
            {

                $Config.OverrideMode = "Allow"
                $null = $IIS.CommitChanges()

            }

            Write-Host "Enabling WebDAV" -ForegroundColor Cyan
            Set-WebConfigurationProperty -filter 'system.webServer/webdav/authoring' -Location $WebsiteName -Name enabled -Value $true
            Write-Host "Done." -ForegroundColor Green

        }

        if (-not(Get-WebConfigurationProperty -filter 'system.webServer/webdav/authoring' -Location $WebsiteName -Name requireSsl).Value -and (Get-WebConfigurationProperty -filter 'system.webServer/webdav/authoring' -Location $WebsiteName -Name requireSsl).Value -ne $RequireSSL.IsPresent)
        {

            Write-Host "Enabling WebDAV SSL" -ForegroundColor Cyan
            Set-WebConfigurationProperty -filter 'system.webServer/webdav/authoring' -Location $WebsiteName -Name requireSsl -Value $RequireSSL.IsPresent
            Write-Host "Done." -ForegroundColor Green

        }

        #Enable WebDAV properties required
        Set-WebConfigurationProperty -Filter system.webServer/webdav/authoring -PSPath "MACHINE/WEBROOT/APPHOST" -Location $WebsiteName -name properties.allowAnonymousPropFind -Value $true
        Set-WebConfigurationProperty -Filter system.webServer/webdav/authoring -PSPath "MACHINE/WEBROOT/APPHOST" -Location $WebsiteName -name properties.allowInfinitePropfindDepth -Value $true


    }

    Catch
    {

        $PSCmdlet.ThrowTerminatingError($_)

    }

}
# SIG # Begin signature block
# MIIleAYJKoZIhvcNAQcCoIIlaTCCJWUCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCIv+YvF+9xVAkF
# Z+K3xsTXBqtb7zcGf2rsy3ryCcZJvaCCFhowggVMMIIDNKADAgECAhMzAAAANdjV
# WVsGcUErAAAAAAA1MA0GCSqGSIb3DQEBBQUAMH8xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKTAnBgNVBAMTIE1pY3Jvc29mdCBDb2RlIFZlcmlm
# aWNhdGlvbiBSb290MB4XDTEzMDgxNTIwMjYzMFoXDTIzMDgxNTIwMzYzMFowbzEL
# MAkGA1UEBhMCU0UxFDASBgNVBAoTC0FkZFRydXN0IEFCMSYwJAYDVQQLEx1BZGRU
# cnVzdCBFeHRlcm5hbCBUVFAgTmV0d29yazEiMCAGA1UEAxMZQWRkVHJ1c3QgRXh0
# ZXJuYWwgQ0EgUm9vdDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALf3
# GjPm8gAELTngTlvtH7xsD821+iO2zt6bETOXpClMfZOfvUq8k+0DGuOPz+VtUFrW
# lymUWoCwSXrbLpX9uMq/NzgtHj6RQa1wVsfwTz/oMp50ysiQVOnGXw94nZpAPA6s
# YapeFI+eh6FqUNzXmk6vBbOmcZSccbNQYArHE504B4YCqOmoaSYYkKtMsE8jqzpP
# hNjfzp/haW+710LXa0Tkx63ubUFfclpxCDezeWWkWaCUN/cALw3CknLa0Dhy2xSo
# RcRdKn23tNbE7qzNE0S3ySvdQwAl+mG5aWpYIxG3pzOPVnVZ9c0p10a3CitlttNC
# bxWyuHv77+ldU9U0WicCAwEAAaOB0DCBzTATBgNVHSUEDDAKBggrBgEFBQcDAzAS
# BgNVHRMBAf8ECDAGAQH/AgECMB0GA1UdDgQWBBStvZh6NLQm9/rEJlTvA73gJMtU
# GjALBgNVHQ8EBAMCAYYwHwYDVR0jBBgwFoAUYvsKIVt/Q24R2glUUGv10pZx8Z4w
# VQYDVR0fBE4wTDBKoEigRoZEaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9j
# cmwvcHJvZHVjdHMvTWljcm9zb2Z0Q29kZVZlcmlmUm9vdC5jcmwwDQYJKoZIhvcN
# AQEFBQADggIBADYrovLhMx/kk/fyaYXGZA7Jm2Mv5HA3mP2U7HvP+KFCRvntak6N
# NGk2BVV6HrutjJlClgbpJagmhL7BvxapfKpbBLf90cD0Ar4o7fV3x5v+OvbowXvT
# gqv6FE7PK8/l1bVIQLGjj4OLrSslU6umNM7yQ/dPLOndHk5atrroOxCZJAC8UP14
# 9uUjqImUk/e3QTA3Sle35kTZyd+ZBapE/HSvgmTMB8sBtgnDLuPoMqe0n0F4x6GE
# NlRi8uwVCsjq0IT48eBr9FYSX5Xg/N23dpP+KUol6QQA8bQRDsmEntsXffUepY42
# KRk6bWxGS9ercCQojQWj2dUk8vig0TyCOdSogg5pOoEJ/Abwx1kzhDaTBkGRIywi
# pacBK1C0KK7bRrBZG4azm4foSU45C20U30wDMB4fX3Su9VtZA1PsmBbg0GI1dRtI
# uH0T5XpIuHdSpAeYJTsGm3pOam9Ehk8UTyd5Jz1Qc0FMnEE+3SkMc7HH+x92DBdl
# BOvSUBCSQUns5AZ9NhVEb4m/aX35TUDBOpi2oH4x0rWuyvtT1T9Qhs1ekzttXXya
# Pz/3qSVYhN0RSQCix8ieN913jm1xi+BbgTRdVLrM9ZNHiG3n71viKOSAG0DkDyrR
# fyMVZVqsmZRDP0ZVJtbE+oiV4pGaoy0Lhd6sjOD5Z3CfcXkCMfdhoinEMIIFajCC
# BFKgAwIBAgIRAMrweR1t1bu9z9KSImtE18gwDQYJKoZIhvcNAQELBQAwfTELMAkG
# A1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UEBxMH
# U2FsZm9yZDEaMBgGA1UEChMRQ09NT0RPIENBIExpbWl0ZWQxIzAhBgNVBAMTGkNP
# TU9ETyBSU0EgQ29kZSBTaWduaW5nIENBMB4XDTE3MTEwOTAwMDAwMFoXDTE4MTEw
# OTIzNTk1OVowgdIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQRDAU5NDMwNDELMAkGA1UE
# CAwCQ0ExEjAQBgNVBAcMCVBhbG8gQWx0bzEcMBoGA1UECQwTMzAwMCBIYW5vdmVy
# IFN0cmVldDErMCkGA1UECgwiSGV3bGV0dCBQYWNrYXJkIEVudGVycHJpc2UgQ29t
# cGFueTEaMBgGA1UECwwRSFAgQ3liZXIgU2VjdXJpdHkxKzApBgNVBAMMIkhld2xl
# dHQgUGFja2FyZCBFbnRlcnByaXNlIENvbXBhbnkwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQCfY9MkbYyF0XdnYSEOHfuTNDmCvkzhhNsjbeI9I9/mkvQs
# 2MoyUPKPVNdXom5/FAmY34gOZ1/NEz2Fzbmx1TYfdNyj30iIrXMK5xhSdR3BmAvu
# plWQnlaetJXsjvBJe/DwWYzWolyedV33bFV3owX9GqfkW6b1R4xpnOESfNBh5K7J
# iXKDaK8As++Cx0+f4K77FsTWHflUeao519uIsFbFRxURQjxql0ydx8GpCCzEF6pa
# KQVx/WG7g/368P5GmqxVeH86kN4i1qGudQ+dMLwxdhm3fHNpXBnEOsdGfuWtC2ls
# pBY6LuTNP6fcXBRctJTCH5rA6F+QzhmfmXndMBKzAgMBAAGjggGNMIIBiTAfBgNV
# HSMEGDAWgBQpkWD/ik366/mmarjP+eZLvUnOEjAdBgNVHQ4EFgQUTkSybbkdnXe7
# pkRTy3t6SaDOyKswDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwEwYDVR0l
# BAwwCgYIKwYBBQUHAwMwEQYJYIZIAYb4QgEBBAQDAgQQMEYGA1UdIAQ/MD0wOwYM
# KwYBBAGyMQECAQMCMCswKQYIKwYBBQUHAgEWHWh0dHBzOi8vc2VjdXJlLmNvbW9k
# by5uZXQvQ1BTMEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwuY29tb2RvY2Eu
# Y29tL0NPTU9ET1JTQUNvZGVTaWduaW5nQ0EuY3JsMHQGCCsGAQUFBwEBBGgwZjA+
# BggrBgEFBQcwAoYyaHR0cDovL2NydC5jb21vZG9jYS5jb20vQ09NT0RPUlNBQ29k
# ZVNpZ25pbmdDQS5jcnQwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmNvbW9kb2Nh
# LmNvbTANBgkqhkiG9w0BAQsFAAOCAQEAByfbvBvqLus0/hnJSNhw5PagGRRWNDwf
# sd09noRqnkpgwbtMoP8tHOCuAbEMqrhSczgkhhk6g3yq4GSno2wnJ4qbPG4SP9zt
# HHluPmHLdEQhRQIJ9bphZiItQzbGIHFbM0W4w/U7OT2CPBZbiZ7EXTknGnkOZmJm
# uwk9MUfzgVPRRFlw0gfV10QW2dRGHCQkxtyBe9yArO3Ha6o/qEKV7GAo5dut/Su4
# NRUaUEFTkz3dcOLJN5oVjiGrhGmzgKIiEos4qxpp4Aqba8RNodgLNi3MGeVwCypm
# bzObZPJGrgAxuP1r4KNBQfG9jj/IQb6XMWm0pIy4Pd8AmFwRsl+jlTCCBXQwggRc
# oAMCAQICECdm7lbrSfOOq9dwovyE3iIwDQYJKoZIhvcNAQEMBQAwbzELMAkGA1UE
# BhMCU0UxFDASBgNVBAoTC0FkZFRydXN0IEFCMSYwJAYDVQQLEx1BZGRUcnVzdCBF
# eHRlcm5hbCBUVFAgTmV0d29yazEiMCAGA1UEAxMZQWRkVHJ1c3QgRXh0ZXJuYWwg
# Q0EgUm9vdDAeFw0wMDA1MzAxMDQ4MzhaFw0yMDA1MzAxMDQ4MzhaMIGFMQswCQYD
# VQQGEwJHQjEbMBkGA1UECBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdT
# YWxmb3JkMRowGAYDVQQKExFDT01PRE8gQ0EgTGltaXRlZDErMCkGA1UEAxMiQ09N
# T0RPIFJTQSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTCCAiIwDQYJKoZIhvcNAQEB
# BQADggIPADCCAgoCggIBAJHoVJLSClaxrA0k3cXPRGd0mSs3o30jcABxvFPfxPoq
# Eo9LfxBWvZ9wcrdhf8lLDxenPeOwBGHu/xGXx/SGPgr6Plz5k+Y0etkUa+ecs4Wg
# gnp2r3GQ1+z9DfqcbPrfsIL0FH75vsSmL09/mX+1/GdDcr0MANaJ62ss0+2PmBwU
# q37l42782KjkkiTaQ2tiuFX96sG8bLaL8w6NmuSbbGmZ+HhIMEXVreENPEVg/DKW
# USe8Z8PKLrZr6kbHxyCgsR9l3kgIuqROqfKDRjeE6+jMgUhDZ05yKptcvUwbKIpc
# Inu0q5jZ7uBRg8MJRk5tPpn6lRfafDNXQTyNUe0LtlyvLGMa31fIP7zpXcSbr0WZ
# 4qNaJLS6qVY9z2+q/0lYvvCo//S4rek3+7q49As6+ehDQh6J2ITLE/HZu+GJYLiM
# KFasFB2cCudx688O3T2plqFIvTz3r7UNIkzAEYHsVjv206LiW7eyBCJSlYCTaeiO
# TGXxkQMtcHQC6otnFSlpUgK7199QalVGv6CjKGF/cNDDoqosIapHziicBkV2v4IY
# J7TVrrTLUOZr9EyGcTDppt8WhuDY/0Dd+9BCiH+jMzouXB5BEYFjzhhxayvspoq3
# MVw6akfgw3lZ1iAar/JqmKpyvFdK0kuduxD8sExB5e0dPV4onZzMv7NR2qdH5YRT
# AgMBAAGjgfQwgfEwHwYDVR0jBBgwFoAUrb2YejS0Jvf6xCZU7wO94CTLVBowHQYD
# VR0OBBYEFLuvfgI9+qbxPISOre44mOzZMjLUMA4GA1UdDwEB/wQEAwIBhjAPBgNV
# HRMBAf8EBTADAQH/MBEGA1UdIAQKMAgwBgYEVR0gADBEBgNVHR8EPTA7MDmgN6A1
# hjNodHRwOi8vY3JsLnVzZXJ0cnVzdC5jb20vQWRkVHJ1c3RFeHRlcm5hbENBUm9v
# dC5jcmwwNQYIKwYBBQUHAQEEKTAnMCUGCCsGAQUFBzABhhlodHRwOi8vb2NzcC51
# c2VydHJ1c3QuY29tMA0GCSqGSIb3DQEBDAUAA4IBAQBkv4PxX5qF0M24oSlXDeha
# 99HpPvJ2BG7xUnC7Hjz/TQ10asyBgiXTw6AqXUz1uouhbcRUCXXH4ycOXYR5N0AT
# d/W0rBzQO6sXEtbvNBh+K+l506tXRQyvKPrQ2+VQlYi734VXaX2S2FLKc4G/HPPm
# uG5mEQWzHpQtf5GVklnxTM6jkXFMfEcMOwsZ9qGxbIY+XKrELoLL+QeWukhNkPKU
# yKlzousGeyOd3qLzTVWfemFFmBhox15AayP1eXrvjLVri7dvRvR78T1LBNiTgFla
# 4EEkHbKPFWBYR9vvbkb9FfXZX5qz29i45ECzzZc5roW7HY683Ieb0abv8TtvEDhv
# MIIF4DCCA8igAwIBAgIQLnyHzA6TSlL+lP0ct800rzANBgkqhkiG9w0BAQwFADCB
# hTELMAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4G
# A1UEBxMHU2FsZm9yZDEaMBgGA1UEChMRQ09NT0RPIENBIExpbWl0ZWQxKzApBgNV
# BAMTIkNPTU9ETyBSU0EgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkwHhcNMTMwNTA5
# MDAwMDAwWhcNMjgwNTA4MjM1OTU5WjB9MQswCQYDVQQGEwJHQjEbMBkGA1UECBMS
# R3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdTYWxmb3JkMRowGAYDVQQKExFD
# T01PRE8gQ0EgTGltaXRlZDEjMCEGA1UEAxMaQ09NT0RPIFJTQSBDb2RlIFNpZ25p
# bmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCmmJBjd5E0f4rR
# 3elnMRHrzB79MR2zuWJXP5O8W+OfHiQyESdrvFGRp8+eniWzX4GoGA8dHiAwDvth
# e4YJs+P9omidHCydv3Lj5HWg5TUjjsmK7hoMZMfYQqF7tVIDSzqwjiNLS2PgIpQ3
# e9V5kAoUGFEs5v7BEvAcP2FhCoyi3PbDMKrNKBh1SMF5WgjNu4xVjPfUdpA6M0ZQ
# c5hc9IVKaw+A3V7Wvf2pL8Al9fl4141fEMJEVTyQPDFGy3CuB6kK46/BAW+QGiPi
# XzjbxghdR7ODQfAuADcUuRKqeZJSzYcPe9hiKaR+ML0btYxytEjy4+gh+V5MYnmL
# Agaff9ULAgMBAAGjggFRMIIBTTAfBgNVHSMEGDAWgBS7r34CPfqm8TyEjq3uOJjs
# 2TIy1DAdBgNVHQ4EFgQUKZFg/4pN+uv5pmq4z/nmS71JzhIwDgYDVR0PAQH/BAQD
# AgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwEQYD
# VR0gBAowCDAGBgRVHSAAMEwGA1UdHwRFMEMwQaA/oD2GO2h0dHA6Ly9jcmwuY29t
# b2RvY2EuY29tL0NPTU9ET1JTQUNlcnRpZmljYXRpb25BdXRob3JpdHkuY3JsMHEG
# CCsGAQUFBwEBBGUwYzA7BggrBgEFBQcwAoYvaHR0cDovL2NydC5jb21vZG9jYS5j
# b20vQ09NT0RPUlNBQWRkVHJ1c3RDQS5jcnQwJAYIKwYBBQUHMAGGGGh0dHA6Ly9v
# Y3NwLmNvbW9kb2NhLmNvbTANBgkqhkiG9w0BAQwFAAOCAgEAAj8COcPu+Mo7id4M
# bU2x8U6ST6/COCwEzMVjEasJY6+rotcCP8xvGcM91hoIlP8l2KmIpysQGuCbsQci
# GlEcOtTh6Qm/5iR0rx57FjFuI+9UUS1SAuJ1CAVM8bdR4VEAxof2bO4QRHZXavHf
# WGshqknUfDdOvf+2dVRAGDZXZxHNTwLk/vPa/HUX2+y392UJI0kfQ1eD6n4gd2HI
# TfK7ZU2o94VFB696aSdlkClAi997OlE5jKgfcHmtbUIgos8MbAOMTM1zB5TnWo46
# BLqioXwfy2M6FafUFRunUkcyqfS/ZEfRqh9TTjIwc8Jvt3iCnVz/RrtrIh2IC/gb
# qjSm/Iz13X9ljIwxVzHQNuxHoc/Li6jvHBhYxQZ3ykubUa9MCEp6j+KjUuKOjswm
# 5LLY5TjCqO3GgZw1a6lYYUoKl7RLQrZVnb6Z53BtWfhtKgx/GWBfDJqIbDCsUgmQ
# Fhv/K53b0CDKieoofjKOGd97SDMe12X4rsn4gxSTdn1k0I7OvjV9/3IxTZ+evR5s
# L6iPDAZQ+4wns3bJ9ObXwzTijIchhmH+v1V04SF3AwpobLvkyanmz1kl63zsRQ55
# ZmjoIs2475iFTZYRPAmK0H+8KCgT+2rKVI2SXM3CZZgGns5IW9S1N5NGQXwH3c/6
# Q++6Z2H/fUnguzB9XIDj5hY5S6cxgg60MIIOsAIBATCBkjB9MQswCQYDVQQGEwJH
# QjEbMBkGA1UECBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdTYWxmb3Jk
# MRowGAYDVQQKExFDT01PRE8gQ0EgTGltaXRlZDEjMCEGA1UEAxMaQ09NT0RPIFJT
# QSBDb2RlIFNpZ25pbmcgQ0ECEQDK8HkdbdW7vc/SkiJrRNfIMA0GCWCGSAFlAwQC
# AQUAoHwwEAYKKwYBBAGCNwIBDDECMAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
# AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIE
# IHzfQ0t/SOO0CPisEV7SmVpVZID17aF3qFk4pjjJcJb9MA0GCSqGSIb3DQEBAQUA
# BIIBAGSA6/9M1T49rQdeAXzs7elEnLee9usVm17r1SiR9dfce2vJtMMA5ZmdgWNF
# +X4PFbj+tWAjfbU2vZSzj1xhzRQ6Ck4bo2vv7GtN/AepsCRcB+/416B3k4LKl4dh
# WyOInYDItmhVjp4423u9pz8Q9xoXqjVm8Ke0jVvmt4Zn1EJhrozNhPB6FSy5ywxG
# edum7MbcdOtV6w4kmaEMmQQLmYaJFpXAsH2ZwTTU7ygVWMq4ZYoO5hC4qHbqk5LW
# vGJXCDltIoa/682pCbh0PB7zzyGRKdDKGhzmKuvRnLGIjy/WY4YMgsa7I2C4nmMp
# OynrrUjokgv+3NhgAnBXtZ76YAChggx0MIIMcAYKKwYBBAGCNwMDATGCDGAwggxc
# BgkqhkiG9w0BBwKgggxNMIIMSQIBAzEPMA0GCWCGSAFlAwQCAQUAMIGvBgsqhkiG
# 9w0BCRABBKCBnwSBnDCBmQIBAQYJKwYBBAGgMgIDMDEwDQYJYIZIAWUDBAIBBQAE
# IHicvXKTKlmRxBL68cEg+rXhMLXRuv/htLMkvlGUANh9AhRT0TgHFAcocKm63GUR
# czrlRYjeCBgPMjAxODA1MDMwMzIyNDhaoC+kLTArMSkwJwYDVQQDDCBHbG9iYWxT
# aWduIFRTQSBmb3IgQWR2YW5jZWQgLSBHMqCCCNMwggS2MIIDnqADAgECAgwMp89d
# BwckrInnmjowDQYJKoZIhvcNAQELBQAwWzELMAkGA1UEBhMCQkUxGTAXBgNVBAoT
# EEdsb2JhbFNpZ24gbnYtc2ExMTAvBgNVBAMTKEdsb2JhbFNpZ24gVGltZXN0YW1w
# aW5nIENBIC0gU0hBMjU2IC0gRzIwHhcNMTgwMjE5MDAwMDAwWhcNMjkwMzE4MTAw
# MDAwWjArMSkwJwYDVQQDDCBHbG9iYWxTaWduIFRTQSBmb3IgQWR2YW5jZWQgLSBH
# MjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALfHkooo2PORy1ANXesp
# RMGCWaXKZM69g7VR5ZTMboCaF2zc/2LmNkNeAcIMZI3Kd572XXdFuV7IJOtBNxFm
# N6zIzXSbzLPvTOJ/G85zvsmWnTUefPdU92zsoBLWrpmdY8R4X1mpLiL1wyfYsltF
# YyeQ/4yxPam08w7A8SBlBomdAxyjsFJBhTTrvMvOVPYS/rMBiUqm+lTFH/vTHMDj
# v5fjP9Ab+UDHG9XrJnxDMMdw8ngRqoVOpQ4NAEo6EXejyiMBgJ7Ik1ZdRsyK2NKq
# CoSFsolb1TLOQXsYTlTKq9FSXhLTJJ5W8wyP3b2SjnnVQYnDo6DlkfzHZ52HM85x
# MnMCAwEAAaOCAagwggGkMA4GA1UdDwEB/wQEAwIHgDBMBgNVHSAERTBDMEEGCSsG
# AQQBoDIBHjA0MDIGCCsGAQUFBwIBFiZodHRwczovL3d3dy5nbG9iYWxzaWduLmNv
# bS9yZXBvc2l0b3J5LzAJBgNVHRMEAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMI
# MEYGA1UdHwQ/MD0wO6A5oDeGNWh0dHA6Ly9jcmwuZ2xvYmFsc2lnbi5jb20vZ3Mv
# Z3N0aW1lc3RhbXBpbmdzaGEyZzIuY3JsMIGYBggrBgEFBQcBAQSBizCBiDBIBggr
# BgEFBQcwAoY8aHR0cDovL3NlY3VyZS5nbG9iYWxzaWduLmNvbS9jYWNlcnQvZ3N0
# aW1lc3RhbXBpbmdzaGEyZzIuY3J0MDwGCCsGAQUFBzABhjBodHRwOi8vb2NzcDIu
# Z2xvYmFsc2lnbi5jb20vZ3N0aW1lc3RhbXBpbmdzaGEyZzIwHQYDVR0OBBYEFC1u
# btGN5QOA7udj6afZ2gs8VyI9MB8GA1UdIwQYMBaAFJIhp0qVXWSwm7Qe5gA3R+ad
# QStMMA0GCSqGSIb3DQEBCwUAA4IBAQCN/R0fj4jTZfi1XEbpp9O2P0rLQwXCBw3D
# XgOEIzeoeoyXI/8nnIR9YetcS1c+m/RE24BY97Zl3o8SCl+nh3S32hRAh1Hsn0pH
# VfAX6Bg+cSAx5igiq072rpxrwudE+2gNIH9TP5QboYaCTjIWEFo/iVjB0LTbtpD7
# jTLpoUtGUDe2wzrpKSQSklpOK7m2Cno2jJfA74l8M94x4aGX5OdFrWFS/VqJ01R+
# ik1biXsOlnfR8jw4G9r5pWEgZQK8xa3XDGIQKCzdtuOoYQCsCCajjYXQZMk2kB6K
# UsmHjTWlXUiR/JOmpWXx7Pk63KUsngRtYjP+3vqfxc6vlfNRqPLOMIIEFTCCAv2g
# AwIBAgILBAAAAAABMYnGUAQwDQYJKoZIhvcNAQELBQAwTDEgMB4GA1UECxMXR2xv
# YmFsU2lnbiBSb290IENBIC0gUjMxEzARBgNVBAoTCkdsb2JhbFNpZ24xEzARBgNV
# BAMTCkdsb2JhbFNpZ24wHhcNMTEwODAyMTAwMDAwWhcNMjkwMzI5MTAwMDAwWjBb
# MQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2lnbiBudi1zYTExMC8GA1UE
# AxMoR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0EgLSBTSEEyNTYgLSBHMjCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAKqbjsOrEVElAbaWlOJP2MEI9kYj
# 2UXFlZdbqxq/0mxXyTMGH6APxjx+U0h6v52Hnq/uw4xH4ULs4+OhSmwMF8SmwbnN
# W/EeRImO/gveIVgT7k3IxWcLHLKz8TR2kaLLB203xaBHJgIVpJCRqXme1+tXnSt8
# ItgU1/EHHngiNmt3ea+v+X+OTuG1CDH96u1LcWKMI/EDOY9EebZ2A1eerS8IRtzS
# jLz0jnTOyGhpUXYRiw9dJFsZVD0mzECNgicbWSB9WfaTgI74Kjj9a6BAZR9Xdsxb
# jgRPLKjbhFATT8bci7n43WlMiOucezAm/HpYu1m8FHKSgVe3dsnYgAqAbgkCAwEA
# AaOB6DCB5TAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNV
# HQ4EFgQUkiGnSpVdZLCbtB7mADdH5p1BK0wwRwYDVR0gBEAwPjA8BgRVHSAAMDQw
# MgYIKwYBBQUHAgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29tL3JlcG9zaXRv
# cnkvMDYGA1UdHwQvMC0wK6ApoCeGJWh0dHA6Ly9jcmwuZ2xvYmFsc2lnbi5uZXQv
# cm9vdC1yMy5jcmwwHwYDVR0jBBgwFoAUj/BLf6guRSSuTVD6Y5qL3uLdG7wwDQYJ
# KoZIhvcNAQELBQADggEBAARWgkp80M7JvzZm0b41npNsl+gGzjEYWflsQV+ALsBC
# JbgYx/zUsTfEaKDPKGoDdEtjl4V3YTvXL+P1vTOikn0RH56KbO8ssPRijTZz0RY2
# 8bxe7LSAmHj80nZ56OEhlOAfxKLhqmfbs5xz5UAizznO2+Z3lae7ssv2GYadn8jU
# mAWycW9Oda7xPWRqO15ORqYqXQiS8aPzHXS/Yg0jjFwqOJXSwNXNz4jaHyi1uoFp
# ZCq1pqLVc6/cRtsErpHXbsWYutRHxFZ0gEd4WIy+7yv97Gy/0ZT3v1Dge+CQ/SAY
# eBgiXQgujBygl/MdmX2jnZHTBkROBG56HCDjNvC2ULkxggKoMIICpAIBATBrMFsx
# CzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMTEwLwYDVQQD
# EyhHbG9iYWxTaWduIFRpbWVzdGFtcGluZyBDQSAtIFNIQTI1NiAtIEcyAgwMp89d
# BwckrInnmjowDQYJYIZIAWUDBAIBBQCgggEOMBoGCSqGSIb3DQEJAzENBgsqhkiG
# 9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcNMTgwNTAzMDMyMjQ4WjAvBgkqhkiG9w0B
# CQQxIgQg4nyyykiafgrHaXTX1IoI407qp9NcdmH1WVxaN9/UKIQwgaAGCyqGSIb3
# DQEJEAIMMYGQMIGNMIGKMIGHBBSbEgV65yqv9tY3crSfaiNvJknNqTBvMF+kXTBb
# MQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2lnbiBudi1zYTExMC8GA1UE
# AxMoR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0EgLSBTSEEyNTYgLSBHMgIMDKfP
# XQcHJKyJ55o6MA0GCSqGSIb3DQEBAQUABIIBAJTsqrj4hZYUey7EcLv0hBWFInhH
# Ppx982pCoQW7yXJAdvq+0y+kEuhsNNBWZ6ZdtWt7r6kOX1kdxOM9LLBSZ9JoJxZc
# fIX9vDCmRRfIfkw89lSuNaee3Ut5gadY0sWFX1kI0QHQOSkXMOiRoyjKystxAGZg
# DzqRqK4BjNrCAxUZ7Opm4uYdMNsft7NqYi0CH+tFjQH4D5eHa6jgWoiD5ZBoSzX1
# EfWdauUfxf8Kj0a1YAPoQe4PXUNKEUUHWMwjeDZE3noO6yCdp/MUBiGrJpQlfoCH
# 6EPKfv34afqDEAPsinhhINh8P+po5GZR4cDecOh/uGzcC8fECFr5RfTSos4=
# SIG # End signature block
