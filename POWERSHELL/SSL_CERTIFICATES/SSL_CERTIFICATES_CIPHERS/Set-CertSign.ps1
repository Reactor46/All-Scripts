################
# Set-CertSign #
################
#Version: 2015-06-08
#Author: johan.carlsson@innovatum.se , http://www.innovatum.se/personer/johan-carlsson

<#

.SYNOPSIS
Use this script to sign PowerShell scripts.

.DESCRIPTION
The script will check if current user has a Code Signing certificate, if not it requests one from the local Certificate Authority, and then signs the assigned script.

.EXAMPLE
.\Set-CertSign.ps1 -FilePath C:\Temp\PSScriptToSign.ps1
The script will check if current user has a Code Signing certificate, if not it requests one from the local Certificate Authority, and then signs the assigned script.

.NOTES
A working domain Certificate Authority is required.

.LINK
https://gallery.technet.microsoft.com/scriptcenter/Set-CertSign-1399afdb

#>

[CmdletBinding()]
Param
(
    [Parameter(Mandatory=$True,Position=1,HelpMessage="Full path to PowerShell script to sign.")]
    [string]$FilePath
)

#Path to User Certificate Store
$CertStoreLocation = "Cert:\CurrentUser\My"
#URL to TimestampServer to use when signing script
$TimestampServer = "http://timestamp.comodoca.com/authenticode"

#Check if current user has a code signing certificate
Write-Host "Validating Code Signing Certificate for user: $env:USERNAME... " -NoNewline -ForegroundColor Yellow
If ((Get-ChildItem $CertStoreLocation -CodeSigningCert) -eq $null)
    {
        #Request certificate
        Write-Host "Missing" -BackgroundColor Red -ForegroundColor White
        Write-Host "Requesting Code Signing Certificate from CA for user: $env:USERNAME... " -NoNewline -ForegroundColor Yellow
        Get-Certificate -Template "CodeSigning" -Url ldap: -CertStoreLocation $CertStoreLocation | Out-Null
        If ($? -eq $true)
            {
                Write-Host "OK" -BackgroundColor Green -ForegroundColor White
            }
        Else
            {
                Write-Error -Message "Failed to obtain Code Signing Certificate from Certificate Authority" -Category ResourceUnavailable -ErrorAction Stop
            }
    }
Else
    {
        Write-Host "OK" -BackgroundColor Green -ForegroundColor White
    }

#Check if current user has a code signing certificate
If ((Get-ChildItem $CertStoreLocation -CodeSigningCert) -ne $null)
    {
        #Sign script
        Write-Host "Signing: $FilePath... " -NoNewline -ForegroundColor Yellow
        $Certificate=(Get-ChildItem $CertStoreLocation -CodeSigningCert)
        Set-AuthenticodeSignature -FilePath $FilePath -Certificate $Certificate -TimestampServer $TimestampServer | Out-Null
        If ($? -eq $true)
            {
                Write-Host "OK" -BackgroundColor Green -ForegroundColor White
            }
        Else
            {
                Write-Error -Message "Failed to sign script: $FilePath" -Category ResourceUnavailable -ErrorAction Stop
            }
    }
Else
    {
        Write-Error -Message "Failed to obtain Code Signing Certificate from Certificate Authority" -Category ResourceUnavailable -ErrorAction Stop
    }