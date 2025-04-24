﻿<#
.SYNOPSIS
    Retrieve SSL Certificate data.

.DESCRIPTION
    Retrieve SSL Certificate data.

.PARAMETER Servers
    Execute command on specified server(s).

.NOTES
    Written by: JBear
    Date: 9/7/2017
#>

param(
    
    [Parameter(Mandatory=$true)]
    [String[]]$Servers
)

function Check-SSLCertificates {

    foreach($Server in $Servers) {

        $RO = [System.Security.Cryptography.X509Certificates.OpenFlags]"ReadOnly"
        $LM = [System.Security.Cryptography.X509Certificates.StoreLocation]"LocalMachine"

        $Stores = New-Object System.Security.Cryptography.X509Certificates.X509Store("\\$Server\root",$LM)
        $Stores.Open($RO)
        $Certs = $Stores.Certificates
        
        foreach($Cert in $Certs) {

            [PSCustomObject] @{
        
                Server=$Server
                FriendlyName=$Cert.FriendlyName
                DNS=$Cert.DNSNameList
                ExpirationDate=$Cert.NotAfter
                Version=$Cert.Version
                HasPrivateKey=$Cert.HasPrivateKey
            }
        }
    }
}

#Call main function
Check-SSLCertificates