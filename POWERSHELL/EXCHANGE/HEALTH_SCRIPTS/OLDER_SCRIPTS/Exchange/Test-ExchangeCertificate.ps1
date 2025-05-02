function Test-ExchangeCertificate 
    {
        Param(
          [Parameter(Mandatory = $false)]
          [array] $Servers
        )

        #Get a list of Exchange 2010 - 2016 CAS servers if no servers specified
        if(!$servers)
            {
                $servers = Get-ExchangeServer | ? {$_.ServerRole -match "ClientAccess"}
            }
        if($servers)
            {
                $a = @()
                $servers | % {$a += Get-ExchangeServer $_}
                $servers = $a
            }


        #Create empty array 
        $CSV = @()

        #For each server
        foreach($server in $servers)
            {
                #Exchange 2013 SP1 and higher and MAPI over HTTP enabled
                if((Get-OrganizationConfig).MapiHttpEnabled -eq $true `
                -and ((($server.AdminDisplayVersion.Major -eq "15" -and $server.AdminDisplayVersion.Minor -eq 0 -and $server.AdminDisplayVersion.Build -ge 847) `
                -or ($server.AdminDisplayVersion.Major -eq "15" -and $server.AdminDisplayVersion.Minor -ge 1))))
                    {
                        $Mapi = Get-MapiVirtualDirectory -Server $server -ADPropertiesOnly
                        $MapiInternalUrl = $Mapi.InternalUrl
                        $MapiExternalUrl = $Mapi.ExternalUrl
                    }
                else
                    {
                        $MapiInternalUrl = "Not Applicable"
                        $MapiExternalUrl = "Not Applicable"
                    }

                #Get the Outlook Anywhere hostnames 
                $OutlookAnywhere = Get-OutlookAnywhere -Server $server -ADPropertiesOnly
                $OutlookAnywhereExternalHostname = $OutlookAnywhere.ExternalHostname

                if($server.AdminDisplayVersion.Major -ge 15)
                    {
                        $OutlookAnywhereInternalHostname = $OutlookAnywhere.InternalHostname
                        
                    }
                else
                    {
                        $OutlookAnywhereInternalHostname = "Not Applicable"
                    }
                
                #Get the internal and external URL for all virtual directories common in Exchange 2010 - 2016
                $OWA = Get-OwaVirtualDirectory -Server $server -ADPropertiesOnly
                $OWAInternalUrl = $OWA.InternalUrl
                $OWAExternalUrl = $OWA.ExternalUrl
                $ECP = Get-EcpVirtualDirectory -Server $server -ADPropertiesOnly
                $EcpInternalUrl = $ECP.InternalUrl
                $EcpExternalUrl = $ECP.ExternalUrl
                $EWS = Get-WebServicesVirtualDirectory -Server $server -ADPropertiesOnly
                $WebServicesInternalUrl = $EWS.InternalUrl
                $WebServicesExternalUrl = $EWS.ExternalUrl
                $OAB = Get-OABVirtualDirectory -Server $server -ADPropertiesOnly
                $OABInternalUrl = $OAB.InternalUrl
                $OABExternalUrl = $OAB.ExternalUrl
                $ActiveSync = Get-ActiveSyncVirtualDirectory -Server $server -ADPropertiesOnly
                $ActiveSyncInternalUrl = $ActiveSync.InternalUrl
                $ActiveSyncExternalUrl = $ActiveSync.ExternalUrl
                $AutodiscoverServiceInternalUri = (Get-ClientAccessServer -Identity $server.Name -WarningAction SilentlyContinue).AutodiscoverServiceInternalUri.AbsoluteUri

                #Get the certificate thumbprint
                $IISCertificate = Get-ExchangeCertificate -Server $server | ? {$_.Services -match "IIS"}
                if($IISCertificate.Count -gt 1)
                    {
                        $CertificateThumbprint = "MultipleIISCertError"
                        $CertificateNamesMissing = "MultipleIISCertError"
                        $CertificateNameMismatch = "MultipleIISCertError"
                    }
                else
                    {
                        $CertificateThumbprint = $IISCertificate.ThumbPrint
                        $CertificateSelfSigned = $IISCertificate.IsSelfSigned
                        $CertificateExpiry = $IISCertificate.NotAfter
                        
                        #Check that all the FQDNs on the virtual directories and Outlook Anywhere are included on the certificate and provide an error if not
                        $URLs = $OWAInternalUrl,$OWAExternalUrl,$EcpInternalUrl,$EcpExternalUrl,$WebServicesInternalUrl,$WebServicesExternalUrl, `
                        $OABInternalUrl,$OABExternalUrl,$ActiveSyncInternalUrl,$ActiveSyncExternalUrl,$AutodiscoverServiceInternalUri, `
                        $MapiInternalUrl,$MapiExternalUrl

                        $HostNames = $OutlookAnywhereExternalHostname,$OutlookAnywhereInternalHostname
                        
                        $FQDNs = @()

                        $urls | % {$FQDNs += $_.Host}
                        
                        $HostNames | % {$FQDNs += $_} 

                        $FQDNs = $FQDNs | select -Unique | ? {$_ -ne "Not Applicable"}

                        $CertificateNames = @()
                        $IISCertificate.CertificateDomains | % {$CertificateNames += $_.Address}

                        $CertificateNamesMissing = @()

                        #Identify whether certificate is a wildcard certificate
                        if($CertificateNames -match "\*")
                            {
                                #Certificate is a wildcard certificate
                                $CertificateNames = $CertificateNames | ? {$_ -match "\*"}
                                $WildcardDomain = $CertificateNames -replace "\*.",""
                                $CertificateNamesMissing = $FQDNs | ? {$_ -notmatch $WildcardDomain}
                            }
                        else
                            {
                                #Certificate is a SAN certificate
                                foreach($FQDN in $FQDNs)
                                    {
                                        if($CertificateNames -notcontains $FQDN)
                                            {
                                                $CertificateNamesMissing += $FQDN
                                            }
                                    }
                            }

                        if($CertificateNamesMissing.Count -gt 0)
                            {
                                $CertificateNameMismatch = $true
                            }
                        else
                            {
                                $CertificateNameMismatch = $false
                            }
                            

                    }

                $CSVLine = New-Object System.Object
                $CSVLine | Add-Member -Type NoteProperty -Name Server -Value $server.Name
                $CSVLine | Add-Member -Type NoteProperty -Name Site -Value $server.Site.Name
                $CSVLine | Add-Member -Type NoteProperty -Name CertificateThumbprint -Value $CertificateThumbprint
                $CSVLine | Add-Member -Type NoteProperty -Name CertificateNames -Value $CertificateNames
                $CSVLine | Add-Member -Type NoteProperty -Name CertificateNamesMissing -Value $CertificateNamesMissing                
                $CSVLine | Add-Member -Type NoteProperty -Name CertificateNameMismatch -Value $CertificateNameMismatch
                $CSVLine | Add-Member -Type NoteProperty -Name CertificateSelfSigned -Value $CertificateSelfSigned 
                $CSVLine | Add-Member -Type NoteProperty -Name CertificateExpiry -Value $CertificateExpiry
                $CSVLine | Add-Member -Type NoteProperty -Name OWAInternalUrl -Value $OWAInternalUrl
                $CSVLine | Add-Member -Type NoteProperty -Name OWAExternalUrl -Value $OWAExternalUrl
                $CSVLine | Add-Member -Type NoteProperty -Name EcpInternalUrl -Value $EcpInternalUrl
                $CSVLine | Add-Member -Type NoteProperty -Name EcpExternalUrl -Value $EcpExternalUrl
                $CSVLine | Add-Member -Type NoteProperty -Name WebServicesInternalUrl -Value $WebServicesInternalUrl
                $CSVLine | Add-Member -Type NoteProperty -Name WebServicesExternalUrl -Value $WebServicesExternalUrl
                $CSVLine | Add-Member -Type NoteProperty -Name OABInternalUrl -Value $OABInternalUrl
                $CSVLine | Add-Member -Type NoteProperty -Name OABExternalUrl -Value $OABExternalUrl
                $CSVLine | Add-Member -Type NoteProperty -Name ActiveSyncInternalUrl -Value $ActiveSyncInternalUrl
                $CSVLine | Add-Member -Type NoteProperty -Name ActiveSyncExternalUrl -Value $ActiveSyncExternalUrl
                $CSVLine | Add-Member -Type NoteProperty -Name AutodiscoverServiceInternalUri -Value $AutodiscoverServiceInternalUri
                $CSVLine | Add-Member -Type NoteProperty -Name OutlookAnywhereInternalHostname -Value $OutlookAnywhereInternalHostname
                $CSVLine | Add-Member -Type NoteProperty -Name OutlookAnywhereExternalHostname -Value $OutlookAnywhereExternalHostname
                $CSVLine | Add-Member -Type NoteProperty -Name MapiInternalUrl -Value $MapiInternalUrl
                $CSVLine | Add-Member -Type NoteProperty -Name MapiExternalUrl -Value $MapiExternalUrl
                $CSV += $CSVLine

            }
        $CSV
    }



