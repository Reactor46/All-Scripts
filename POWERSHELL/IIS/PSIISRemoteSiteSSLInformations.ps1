Import-Module -Name CertificateHealth

$servers = Get-Content C:\LazyWinAdmin\IIS\IIS-Reports\Web.txt 
#Enumerate the server list from the text file
foreach ($server in $servers) {

Get-CertificateHealth -ComputerName $server | ForEach-Object -Process `
        {
            If ($_.Days -gt 0) 
            {                            
                    [PsCustomObject]@{
                        Server  = $_.ComputerName
                        Issuer  = $_.Subject 
                        ExpirationDate = $_.NotAfter
                        DaysTillExpired = $_.Days
                        Thumbprint = $_.Thumbprint

                        }
            }
        } | Select Server, Issuer, ExpirationDate, DaysTillExpired, Thumbprint
    } 