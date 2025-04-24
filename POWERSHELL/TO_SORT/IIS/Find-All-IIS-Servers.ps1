$Servers = ""
$Servers = (Get-Content D:\AllServers.txt)

foreach($srv in $Servers){
(Get-Service -ComputerName $srv -Name W3SVC -ErrorAction SilentlyContinue | Where Status -eq "Running" -ErrorAction SilentlyContinue).MachineName |
        Out-FileUtf8NoBom D:\All-IIS-Servers.txt -Append
        }



# Define a list of servers or get the list dynamically
$servers = Get-Content D:\All-IIS-Servers.txt  # Add your servers here

# Define the function to get IIS SSL certificates
function Get-IISCertificates {
    param (
        [string]$server
    )

    $certificates = @()

    Invoke-Command -ComputerName $server -ScriptBlock {
        Import-Module WebAdministration

        $bindings = Get-WebBinding -ErrorAction SilentlyContinue | Where-Object { $_.protocol -eq 'https' }

        foreach ($binding in $bindings) {
            $certThumbprint = $binding.certificateHash
            $cert = Get-ChildItem -Path cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $certThumbprint }
            
            [PSCustomObject]@{
                ServerName = $env:COMPUTERNAME
                SiteName   = $binding.ItemXPath.Split('/')[-2]
                IP         = $binding.bindingInformation.Split(':')[0]
                Port       = $binding.bindingInformation.Split(':')[1]
                Hostname   = $binding.bindingInformation.Split(':')[2]
                Thumbprint = $certThumbprint
                Subject    = $cert.Subject
                Issuer     = $cert.Issuer
                ExpiryDate = $cert.NotAfter
            }
        }
    }
}

# Loop through all servers and get their IIS certificates
$allCertificates = foreach ($server in $servers) {
    Get-IISCertificates -server $server
}

# Export the results to a CSV file
$csvPath = "D:\IIS_SSL_Certificates.csv"
$allCertificates | Export-Csv -Path $csvPath -NoTypeInformation

# Export the results to an HTML file
$htmlPath = "D:\IIS_SSL_Certificates.html"
$allCertificates | ConvertTo-Html -Property ServerName, SiteName, IP, Port, Hostname, Thumbprint, Subject, Issuer, ExpiryDate -Head "<style>table { font-family: Arial; border-collapse: collapse; width: 100%; } th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; } th { background-color: #f2f2f2; }</style>" -Title "IIS SSL Certificates Report" | Out-File -FilePath $htmlPath

Write-Host "SSL Certificate report generated:"
Write-Host "CSV: $csvPath"
Write-Host "HTML: $htmlPath"
