# List of domain names (you can load them from a text file)
$Domains = @('kelsey-seybold.com','kelseycareadvantage.com','kelseycare.com')


#$Domains = Get-Content -Path "D:\RewriteMaps\Combined-Domain-List.txt"
#& "D:\RewriteMaps\CertificateScanner.ps1"
# Function to check if a port is open (e.g., 80 or 443)
function Check-Port {
    param (
        [string]$hostname,
        [int]$port
    )

    try {
        $connection = Test-NetConnection -ComputerName $hostname -Port $port
        if ($connection.TcpTestSucceeded) {
            return $true
        }
        return $false
    } catch {
        Write-Host "Error connecting to $hostname on port $port"
        return $false
    }
}

# Loop through each domain and check ports 80 and 443
foreach ($domain in $domains) {
    Write-Host "`nChecking domain: $domain"

    # Initialize a result object for this domain
    $domainResult = [PSCustomObject]@{
        Domain        = $domain
        Port80Status  = "Closed"
        Port443Status = "Closed"
        SSLIssuer     = "N/A"
        SSLSubject    = "N/A"
        SSLExpiration = "N/A"
        SSLThumbprint = "N/A"
    }

    # Check port 80
    if (Check-Port -hostname $domain -port 80) {
        $domainResult.Port80Status = "Open"
    }

    # Check port 443
    if (Check-Port -hostname $domain -port 443) {
        $domainResult.Port443Status = "Open"

        # If port 443 is open, get SSL certificate details
        $certDetails = D:\RewriteMaps\CertificateScanner.ps1 -SiteToScan $domain
        
        if ($certDetails) {
            $domainResult.SSLIssuer = $certDetails.Issuer
            $domainResult.SSLSubject = $certDetails.Subject
            $domainResult.SSLExpiration = $certDetails.EndDate
            #$domainResult.SSLThumbprint = $certDetails.Thumbprint
        }
        
    }

    # Add the domain result to the results array
    $results += $domainResult
}

# Export the results to an HTML report
$reportPath = "D:\RewriteMaps\ssl_report.html"
$results | ConvertTo-Html -Property Domain, Port80Status, Port443Status, SSLIssuer, SSLSubject, SSLExpiration -Head "<style>table { width: 100%; border-collapse: collapse; } th, td { padding: 8px; text-align: left; border-bottom: 1px solid #ddd; } th { background-color: #f2f2f2; }</style>" -Title "SSL Port and Certificate Report" | Out-File $reportPath
Write-Host "Report has been saved to $reportPath"
