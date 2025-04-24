# Define the list of domains
$Domains = Get-Content -Path "D:\RewriteMaps\Combined-Domain-List.txt"

# Function to check domain status and gather information
function Test-Domain {
    param (
        [string]$domain
    )

    # Check if domain is active
    $pingResult = Test-Connection -ComputerName $domain -Count 1 -ErrorAction SilentlyContinue
    if (!$pingResult) {
        Write-Host "Domain Not Found" -ForegroundColor Red
        return
    }

    # Check ports 80 and 443
    $ports = @(80, 443)
    foreach ($port in $ports) {
        $tcpConnection = Test-NetConnection -ComputerName $domain -Port $port
        if ($tcpConnection.TcpTestSucceeded) {
            Write-Host "Port $port is open on $domain"
        } else {
            Write-Host "Port $port is closed on $domain"
        }
    }

    # Check if domain redirects
    $redirect = Invoke-WebRequest -Uri "http://$domain" -MaximumRedirection 0 -ErrorAction SilentlyContinue
    if ($redirect.StatusCode -eq 301 -or $redirect.StatusCode -eq 302) {
        Write-Host "Domain redirects to $($redirect.Headers.Location)"
    } else {
        Write-Host "No redirection detected for $domain"
    }

    # Check SSL certificate if port 443 is open
    if ($tcpConnection.TcpTestSucceeded -and $port -eq 443) {
        $sslDetails = [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        $sslRequest = [System.Net.HttpWebRequest]::Create("https://$domain")
        $sslRequest.GetResponse() | Out-Null
        $cert = $sslRequest.ServicePoint.Certificate
        $certDetails = New-Object PSObject -Property @{
            Issuer = $cert.Issuer
            Subject = $cert.Subject
            DateCreated = $cert.GetEffectiveDateString()
            DateExpires = $cert.GetExpirationDateString()
        }
        Write-Host "SSL Certificate Details:"
        $certDetails | Format-List

        # Check SSL expiration
        $expirationDate = [datetime]::Parse($certDetails.DateExpires)
        $daysLeft = ($expirationDate - (Get-Date)).Days
        if ($daysLeft -le 30) {
            Write-Host "SSL certificate expires in $daysLeft days" -ForegroundColor Red
        } elseif ($daysLeft -le 90) {
            Write-Host "SSL certificate expires in $daysLeft days" -ForegroundColor Yellow
        } else {
            Write-Host "SSL certificate expires in $daysLeft days" -ForegroundColor Green
        }
    }

    # Gather Whois information
    $whoisUrl = "https://api.ip2whois.com/v2?key=3F98CEAB0C7761C6A21471C12D3393D5&domain=$($domain)"
    $whoisInfo = Invoke-RestMethod -Uri $whoisUrl
    $whoisDetails = New-Object PSObject @{
        DomainName = $whoisInfo.domain
        Registrar = $whoisInfo.registrar.name
        WhoIsServer = $whoisInfo.whois_server
        NameServers = $whoisInfo.nameservers -join ", "
        LastUpdated = $whoisInfo.update_date
        Created = $whoisInfo.create_date
        Expiration = $whoisInfo.expire_date
        DomainAge = $whoisInfo.domain_age
    }
    Write-Host "Whois Information:"
    $whoisDetails | Format-List
}

# Loop through each domain and test
foreach ($domain in $Domains) {
    Test-Domain -domain $domain
}

# Output results in desired format
$OutputFormat = "html" # Change to "xml" or "csv" as needed
switch ($OutputFormat) {
    "html" { $whoisDetails | ConvertTo-Html | Out-File "domain_report.html" }
    "xml" { $whoisDetails | Export-Clixml -Path "domain_report.xml" }
    "csv" { $whoisDetails | Export-Csv -Path "domain_report.csv" -NoTypeInformation }
}