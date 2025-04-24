# Function to perform WHOIS lookup
function Get-WhoisInfo {
    param(
        [parameter(Mandatory = $false)][string]$PublicIPaddressOrName
    )

    #Check if the module PSParseHTML is installed and install
    #the module if it's not installed
    if (-not (Get-Command ConvertFrom-HTMLClass -ErrorAction SilentlyContinue)) {
        Install-Module PSParseHTML -SkipPublisherCheck -Force:$true -Confirm:$false
    }

    try {
        #Get results from your own Public IP Address
        if (-not ($PublicIPaddressOrName)) {
            $ProgressPreference = "SilentlyContinue"
            $PublicIPaddressOrName = (Invoke-WebRequest -uri https://api.ipify.org?format=json | ConvertFrom-Json -ErrorAction Stop).ip
            $whoiswebresult = Invoke-Restmethod -Uri "https://www.whois.com/whois/$($PublicIPaddressOrName)" -TimeoutSec 15 -ErrorAction SilentlyContinue
            $whoisinfo = ConvertFrom-HTMLClass -Class 'whois-data' -Content $whoiswebresult -ErrorAction SilentlyContinue
            write-host ("Getting WHOIS details for {0}" -f $PublicIPaddressOrName) -ForegroundColor Green
        }
        #Get results from the Public IP or name specified
        else {
            $ProgressPreference = "SilentlyContinue"
            if ((($PublicIPaddressOrName).Split('.').Length -eq 4)) {
                $whoiswebresult = Invoke-Restmethod -Uri "https://www.whois.com/whois/$($PublicIPaddressOrName)" -TimeoutSec 15 -ErrorAction SilentlyContinue
                $whoisinfo = ConvertFrom-HTMLClass -Class 'whois-data' -Content $whoiswebresult -ErrorAction SilentlyContinue
                write-host ("Getting WHOIS details for {0}" -f $PublicIPaddressOrName) -ForegroundColor Green
            }
            else {
                $whoiswebresult = Invoke-Restmethod -Uri "https://www.who.is/whois/$($PublicIPaddressOrName)" -TimeoutSec 30 -ErrorAction SilentlyContinue
                $whoisinfo = ConvertFrom-HTMLClass -Class 'df-raw' -Content $whoiswebresult -ErrorAction SilentlyContinue
                write-host ("Getting WHOIS details for {0}" -f $PublicIPaddressOrName) -ForegroundColor Green
            }
        }
    
        Return $whoisinfo   
    }
    catch {
        Write-Warning ("Error getting WHOIS details")
    }
}

# Path to your domain list text file
$domainListFile = "D:\RewriteMaps\Redirect-Domains.txt"

# Read domains from the text file
$domains = Get-Content -Path $domainListFile

# Loop through each domain in the list
foreach ($domain in $domains) {
    $domain = $domain.Trim()

    # Resolve DNS for the domain
    try {
        $dnsInfo = Resolve-DnsName -Name $domain -ErrorAction Stop
        Write-Host "DNS for $domain resolved to: $($dnsInfo[0].IPAddress)"
        
        # Ping the domain
        $pingResult = Test-Connection -ComputerName $domain -Count 1 -Quiet
        
        if ($pingResult) {
            Write-Host "$domain is reachable via ping."

            # Get WHOIS information
            $whoisData = Get-WhoisInfo $domain
            Write-Host "WHOIS information for $domain :`n$whoisData"
        }
        else {
            Write-Host "$domain is not reachable via ping."
        }
    }
    catch {
        Write-Host "Error resolving DNS or pinging $domain. Skipping..."
    }
    
    Write-Host "------------------------------------------"
}

