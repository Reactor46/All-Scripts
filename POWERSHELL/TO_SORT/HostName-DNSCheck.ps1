# Array of hostnames
$hostnames = @("FBV-SCRCD10-T01-3","FBV-SCRCD10-T02-3","FBV-SCRCD10-T01-2","FBV-SCRCD10-T02-2")

# Function to resolve a hostname and test connectivity
function Resolve-And-TestConnectivity {
    param (
        [string]$hostname
    )

    $result = @()

    try {
        # Resolve the hostname to an IP address
        $ipEntry = [System.Net.Dns]::GetHostAddresses($hostname) | Where-Object { $_.AddressFamily -eq 'InterNetwork' }

        if ($ipEntry) {
            foreach ($ip in $ipEntry) {
                # Test connectivity (ping) to the resolved IP address
                $pingResult = Test-Connection -ComputerName $ip.IPAddressToString -Count 1 -Quiet

                $result += [pscustomobject]@{
                    Hostname       = $hostname
                    IPAddress      = $ip.IPAddressToString
                    Connectivity   = if ($pingResult) { "Success" } else { "Failed" }
                }
            }
        }
        else {
            $result += [pscustomobject]@{
                Hostname       = $hostname
                IPAddress      = "Resolution Failed"
                Connectivity   = "N/A"
            }
        }
    }
    catch {
        $result += [pscustomobject]@{
            Hostname       = $hostname
            IPAddress      = "Error"
            Connectivity   = $_.Exception.Message
        }
    }

    return $result
}

# Initialize an array to store results
$results = @()

# Resolve and test connectivity for each hostname
foreach ($hostname in $hostnames) {
    Write-Output "Resolving and testing connectivity for $hostname"
    $results += Resolve-And-TestConnectivity -hostname $hostname
}

# Output results
$results | Format-Table -AutoSize

# Optionally, export results to a CSV file
$outputFile = "D:\ConnectivityResults.csv"
$results | Export-Csv -Path $outputFile -NoTypeInformation

Write-Output "Hostname resolution and connectivity testing completed and saved to $outputFile"
