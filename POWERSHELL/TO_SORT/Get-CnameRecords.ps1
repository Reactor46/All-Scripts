function Get-CnameRecords {
    param (
        [string]$ip
    )
    
    try {
        # Perform a reverse DNS lookup to get the hostname
        $hostEntry = [System.Net.Dns]::GetHostEntry($ip)
        $hostname = $hostEntry.HostName

        # Perform a forward DNS lookup on the hostname to check for CNAME records
        $dnsQuery = Resolve-DnsName -Name $hostname -Type CNAME -ErrorAction Stop

        $cnameRecords = $dnsQuery | ForEach-Object {
            [pscustomobject]@{
                IPAddress = $ip
                Hostname = $hostname
                CName = $_.Name
                CanonicalName = $_.AliasName
            }
        }

        return $cnameRecords
    }
    catch {
        Write-Output "No CNAME records found for IP $ip or failed to resolve. Error: $_"
        return $null
    }
}