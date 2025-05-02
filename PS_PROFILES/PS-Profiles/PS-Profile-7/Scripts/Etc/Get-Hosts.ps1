param($netbase,$Name)
$ErrorActionPreference = "SilentlyContinue"
for ( $i -eq 1; $i -le 254; $i++ )
    {
    $ip = "$($netbase)$($i)"
    if ( $hostname = [System.Net.Dns]::GetHostEntry($ip).hostname )
        {
        $hostname = $hostname.replace('.ccmhmr.local','')
        $filename = "C:\Users\wreeves\Documents\IPs\" + $Name + ".txt"
        "|$($ip)|[[MIS:$($hostname)|$hostname]]|" | Out-File $filename -append
        }
    }