$servers = get-content c:\users\wreeves\desktop\servers.txt

foreach ( $server in $servers )
    {
    Write-Host "Getting Services for $server"':' -ForegroundColor Green
    Get-wmiobject win32_service -ComputerName $server -ErrorAction SilentlyContinue | ? { $_.startname -match "ccmhmr" } | Format-Table name, startname, startmode
    }
