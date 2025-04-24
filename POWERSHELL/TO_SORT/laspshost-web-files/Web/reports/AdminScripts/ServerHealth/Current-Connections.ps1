# Define the site, host, and page
$siteName = "ksprod-new-cd.ksnet.com"  # Replace with your site name
$hostName = $env:COMPUTERNAME         # Replace with your host name
$pagePath = "/"                 # Replace with your page path

# Define the counter paths
$siteCounterPath = "\Web Service($siteName)\Current Connections"
$hostCounterPath = "\Web Service($siteName)\$hostName Current Connections"
$pageCounterPath = "\Web Service($siteName)\$hostName\$pagePath Current Connections"

# Get counter values
$siteConnections = (Get-Counter -Counter $siteCounterPath).CounterSamples.CookedValue
$hostConnections = (Get-Counter -Counter $hostCounterPath).CounterSamples.CookedValue
$pageConnections = (Get-Counter -Counter $pageCounterPath).CounterSamples.CookedValue

# Display the number of connections
Write-Host "Active connections to site '$siteName': $siteConnections"
Write-Host "Active connections to host '$hostName' on site '$siteName': $hostConnections"
Write-Host "Active connections to page '$pagePath' on host '$hostName' on site '$siteName': $pageConnections"
