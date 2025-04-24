# Define your LDAP filter to find Windows Server machines in AD
$ldapFilter = "(&(operatingSystem=*Windows Server*)(|(managedBy=CN=resManagedBy-AppTech-WebSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)(managedBy=CN=resManagedBy-AppTech-SiteCoreSupport,OU=SecurityGroups,OU=KelseyGroups,DC=ksnet,DC=com)))"

# Define Active Directory search base (can be set to the root or an OU)
$searchBase = "DC=ksnet,DC=com"

# Query AD for servers matching the filter
$servers = Get-ADComputer -LDAPFilter $ldapFilter -SearchBase $searchBase -Properties Name

# Define registry paths for uninstall strings
$registryPaths = @(
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    "Software\Microsoft\Windows\CurrentVersion\Uninstall" # For user-specific apps (HKCU)
)

# Initialize an empty array to store all reports
$allReports = @()

# Function to retrieve installed applications and uninstall strings from a remote server
function Get-InstalledAppsFromRegistry {
    param (
        [string]$server
    )

    $installedApps = @()

    foreach ($path in $registryPaths) {
        # Access registry remotely via Invoke-Command (or use New-PSDrive for remote registry access)
        $command = {
            param($regPath)

            $installedApps = @()

            # Check if registry path exists
            if (Test-Path "HKLM:\$regPath") {
                # Get subkeys (applications)
                $keys = Get-ChildItem -Path "HKLM:\$regPath"

                foreach ($key in $keys) {
                    $appName = (Get-ItemProperty -Path $key.PSPath -Name "DisplayName" -ErrorAction SilentlyContinue)
                    $uninstallString = (Get-ItemProperty -Path $key.PSPath -Name "UninstallString" -ErrorAction SilentlyContinue)

                    if ($appName -and $uninstallString) {
                        $installedApps += [PSCustomObject]@{
                            ApplicationName = $appName.DisplayName
                            UninstallString = $uninstallString.UninstallString
                        }
                    }
                }
            }

            return $installedApps
        }

        # Execute the command remotely on the server
        $apps = Invoke-Command -ComputerName $server -ScriptBlock $command -ArgumentList $path
        $installedApps += $apps
    }

    return $installedApps
}

# Loop through each remote server and get the installed apps
foreach ($server in $servers) {
    try {
        Write-Host "Retrieving applications from $($server.Name)..."

        # Get installed applications from the remote server
        $apps = Get-InstalledAppsFromRegistry -server $server.Name

        if ($apps.Count -gt 0) {
            # Create CSV and HTML report paths
            $csvReportPath = "D:\Reports\InstalledApps\$($server.Name)-InstalledApps.csv"
            $htmlReportPath = "D:\Reports\InstalledApps\$($server.Name)-InstalledApps.html"

            # Export to CSV
            $apps | Export-Csv -Path $csvReportPath -NoTypeInformation

            # Create HTML report with vibrant CSS
            $htmlContent = $apps | ConvertTo-Html -Property ApplicationName, UninstallString -PreContent "<h1>Installed Applications on $($server.Name)</h1>" -PostContent "<footer>Generated on $(Get-Date)</footer>"

            # Add vibrant CSS
            $css = @"
<style>
    body {
        font-family: Arial, sans-serif;
        color: #333;
        background-color: #f4f4f4;
        margin: 0;
        padding: 0;
    }
    h1 {
        background-color: #0078D4;
        color: white;
        padding: 10px;
        text-align: center;
    }
    table {
        width: 100%;
        margin: 20px 0;
        border-collapse: collapse;
        background-color: #ffffff;
    }
    th, td {
        padding: 10px;
        text-align: left;
        border: 1px solid #ddd;
    }
    th {
        background-color: #0078D4;
        color: white;
    }
    tr:nth-child(even) {
        background-color: #f2f2f2;
    }
    tr:hover {
        background-color: #ddd;
    }
    footer {
        text-align: center;
        font-size: 12px;
        color: #888;
    }
</style>
"@
            $htmlContent = $htmlContent -replace "</head>", "$css</head>"
            $htmlContent | Out-File -FilePath $htmlReportPath

            Write-Host "Reports generated for $($server.Name): $csvReportPath and $htmlReportPath"
        } else {
            Write-Host "No applications found for $($server.Name)."
        }
    }
    catch {
        Write-Host "Failed to retrieve applications from $($server.Name). Error: $_"
    }
}
