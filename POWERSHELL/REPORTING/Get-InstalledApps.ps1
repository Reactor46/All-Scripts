# Define registry paths for uninstall strings
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
)

# Initialize an array to hold the results
$installedApps = @()

# Loop through each registry path and get installed application details
foreach ($path in $registryPaths) {
    # Check if the registry path exists
    if (Test-Path $path) {
        # Get all subkeys (installed applications)
        $keys = Get-ChildItem -Path $path

        foreach ($key in $keys) {
            # Get the DisplayName and UninstallString (if available)
            $appName = (Get-ItemProperty -Path $key.PSPath -Name "DisplayName" -ErrorAction SilentlyContinue)
            $uninstallString = (Get-ItemProperty -Path $key.PSPath -Name "UninstallString" -ErrorAction SilentlyContinue)

            # If the DisplayName and UninstallString are available, add them to the array
            if ($appName -and $uninstallString) {
                $installedApps += [PSCustomObject]@{
                    ApplicationName = $appName.DisplayName
                    UninstallString = $uninstallString.UninstallString
                }
            }
        }
    }
}

# Output the results (to console, CSV or HTML as needed)

# Option 1: Display results in console
$installedApps | Format-Table -Property ApplicationName, UninstallString

# Option 2: Export results to a CSV file
$installedApps | Export-Csv -Path "D:\InstalledApps.csv" -NoTypeInformation

# Option 3: Export results to an HTML file
$installedApps | ConvertTo-Html -Property ApplicationName, UninstallString -Title "Installed Applications" | Out-File "D:\InstalledApps.html"

Write-Host "Reports generated: D:\InstalledApps.csv and D:\InstalledApps.html"
