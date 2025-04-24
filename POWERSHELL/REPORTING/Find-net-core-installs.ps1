# Define registry paths for .NET-related installations
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

# Define a list of keywords to filter specific .NET components
$keywords = @(
    ".NET",
    "ASP.NET",
    "Microsoft Windows Desktop Runtime",
    "Microsoft Windows Desktop Targeting Pack"
)

# Function to retrieve .NET-related installation details
function Get-DotNetInstallations {
    param ([string[]]$Paths, [string[]]$FilterKeywords)

    $dotNetInfo = @()

    foreach ($path in $Paths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
                $properties = Get-ItemProperty $_.PSPath
                if ($properties -and $FilterKeywords | ForEach-Object { $properties.DisplayName -like "*$_*" }) {
                    $dotNetInfo += [PSCustomObject]@{
                        Name            = $properties.DisplayName
                        Version         = $properties.DisplayVersion
                        InstalledDate   = $properties.InstallDate -as [datetime]
                        UninstallString = $properties.UninstallString
                    }
                }
            }
        }
    }

    return $dotNetInfo
}

# Retrieve .NET installation details
$dotNetInstallations = Get-DotNetInstallations -Paths $registryPaths -FilterKeywords $keywords

# Sort by version and installation date
$sortedDotNetInstallations = $dotNetInstallations | Sort-Object Version, InstalledDate

# Generate HTML report
$htmlFilePath = "DotNetInstallationsReport.html"

# CSS styling for the HTML report
$css = @"
<style>
    body { font-family: Arial, sans-serif; margin: 20px; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background-color: #f4f4f4; }
    tr:nth-child(even) { background-color: #f9f9f9; }
    tr:hover { background-color: #f1f1f1; }
</style>
"@

# Create the HTML content
$htmlContent = @"
<html>
<head>
    <title>.NET Installations Report</title>
    $css
</head>
<body>
    <h1>.NET Installations Report</h1>
    <table>
        <thead>
            <tr>
                <th>Name</th>
                <th>Version</th>
                <th>Installed Date</th>
                <th>Uninstall String</th>
            </tr>
        </thead>
        <tbody>
"@

foreach ($item in $sortedDotNetInstallations) {
    $htmlContent += @"
            <tr>
                <td>$($item.Name)</td>
                <td>$($item.Version)</td>
                <td>$($item.InstalledDate)</td>
                <td>$($item.UninstallString)</td>
            </tr>
"@
}

$htmlContent += @"
        </tbody>
    </table>
</body>
</html>
"@

# Save the HTML report to a file
Set-Content -Path $htmlFilePath -Value $htmlContent -Encoding UTF8

Write-Host "Report saved to $htmlFilePath"
