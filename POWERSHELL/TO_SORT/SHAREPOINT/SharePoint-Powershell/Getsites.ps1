# Add SharePoint PowerShell snap-in
if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

# Set the SharePoint site URL
$siteUrl = "https://pulse.kscpulse.com"

# Output CSV file path
$csvFilePath = "C:\Path\output.csv"

# Get all site collections
$siteCollections = Get-SPSite -Limit All -WebApplication $siteUrl

# Initialize an array to store site collection details
$siteCollectionDetails = @()

# Iterate through each site collection
foreach ($siteCollection in $siteCollections) {
    # Get site collection properties
    $siteCollectionDetails += [PSCustomObject]@{
        "Site Collection URL" = $siteCollection.Url
        "Title" = $siteCollection.RootWeb.Title
    }
}

# Export site collection details to CSV file
$siteCollectionDetails | Export-Csv -Path $csvFilePath -NoTypeInformation

# Dispose SharePoint objects
$siteCollections.Dispose()
