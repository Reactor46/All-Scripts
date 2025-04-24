# Load the SharePoint PowerShell module (if not already loaded)
if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

# Get all site collections
$siteCollections = Get-SPSite -Limit All

# Create an array to store site collection data
$siteCollectionData = @()

# Loop through and collect information about each site collection
foreach ($siteCollection in $siteCollections)
{
    $siteCollectionData += [PSCustomObject]@{
        "SiteCollectionURL" = $siteCollection.Url
        "SiteCollectionOwner" = $siteCollection.Owner
        "SiteCollectionTitle" = $siteCollection.RootWeb.Title
    }
}

# Export site collection data to a CSV file
$siteCollectionData | Export-Csv -Path "D:\SPScripts\SiteCollections.csv" -NoTypeInformation -Append

# Dispose of the site collection objects to release resources
$siteCollections.Dispose()
