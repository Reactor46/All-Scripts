# Load the SharePoint PowerShell module (if not already loaded)
if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

# Specify the URL prefix to filter site collections
$urlPrefix = "https://prod-pulse.kscpulse.com"

# Get all site collections
$siteCollections = Get-SPSite -Limit All | Where-Object { $_.Url.StartsWith($urlPrefix) }


# Create an array to store site collection data
$siteCollectionData = @()

# Loop through and collect information about each site collection
foreach ($siteCollection in $siteCollections)
{
    $siteCollectionUrl = $siteCollection.Url
    $siteCollectionTitle = $siteCollection.RootWeb.Title

    # Get all SharePoint Groups within the site collection
    $groups = $siteCollection.RootWeb.SiteGroups

    # Get all Term Sets within the site collection
    #$taxonomySession = [Microsoft.SharePoint.Taxonomy.TaxonomySession]::GetTaxonomySession($siteCollection)
    $taxonomySession = Get-SPTaxonomySession -Site $siteCollectionUrl
    $termStores = $taxonomySession.TermStores

    # Create an array to store group and term set data
    $groupAndTermSetData = @()

    # Loop through SharePoint Groups and collect their names
    foreach ($group in $groups)
    {
        $groupAndTermSetData += [PSCustomObject]@{
            "SiteCollectionURL" = $siteCollectionUrl
            "SiteCollectionTitle" = $siteCollectionTitle
            "GroupName" = $group.Name
            "TermSetName" = $null  # Initialize Term Set name to null
        }
    }

    # Loop through Term Stores and collect Term Set names
    foreach ($termStore in $termStores)
    {
        $termSets = $termStore.Groups | ForEach-Object { $_.TermSets }

        foreach ($termSet in $termSets)
        {
            $groupAndTermSetData += [PSCustomObject]@{
                "SiteCollectionURL" = $siteCollectionUrl
                "SiteCollectionTitle" = $siteCollectionTitle
                "GroupName" = $null  # Initialize Group name to null
                "TermSetName" = $termSet.Name
            }
        }
    }

    $groupAndTermSetData | ForEach-Object { $siteCollectionData += $_ }
}

# Export site collection data to a CSV file
$siteCollectionData | Export-Csv -Path "D:\SPScripts\SiteCollectionsWithGroupsAndTermSets.csv" -NoTypeInformation -Append

# Dispose of the site collection objects to release resources
$siteCollections.Dispose()
