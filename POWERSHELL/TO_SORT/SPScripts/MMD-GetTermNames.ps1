# Load the SharePoint PowerShell module (if not already loaded)
Add-PSSnapin Microsoft.SharePoint.PowerShell

# Get the SharePoint service context
#$site = Get-SPSite "https://prod-pulse.kscpulse.com"
$siteCollectionUrl = "https://prod-pulse.kscpulse.com/strategic-planning/risk-management/"

# Get the SharePoint service context for the site collection
$site = Get-SPSite $siteCollectionUrl
$taxonomySession = Get-SPTaxonomySession -Site $site

# Get all Term Stores
$termStores = $taxonomySession.TermStores

# Loop through Term Stores and get Term Set names
foreach ($termStore in $termStores)
{
    Write-Host "Term Store Name: $($termStore.Name)"
    Write-Host "-----------------------------------"
    
    $groups = $termStore.Groups
    
    foreach ($group in $groups)
    {
        $termSets = $group.TermSets

        foreach ($termSet in $termSets)
        {
            Write-Host "Term Set Name: $($termSet.Name)"
        }
    }
}

# Dispose of the site object to release resources
$site.Dispose()
