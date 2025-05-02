# Specify the URL of the site collection
$siteCollectionUrl = 'https://prod-pulse.kscpulse.com/'

# Get the SharePoint site collection
$site = Get-SPSite $siteCollectionUrl

# Get all global SharePoint groups for the site collection
$groups = $site.RootWeb.SiteGroups | Where-Object { $_.IsAssociatedMemberGroup -eq $true -or $_.IsAssociatedOwnerGroup -eq $true -or $_.IsAssociatedVisitorGroup -eq $true }

# Loop through and display the names of the global groups
foreach ($group in $groups)
{
    Write-Host "Global Group Name: $($group.Name)"
}

# Dispose of the site object to release resources
$site.Dispose()