<# # Load the SharePoint PowerShell module (if not already loaded)
if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
} #>

# Get the SharePoint service context
$sitePreFix = "https://prod-pulse.kscpulse.com"
$sites = Get-Content -Path "D:\SPScripts\Sites.txt"

ForEach($Url in $sites){
# Get the SharePoint service context for the site collection
$site = Get-SPSite $sitePrefix$Url

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
            $termSet.Name
        }
    }
}}

# Dispose of the site object to release resources
#$site.Dispose()
