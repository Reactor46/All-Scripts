# Add the SharePoint PowerShell Snap-in
if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell
}

# Create an empty array to store the site owners
$siteOwners = @()

# Get the SharePoint farm
$spFarm = Get-SPFarm

# Iterate through all site collections
foreach ($spWebApp in $spFarm.WebApplications) {
    foreach ($site in $spWebApp.Sites) {
        foreach ($web in $site.AllWebs) {
            # Get the site owners
            $owners = Get-SPUser -Web $web.Url | Where-Object { $_.IsSiteAdmin -eq $true }

            # Add site owners to the array
            foreach ($owner in $owners) {
                $siteOwners += [PSCustomObject]@{
                    "Site URL" = $web.Url
                    "Owner"    = $owner.LoginName
                }
            }

            # Dispose of the web object
            $web.Dispose()
        }
        # Dispose of the site object
        $site.Dispose()
    }
}
<#
# Export site owners to CSV
$siteOwners | Export-Csv -Path "SiteOwners.csv" -NoTypeInformation

# Export site owners to XLSX
$siteOwners | Export-Excel -Path "SiteOwners.xlsx" -WorksheetName "SiteOwners"

# Export site owners to HTML
$siteOwners | ConvertTo-Html -Property "Site URL", "Owner" | Out-File -FilePath "SiteOwners.html"
#>

$siteOwners