Add-PSSnapin "Microsoft.SharePoint.PowerShell"
# ----- For publishing sites and non publishing sites
$site = Get-SPSite http://ServerName
$customMasterPage = "/_catalogs/masterpage/name.master"
$defaultMasterPage = "/_catalogs/masterpage/name.master"

# -----Changing custom MasterPage
foreach ($web in $site.AllWebs) {
    $web; $web.CustomMasterUrl = $customMasterPage ; 
    $web.Update(); $web.CustomMasterUrl;
    $web.Dispose()
}

# ----- Change Default MasterPage
foreach ($web in $site.AllWebs) {
    $web; $web.MasterUrl = $defaultMasterPage ; 
    $web.Update(); $web.MasterUrl;
    $web.Dispose()
}
$site.Dispose()

write-host "Complete! New MasterPage is now applied";