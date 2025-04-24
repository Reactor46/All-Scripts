Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 

$site = Get-SPSite https://apps.kscpulse.com/sites/Enrollment 
$siteId = $site.Id
$siteDatabase = $site.ContentDatabase 
$siteDatabase.ForceDeleteSite($siteId, $false, $false)