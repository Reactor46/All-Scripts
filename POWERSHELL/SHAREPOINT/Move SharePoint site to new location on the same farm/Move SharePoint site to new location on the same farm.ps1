# add SharePoint snapin
Add-PSSnapin Microsoft.SharePoint.PowerShell –ea SilentlyContinue

# set variables
$exportfolder = "C:\Site exports"
$exportfile = "\site_export.cmp"
$exportsite = "http://vm353/PWA/PAULMATHERTESTSITE"
$exportlocation = $exportfolder+$exportfile
$importlocation = "http://vm353/PAULMATHERTESTSITE"

#get export site's template
$web = Get-SPWeb $exportsite
$webTemp =  $web.WebTemplate
$webTempID = $web.Configuration
$webTemplate = "$webTemp#$webTempID"
$web.Dispose()

#create export folder
$null = New-Item $exportfolder -type directory
#export site
Export-SPWeb $exportsite –Path $exportlocation -IncludeUserSecurity -IncludeVersions 4
Write-host "$exportsite has been exported to $exportlocation"
#create new site ready for import
$null = New-SPWeb $importlocation -Template "$webTemplate"
Write-host "$importlocation created ready for import"
#import site
Import-SPWeb $importlocation –Path $exportlocation -IncludeUserSecurity –UpdateVersions 2
Write-host "$exportsite has been imported to $importlocation" -foregroundcolor "Green"


