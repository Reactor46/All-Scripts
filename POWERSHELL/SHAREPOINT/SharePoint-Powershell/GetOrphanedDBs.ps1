Add-PSSnapin Microsoft.SharePoint.PowerShell –erroraction SilentlyContinue

Get-SPDatabase
Write-Host ""

#Find all deleted databases which are still referenced in SharePoint Farm
$OrphanDBs = Get-SPDatabase | Where {$_.Exists -eq $false}
$OrphanDBs

