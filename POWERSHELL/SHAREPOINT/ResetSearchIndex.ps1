Set-ExecutionPolicy Unrestricted

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$SearchServiceApp = Get-SPEnterpriseSearchServiceApplication -identity "Search Service - eFlipChart" 
Write-Host $SearchServiceApp
$SearchServiceApp.Reset($true, $true)




