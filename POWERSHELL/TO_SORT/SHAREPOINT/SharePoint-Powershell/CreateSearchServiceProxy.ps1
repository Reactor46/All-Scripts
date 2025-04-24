Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
$ssa = Get-SPEnterpriseSearchServiceApplication
New-SPEnterpriseSearchServiceApplicationProxy -Name 'Search Service Application Proxy' -SearchApplication $ssa

Write-Host "Search Service Application Proxy created. "