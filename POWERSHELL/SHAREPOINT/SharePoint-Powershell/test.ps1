Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$searchServerName = (Get-ChildItem env:computername).value

Write-Host searchServerName