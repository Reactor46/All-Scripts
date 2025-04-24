

Set-ExecutionPolicy Unrestricted

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$spapp = Get-SPServiceApplication -Name "Search Service Application"
Write-Host $spapp

Write-Host "Removing Search Service Application"
#Remove-SPServiceApplication $spapp -RemoveData
Remove-SPServiceApplication $spapp 

Write-Host "Finished"

