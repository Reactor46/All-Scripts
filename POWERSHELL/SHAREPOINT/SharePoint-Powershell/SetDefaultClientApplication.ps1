Set-ExecutionPolicy Unrestricted

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$web = Get-SPWeb https://teams.kscpulse.com/sites/BPA
$list = $web.Lists["BPAs"]
$list.DefaultItemOpen = [Microsoft.SharePoint.DefaultItemOpen]::PreferClient
$list.Update()
