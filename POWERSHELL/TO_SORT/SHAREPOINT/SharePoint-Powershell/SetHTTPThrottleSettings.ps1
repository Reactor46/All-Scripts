Add-PSSnapin Microsoft.SharePoint.PowerShell –erroraction SilentlyContinue

$w = Get-SPWebApplication https://www.kelseycommunity.com
$w.HttpThrottleSettings
$w.Update()

