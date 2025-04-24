Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop 

$sptokensvc= Get-SPSecurityTokenServiceConfig
$sptokensvc.FormsTokenLifetime = (New-TimeSpan -minutes 1440)
$sptokensvc.WindowsTokenLifetime = (New-TimeSpan -minutes 1440)
$sptokensvc.LogonTokenCacheExpirationWindow = (New-TimeSpan -minutes 1)
$sptokensvc.Update()
#iisreset