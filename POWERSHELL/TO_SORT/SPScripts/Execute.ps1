."D:\SPScripts\SharePoint2016Config.ps1"
 $CredSSPDelegates = "*.ksnet.com"
 $SPBinaryPath = "F:\"
 $SPBinaryPathCredential = Get-Credential -UserName "contoso\sp_install" -Message "SP_Install"
 $FarmAccount = Get-Credential -UserName "KELSEY-SEYBOLD\sp_farm" -Message "SP_Farm"
 $InstallAccount = Get-Credential -UserName "KELSEY-SEYBOLD\sp_install" -Message "SP_Install"
 $ProductKey ="7G7R6-N6QJC-JFPJX-CK8WX-66QW4"
 $DatabaseServer = "SP2016-DSCTBD"
 $FarmPassPhrase = "pass@word1"
 $WebPoolManagedAccount = Get-Credential -UserName "KELSEY-SEYBOLD\sp_webapp" -Message "SP_WebApp"
 $ServicePoolManagedaccount = Get-Credential -UserName "KELSEY-SEYBOLD\sp_serviceapps" -Message "SP_ServiceApps"
 $WebAppUrl = "SP2016-DSCTBD"
 $TeamSiteUrl = "/"
 $MySiteHostUrl = "/personnal/"
 $CacheSizeInMB = 300
 $ConfigData = @{
 AllNodes = @(
 @{
 NodeName = "SP2016-DSCTBD"
 PSDscAllowPlainTextPassword = $true
 })
 }
 SharePointServer -ConfigurationData $ConfigData -CredSSPDelegates $CredSSPDelegates -SPBinaryPath $SPBinaryPath -ULSViewerPath "D:\Tools\ULSViewer.exe" -SPBinaryPathCredential $SPBinaryPathCredential -FarmAccount $FarmAccount -InstallAccount $InstallAccount -ProductKey $ProductKey -DatabaseServer $DatabaseServer -FarmPassPhrase $FarmPassPhrase -WebPoolManagedAccount $WebPoolManagedAccount -ServicePoolManagedAccount $ServicePoolManagedAccount -WebAppUrl $WebAppurl -TeamSiteUrl $TeamSiteUrl -MySiteHostUrl $MySiteHostUrl -CacheSizeInMB $CacheSizeInMB