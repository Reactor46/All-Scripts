

Set-ExecutionPolicy Unrestricted

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

New-SPConfigurationDatabase -DatabaseName "SP2013Ext_CentralAdmin_Config" -DatabaseServer "SPProd\KSExternal" -Passphrase (ConvertTo-SecureString "KSCSP2013x_Prod" -AsPlainText -force) -FarmCredentials (Get-Credential)