Add-PSSnapin microsoft.sharepoint.powershell -ea SilentlyContinue

Get-SPEnterpriseSearchServiceInstance -Local | Start-SPEnterpriseSearchServiceInstance