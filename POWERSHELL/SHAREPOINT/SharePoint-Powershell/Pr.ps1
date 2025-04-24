Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$service = Get-SPServiceInstance -Identity fa1cb1cc-e919-4cb7-b850-bde8f0a75fe1

$service.provision()

$service.update()

iisreset /noforce