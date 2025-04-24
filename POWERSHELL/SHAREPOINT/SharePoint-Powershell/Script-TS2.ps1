Asnp *sh* # same as Add-PSSnapin microsoft.sharepoint.powershell -ea SilentlyContinue
$farm = Get-SPFarm
$obj = $farm.GetObject('5792f3ea-784c-4f97-8537-6eb9b062ff38')
$obj.Status = [Microsoft.SharePoint.Administration.SPObjectStatus]::Online
$obj.Update()