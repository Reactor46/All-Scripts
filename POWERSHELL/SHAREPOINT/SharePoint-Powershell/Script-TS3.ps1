Asnp *sh* # same as Add-PSSnapin microsoft.sharepoint.powershell -ea SilentlyContinue
$farm = Get-SPFarm
$obj = $farm.GetObject('53ff7253-c346-41db-9399-716459d5c65e')
$obj.Status = [Microsoft.SharePoint.Administration.SPObjectStatus]::Online
$obj.Update()