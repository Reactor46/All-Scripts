Add-PSSnapin microsoft.sharepoint.powershell -ea SilentlyContinue

$db2delete = Get-SPDatabase “76b508b8-4ac5-441f-8571-feeef6255510”
Write-Host $db2delete
$db2delete
$db2delete.status
$db2delete.Delete()