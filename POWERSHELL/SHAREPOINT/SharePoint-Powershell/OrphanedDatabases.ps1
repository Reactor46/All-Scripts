Add-PSSnapin microsoft.sharepoint.powershell -ea SilentlyContinue

$OrphanDBs = Get-SPDatabase | Where {$_.Exists -eq $false}
Write-Host $OrphanDBs
#clean any Orphans DB
#Get-SPDatabase | Where{$_.Exists -eq $false} | ForEach {$_.Delete()}