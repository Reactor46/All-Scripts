Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue

$dbs = Get-SPContentDatabase
$result = @();
foreach($db in $dbs){
    $obj = New-Object PSObject
    $obj | Add-Member -Type NoteProperty -Name "Type" -Value $db.Type
    $obj | Add-Member -Type NoteProperty -Name "Attached" -Value $db.IsAttachedToFarm
    $obj | Add-Member -Type NoteProperty -Name "Exists" -Value $db.ExistsInFarm
    $obj | Add-Member -Type NoteProperty -Name "Site Count" -Value $db.CurrentSiteCount
    $obj | Add-Member -Type NoteProperty -Name "Paired" -Value $db.IsPaired
    $obj | Add-Member -Type NoteProperty -Name "Can Migrate" -Value $db.CanMigrate
    $obj | Add-Member -Type NoteProperty -Name "Normalized Data Source" -Value $db.NormalizedDataSource
    $obj | Add-Member -Type NoteProperty -Name "Service Instance" -Value $db.ServiceInstance
    #$obj | Add-Member -Type NoteProperty -Name "Failover Server" -Value $db.FailoverServer
    #$obj | Add-Member -Type NoteProperty -Name "Failover Service Instance" -Value $db.FailoverServiceInstance
    $obj | Add-Member -Type NoteProperty -Name "URL" -Value $db.WebApplication.URL
    $obj | Add-Member -Type NoteProperty -Name "Site Names" -Value $db.WebApplication.Name
    $obj | Add-Member -Type NoteProperty -Name "Name" -Value $db.Name
    $obj | Add-Member -Type NoteProperty -Name "DatabaseServer" -Value $db.Server
    $obj | Add-Member -Type NoteProperty -Name "Database Connection String" -Value $db.DatabaseConnectionString
    $obj | Add-Member -Type NoteProperty -Name "MaximumSiteCount" -Value $db.MaximumSiteCount
    $obj | Add-Member -Type NoteProperty -Name "WarningSiteCount" -Value $db.WarningSiteCount
    $result += $obj
}

$result | Export-Csv .\$env:COMPUTERNAME-DatabaseExport.csv -NoTypeInformation