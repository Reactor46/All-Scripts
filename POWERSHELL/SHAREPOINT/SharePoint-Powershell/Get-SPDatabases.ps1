Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$spConfigDb = Get-SPDatabase | Where-Object { $_.Type -eq "Configuration Database" }
$spContentDbs = Get-SPDatabase | Where-Object { $_.Type -eq "Content Database" }

$databases = @()

$databases += [PSCustomObject]@{
    Type = "Configuration Database"
    Name = $spConfigDb.Name
    Server = $spConfigDb.Server
    Size_GB = $spConfigDb.DiskSizeRequired/1GB
}

foreach ($contentDb in $spContentDbs) {
    $databases += [PSCustomObject]@{
        Type = "Content Database"
        Name = $contentDb.Name
        Server = $contentDb.Server
        Size_GB = $contentDb.DiskSizeRequired/1GB
    }
}

$databases | Export-Csv -Path "D:\Reports\Databases.csv" -NoTypeInformation -Append