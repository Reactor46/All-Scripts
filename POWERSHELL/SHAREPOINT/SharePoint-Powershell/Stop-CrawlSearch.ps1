Add-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue
Get-SPEnterpriseSearchCrawlContentSource -SearchApplication "Search Service - eFlipChart" | ForEach-Object {
    if ($_.CrawlStatus -ne "Idle")
    {
        Write-Host "Stopping currently running crawl for content source $($_.Name)"
        $_.StopCrawl()
        
        do { Start-Sleep -Seconds 1 }
        while ($_.CrawlStatus -ne "Idle")
    }
}

Get-SPEnterpriseSearchCrawlContentSource -SearchApplication "Intranet Search Service Application" | ForEach-Object {
    if ($_.CrawlStatus -ne "Idle")
    {
        Write-Host "Stopping currently running crawl for content source $($_.Name)"
        $_.StopCrawl()
        
        do { Start-Sleep -Seconds 1 }
        while ($_.CrawlStatus -ne "Idle")
    }
}
Remove-PSSnapin Microsoft.SharePoint.PowerShell -erroraction SilentlyContinue