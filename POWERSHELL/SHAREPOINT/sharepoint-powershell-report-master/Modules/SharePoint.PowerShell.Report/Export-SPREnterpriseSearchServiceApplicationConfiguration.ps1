function Export-SPREnterpriseSearchServiceApplicationConfiguration {
    param(
        [string]$Path
    )

    $file = '{0}\SPEnterpriseSearchServiceApplication.xml' -f $Path
    Start-Job -ScriptBlock {
        
        Add-PSSnapin -Name Microsoft.SharePoint.PowerShell

        $searchServiceApplications = Get-SPEnterpriseSearchServiceApplication
        $object = @()
        foreach ($searchServiceApplication in $searchServiceApplications)
        {
          $content = New-Object -TypeName Microsoft.Office.Server.Search.Administration.Content -ArgumentList $searchServiceApplication
          $searchServiceApplication | Add-Member -MemberType NoteProperty -Name 'Content' -Value $content

          $crawlContentSource = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $searchServiceApplication
          $searchServiceApplication | Add-Member -MemberType NoteProperty -Name 'CrawlContentSources' -Value $crawlContentSource

          $searchCrawlDatabase = Get-SPEnterpriseSearchCrawlDatabase -SearchApplication $searchServiceApplication
          $searchServiceApplication | Add-Member -MemberType NoteProperty -Name 'SearchCrawlDatabase' -Value $searchCrawlDatabase

          $searchFileExtension = Get-SPEnterpriseSearchCrawlExtension -SearchApplication $searchServiceApplication
          $searchServiceApplication | Add-Member -MemberType NoteProperty -Name 'CrawlFileExtensions' -Value $searchFileExtension

          $searchOwner = Get-SPEnterpriseSearchOwner -Level Ssa
          $searchQueryAuthority = Get-SPEnterpriseSearchQueryAuthority -SearchApplication $searchServiceApplication -Owner $searchOwner
          $searchServiceApplication | Add-Member -MemberType NoteProperty -Name 'SearchQueryAuthority' -Value $searchQueryAuthority

          $object += $searchServiceApplication
        }

        $object | Export-Clixml -Path $args[0] -Depth 3
    } -ArgumentList $file

}