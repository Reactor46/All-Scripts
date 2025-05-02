Param([string] $queryText = 'SharePoint', [string] $siteCollectionUrl = "http://server/sitecollection/site")

# Add-PSSnapin -Name Microsoft.SharePoint.PowerShell
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Search")


$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($siteCollectionUrl)

$keywordQuery = New-Object Microsoft.SharePoint.Client.Search.Query.KeywordQuery($clientContext)
$keywordQuery.QueryText = $queryText

$searchExecutor = New-Object Microsoft.SharePoint.Client.Search.Query.SearchExecutor($clientContext)

$results = $searchExecutor.ExecuteQuery($keywordQuery)

$clientContext.ExecuteQuery()

$formattedResults = @()

foreach($result in $results.Value[0].ResultRows)
{
    $formattedResult = New-Object -TypeName PSObject
    foreach($key in $result.Keys)
    {
        $formattedResult | Add-Member -Name $key -MemberType NoteProperty -Value $result[$key]
    }
    $formattedResults += $formattedResult
}

$formattedResults