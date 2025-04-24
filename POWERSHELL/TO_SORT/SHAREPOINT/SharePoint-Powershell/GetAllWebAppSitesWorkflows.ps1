Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Specify the output file path
$csvPath = "D:\Sites\PulseSitesWF.csv"

# Initialize an empty array to store the site data
$sitesData = @()

# Get all site collections
$siteCollections = Get-SPSite -Limit All

# Iterate through each site collection
foreach ($siteCollection in $siteCollections) {
    # Get the root web of the site collection
    $rootWeb = $siteCollection.RootWeb
    
    # Get all webs (sites) within the root web
    $webs = $rootWeb.Webs
    
    # Iterate through each web
    foreach ($web in $webs) {
        # Get the web URL, title, and description
        $webUrl = $web.Url
        $webTitle = $web.Title
        $webDescription = $web.Description
        
        # Get all workflows associated with the web
        $workflows = $web.Workflows
        
        # Iterate through each workflow
        foreach ($workflow in $workflows) {
            # Get the workflow name and status
            $workflowName = $workflow.Name
            $workflowStatus = $workflow.InternalState
            
            # Create an object to store the site data
            $siteData = New-Object PSObject
            $siteData | Add-Member -MemberType NoteProperty -Name "Site Collection URL" -Value $siteCollection.Url
            $siteData | Add-Member -MemberType NoteProperty -Name "Web URL" -Value $webUrl
            $siteData | Add-Member -MemberType NoteProperty -Name "Web Title" -Value $webTitle
            $siteData | Add-Member -MemberType NoteProperty -Name "Web Description" -Value $webDescription
            $siteData | Add-Member -MemberType NoteProperty -Name "Workflow Name" -Value $workflowName
            $siteData | Add-Member -MemberType NoteProperty -Name "Workflow Status" -Value $workflowStatus
            
            # Add the site data object to the array
            $sitesData += $siteData
        }
        
        # Dispose of the web object
        $web.Dispose()
    }
    
    # Dispose of the site collection object
    $siteCollection.Dispose()
}

# Export the site data array to a CSV file
$sitesData | Export-Csv -Path $csvPath -NoTypeInformation

Write-Host "Export completed. The site data is saved to: $csvPath"
