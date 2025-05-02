Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Connect to the SharePoint Farm
#Connect-SPFarm -ErrorAction Stop

# Function to search for lists with "providerrequest" in their names
function SearchLists($web) {
    #Write-Host "Enter Web: " $web.Name
    # Get all lists in the current web
    $lists = $web.Lists


    
    foreach ($list in $lists) {
        #Write-Host "List Title: $($list.Title)"
        # Check if the list name contains "providerrequest"
        if ($list.Title -like "*Request*") {
            $data = @(
                [PSCustomObject]@{
                     Name = "List Title: $($list.Title)";
                     ListView = "List URL: $($list.DefaultViewUrl)";
                     Space = " "
                     }
            )
            # Export data to CSV
            $data | Export-Csv -Path $outputPath -NoTypeInformation -Append
            Write-Host "List Title: $($list.Title)"
            Write-Host "List URL: $($list.DefaultViewUrl)"
            Write-Host ""
            
        }
    }

    # Recursively search in all subsites
    foreach ($subweb in $web.Webs) {
        SearchLists $subweb
    }
}


# Get all web applications in the SharePoint Farm
$webApplications = Get-SPWebApplication

# Output file path
$outputPath = "d:\output.csv"

# Iterate through each web application
foreach ($webApp in $webApplications) {
    Write-Host "Web Application: $($webApp.Name)"

    # Get all site collections in the web application
    $siteCollections = $webApp.Sites

    # Iterate through each site collection
    foreach ($siteCollection in $siteCollections) {
        #Write-Host "`tSite Collection: $($siteCollection.Url)"

        # Get all sites in the site collection
        $sites = $siteCollection.AllWebs

        # Iterate through each site
        foreach ($site in $sites) {
            #Write-Host "`t`tSite: $($site.Url)"
            # Perform any desired operations on the site here
            
                        # Start searching from the root web of the site collection
            SearchLists $site

            # Dispose of the site object
            $site.Dispose()
        }

        # Dispose of the site collection object
        $siteCollection.Dispose()
    }
}



# Disconnect from the SharePoint Farm
#Disconnect-SPFarm
