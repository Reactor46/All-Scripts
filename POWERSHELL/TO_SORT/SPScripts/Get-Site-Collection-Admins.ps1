# Get all site collections in SharePoint
$AllSites = Get-SPSite -Limit All

# Create an array list to store the results
$myArray = [System.Collections.ArrayList]@()

# Loop through each site collection URL
foreach ($url in $AllSites.Url) {

    # Get all users who are site collection admins
    $siteAdmins = Get-SPUser -Web $url -Limit All | Where-Object {$_.IsSiteAdmin}

    # Add each admin's details to the array
    foreach ($admin in $siteAdmins) {
        $myArray.Add([PSCustomObject]@{
            DisplayName = $admin.DisplayName
            LoginName   = $admin.LoginName
            Groups      = $admin.Groups -join ', '  # Join multiple groups into a single string
            URL         = $url
        })
    }
}

# Output to HTML with colorful CSS
$htmlOutput = $myArray | Select-Object DisplayName, LoginName, Groups, URL | ConvertTo-Html -Head "<style>
    table {border-collapse: collapse; width: 100%;}
    th, td {padding: 8px; text-align: left;}
    th {background-color: #4CAF50; color: white;}
    tr:nth-child(even) {background-color: #f2f2f2;}
    tr:hover {background-color: #ddd;}
</style>" -Body "<h2>Site Collection Administrators Report</h2>"

# Save HTML output to a file
$htmlOutput | Out-File "SiteCollectionAdminsReport.html"

# Output to CSV
$myArray | Export-Csv -Path "SiteCollectionAdminsReport.csv" -NoTypeInformation

Write-Host "HTML and CSV files have been generated."
