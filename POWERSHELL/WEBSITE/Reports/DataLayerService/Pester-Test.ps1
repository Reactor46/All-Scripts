# Invoke pester and store the results so that they can be referenced multiple times
$results = Invoke-Pester -show none -passthru

# Format cells where the value is greater than 0
$params = @{
    ScriptBlock = {$args[0] -gt 0}
}

# Create a summary table and add spaces into the column names for asthetics, adding colours for passed and failed tests
$summaryTable = $results | Select-Object `
     @{Name="Passed Count";Expression={$_.PassedCount}},
     @{Name="Failed Count";Expression={$_.FailedCount}},
     @{Name="Skipped Count";Expression={$_.SkippedCount}},
     @{Name="Pending Count";Expression={$_.PendingCount}},
     @{Name="Total Count";Expression={$_.TotalCount}} | New-HtmlTable |
    Add-HTMLTableColor -Argument "Failed" -Column "Failed Count" -AttrValue "background-color:#ffb3b3;" @params |
    Add-HTMLTableColor -Argument "Passed" -Column "Passed Count" -AttrValue "background-color:#c6ffb3;" @params

#Compose the html adding headers
$HTML = New-HTMLHead
$HTML += "<h3>Post Build Test Summary</h3>"
$HTML += $summaryTable
$HTML += "<h3>Post Build Test Results</h3>"

# Create a seperate table for each describe block
foreach ($section in ($results |Select-Object -ExpandProperty TestResult | Select-Object Describe -Unique)) {
    # Add a title based on the descibe block
    $HTML += ("<h4>{0}</h4>" -f $section.Describe)
    $HTML += $results | Select-Object -ExpandProperty TestResult | Where-Object -FilterScript { $_.Describe -eq $section.Describe } |
        Select-Object Context, Name, Result |
        New-HtmlTable |
            Add-HTMLTableColor -Argument "Failed" -Column "Result" -AttrValue "background-color:#ffb3b3;" |
            Add-HTMLTableColor -Argument "Passed" -Column "Result" -AttrValue "background-color:#c6ffb3;"
}

$HTML += "" | Close-HTML

#test it out
set-content C:\inetpub\wwwroot\Pester-Test.html $HTML
& 'C:\Program Files\Internet Explorer\iexplore.exe' C:\inetpub\wwwroot\Pester-Test.html