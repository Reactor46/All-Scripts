### Created By Martin ###

# Define the input file and output file paths
$cveFilePath = "C:\golive\CVE.txt"
$outputFilePath = "C:\golive\CVE_OUTPUT.html"

# Read the CVE list from the file
$cveList = Get-Content -Path $cveFilePath -Raw
$cveArray = $cveList -split ",\s*"

# Initialize an array to hold the HTML rows
$htmlRows = @()

# Loop through each CVE in the list
foreach ($cve in $cveArray) {
    # Trim any leading/trailing whitespace
    $cve = $cve.Trim()

    # Construct the URL for the CVE
    $url = "https://cve.mitre.org/cgi-bin/cvename.cgi?name=$cve"

    # Fetch the web page content
    try 
    {
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing
        $content = $response.Content

        # Extract the description using a regular expression
        if ($content -match '<th colspan="2">Description<\/th>\s*<\/tr>\s*<tr>\s*<td colspan="2">\s*(.*?)\s*<\/td>') 
        {
            $description = $matches[1].Trim()
            $description = $description -replace '\s+', ' ' # Normalize whitespace
        } 
        else 
        {
            $description = "Description not found"
        }

        # Extract the AssigningCNA using a regular expression
        if ($content -match '<th colspan="2">Assigning CNA<\/th>\s*<\/tr>\s*<tr>\s*<td colspan="2">\s*(.*?)\s*<\/td>') 
        {
            $cna = $matches[1].Trim()
            $cna = $cna -replace '\s+', ' ' # Normalize whitespace
        } 
        else 
        {
            $cna = "Assigning CNA not found"
        }

        # Add the CVE and description to the HTML rows
        $htmlRows += "<tr><td>$cve</td><td>$description</td><td>$cna</td></tr>"
    } 
    catch 
    {
        # Handle any errors during the web request
        $htmlRows += "<tr><td>$cve</td><td>WEB REQUEST ERROR</td></tr>"
    }
}

# Construct the HTML content
$htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CVE Descriptions v1.1</title>
    <style>
        table 
        {
            width: 100%;
            border-collapse: collapse;
        }
        th, td {
            border: 1px solid black;
            padding: 8px;
            text-align: left;
        }
        th 
        {
            background-color: #f2f2f2;
        }
    </style>
</head>
<body>
    <table>
        <thead>
            <tr>
                <th>CVE</th>
                <th>Description</th>
                <th>Manufacturer</th>
            </tr>
        </thead>
        <tbody>
            $(foreach ($row in $htmlRows) { $row })
        </tbody>
    </table>
</body>
</html>
"@

# Write the HTML content to the output file
$htmlContent | Set-Content -Path $outputFilePath
