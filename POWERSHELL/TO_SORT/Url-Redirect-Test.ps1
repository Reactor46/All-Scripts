$CsvImport = Import-Csv -Path D:\RewriteMaps\KSC-Current-Redirects.csv
# Function to get color based on StatusCode
function Get-StatusColor {
    param (
        [int]$statusCode
    )
    
    if ($statusCode -eq 301 -or $statusCode -eq 302) {
        return "green"
    } elseif ($statusCode -ge 303 -and $statusCode -lt 400) {
        return "yellow"
    } elseif ($statusCode -ge 400) {
        return "red"
    } else {
        return "black"
    }
}

# Create HTML content
$html = @"
<html>
<head>
    <style>
        table {
            width: 100%;
            border-collapse: collapse;
            background-color: lightblue;
        }
        table, th, td {
            border: 1px solid black;
        }
        th, td {
            padding: 10px;
            text-align: left;
        }
        th {
            cursor: pointer;
        }
    </style>
    <script>
        function sortTable(n) {
            var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
            table = document.getElementById("statusTable");
            switching = true;
            dir = "asc"; 
            while (switching) {
                switching = false;
                rows = table.rows;
                for (i = 1; i < (rows.length - 1); i++) {
                    shouldSwitch = false;
                    x = rows[i].getElementsByTagName("TD")[n];
                    y = rows[i+1].getElementsByTagName("TD")[n];
                    if (dir == "asc") {
                        if (x.innerHTML.toLowerCase() > y.innerHTML.toLowerCase()) {
                            shouldSwitch = true;
                            break;
                        }
                    } else if (dir == "desc") {
                        if (x.innerHTML.toLowerCase() < y.innerHTML.toLowerCase()) {
                            shouldSwitch = true;
                            break;
                        }
                    }
                }
                if (shouldSwitch) {
                    rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                    switching = true;
                    switchcount ++; 
                } else {
                    if (switchcount == 0 && dir == "asc") {
                        dir = "desc";
                        switching = true;
                    }
                }
            }
        }
    </script>
</head>
<body>
    <table id="statusTable">
        <tr>
            <th onclick="sortTable(0)">No.</th>
            <th onclick="sortTable(1)">Original</th>
            <th onclick="sortTable(2)">Target</th>
            <th onclick="sortTable(3)">StatusCode</th>
        </tr>
"@

# Counter for numbering rows
$rowNumber = 1
#$Links = Get-Content -Path D:\RewriteMaps\KSC-Redirect-Test.txt
$Links = $CsvImport.'Test URL'

Foreach($link in $Links){

$results = Test-Redirects -Url $link -ErrorAction SilentlyContinue 



foreach ($entry in $results) {
    $color = Get-StatusColor -statusCode $entry.StatusCode
    $html += "<tr style='color:$color'>"
    $html += "<td>$rowNumber</td>"
    $html += "<td>$($entry.Original)</td>"
    $html += "<td>$($entry.Target)</td>"
    $html += "<td>$($entry.StatusCode)</td>"
    $html += "</tr>"
    $rowNumber++
}

$html += @"
    </table>
</body>
</html>
"@


}

# Output HTML to file
$html | Out-File -FilePath "output.html" -Encoding utf8
# Optional: Open the HTML file in the default browser
Start-Process "output.html"
