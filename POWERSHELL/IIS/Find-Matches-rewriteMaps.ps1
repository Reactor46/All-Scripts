# Path to the CSV file
$csvFilePath = "D:\VULS\Exported-rewriteMaps.csv"

# Import the CSV file
Write-Verbose "Importing CSV file from $csvFilePath"
$csvData = Import-Csv -Path $csvFilePath -Verbose

# Initialize an array to store matching pairs
$matchingPairs = @()

# Iterate through each row in Column FromURL
for ($i = 0; $i -lt $csvData.FromURL.Count; $i++) {
    $valueFromURL = $csvData.FromURL[$i]

    Write-Verbose "Processing row $($i + 1) in Column FromURL: $valueFromURL"

    # Iterate through each row in Column ToURL
    for ($j = 0; $j -lt $csvData.ToURL.Count; $j++) {
        $valueToURL = $csvData.ToURL[$j]

        Write-Verbose "Comparing with row $($j + 1) in Column ToURL: $valueToURL"

        # Check if the values in Column FromURL and Column ToURL match exactly
        if ($valueFromURL -eq $valueToURL) {
            Write-Verbose "Exact match found:"
            Write-Verbose "Row number of FromURL: $($i + 1)"
            Write-Verbose "Row number of ToURL: $($j + 1)"
            Write-Verbose "FromURL: $valueFromURL"
            Write-Verbose "ToURL: $valueToURL"

            $matchingPairs += @{
                RowNumberFromURL = $i + 1
                RowNumberToURL = $j + 1
                FromURL = $valueFromURL
                ToURL = $valueToURL
            }
        }
    }
}

# Output the matching pairs
if ($matchingPairs.Count -gt 0) {
    Write-Output "Matching pairs found:"
    $matchingPairs
} else {
    Write-Output "No matching pairs found."
}
