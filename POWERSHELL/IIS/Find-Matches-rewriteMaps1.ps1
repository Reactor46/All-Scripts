# Path to the CSV file
$csvFilePath = "D:\VULS\Exported-rewriteMaps.csv"

# Import the CSV file
$csvData = Import-Csv -Path $csvFilePath -Verbose

# Initialize an array to store matching pairs
$matchingPairs = @()

# Iterate through each row in Column A
for ($i = 0; $i -lt $csvData.FromURL.Count; $i++) {
    $valueA = $csvData.FromURL[$i]

    # Iterate through each row in Column B
    for ($j = 0; $j -lt $csvData.ToURL.Count; $j++) {
        $valueB = $csvData.ToURL[$j]

        # Check if the values in Column FromURL and Column ToURL match exactly
        if ($valueA -eq $valueB) {
            $matchingPairs += @{
                RowNumberFromURL = $i + 1
                RowNumberToURL = $j + 1
                FromURL = $valueA
                ToURL = $valueB
            }
        }
    }
}

# Output the matching pairs
if ($matchingPairs.Count -gt 0) {
    Write-Output "Matching pairs found:" -Verbose
    $matchingPairs
} else {
    Write-Output "No matching pairs found." -Verbose
}
