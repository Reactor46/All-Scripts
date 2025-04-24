# Define the paths for the CSV and JSON files
$csvFilePath = 'C:\Users\silad\Documents\Powershell_stuff\PowerShellLearning\cities.csv'
$jsonFilePath = 'C:\Users\silad\Documents\Powershell_stuff\PowerShellLearning\cities.json'

# Import CSV data
$csvData = Import-Csv -Path $csvFilePath

# Convert CSV data to JSON format
$jsonData = $csvData | ConvertTo-Json

# Save JSON data to a file
$jsonData | Out-File -FilePath $jsonFilePath

Write-Host "Conversion from CSV to JSON complete. JSON file saved at: $jsonFilePath"
