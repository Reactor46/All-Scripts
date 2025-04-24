# Path to the rewriteMaps.config file
$rewriteMapsConfigPath = "\\fbv-scrcd10-p01\C$\inetpub\wwwroot\ksprod-new-cd.ksnet.com\rewriteMaps.config"

# Read the contents of the rewriteMaps.config file
$rewriteMapsContent = Get-Content -Path $rewriteMapsConfigPath -Raw

# Extract all source and value pairs from the rewriteMaps.config file
$sourceValuePairs = [regex]::Matches($rewriteMapsContent, '<add\s+key="([^"]+)"\s+value="([^"]+)"')

# Initialize an array to store data
$rewriteMapsData = @()

# Iterate through each source and value pair
foreach ($pair in $sourceValuePairs) {
    $source = $pair.Groups[1].Value
    $value = $pair.Groups[2].Value

    # Create a custom object for each source and value pair
    $entry = [PSCustomObject]@{
        FromURL = $source
        ToURL = $value
    }
    $rewriteMapsData += $entry
}

# Export the data to a CSV file
$rewriteMapsData | Export-Csv -Path "D:\VULS\Exported-rewriteMaps.csv" -NoTypeInformation
