# List of servers to check
$Servers = @("fbv-scrcd10-p01", "fbv-scrcd10-p02", "fbv-scrcd10-p03", "fbv-scrcd10-p04")  # Add your server names here

# Iterate through each server
foreach ($Server in $Servers) {
    Write-Host "Checking $Server..."

    # Get the Web-Dyn-Compression feature status
    $FeatureStatus = Get-WindowsFeature -ComputerName $Server | Where-Object { $_.Name -eq "Web-Dyn-Compression" }

    # Check if the feature is enabled
    if ($FeatureStatus.Installed) {
        Write-Host "Web-Dyn-Compression is enabled on $Server."
    } else {
        Write-Host "Web-Dyn-Compression is not enabled on $Server."
    }

    Write-Host "-----------------------"
}
