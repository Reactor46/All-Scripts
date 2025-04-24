# Import the CSV files
$CSV1 = Import-Csv -Path 'C:\LazyWinAdmin\InstalledKBs.csv'
$CSV2 = Import-Csv -Path 'C:\LazyWinAdmin\MatchMe.csv'


# Extract server names and installed KBs from CSV1
$Servers = $CSV1.ComputerName
$InstalledKBs = $CSV1.HotFixId

# Initialize an array to store missing KB patches
$MissingKBs = @()

# Iterate through each row in CSV2
foreach ($row in $CSV2) {
    $ServerName = $row.Endpoint
    $KB = $row.CheckID

    # Check if the server exists in CSV1
    if ($Servers -contains $ServerName) {
        # Check if the installed KB is found
        if (-not ($InstalledKBs -contains $KB)) {
            # Add the missing KB patch to the array
            $MissingKBs += [PSCustomObject]@{
                ServerName = $ServerName
                KB = $KB
            }
        }
    } else {
        Write-Output "Server $ServerName not found in InstalledKBs.csv."
        # Add your actions here for when the server is not found
    }
}

# Export the missing KB patches to a CSV file
$MissingKBs | Export-Csv -Path 'C:\LazyWinAdmin\Missing KB Patches.csv' -NoTypeInformation
