# Import the CSV files
$CSV1 = Import-Csv -Path 'C:\LazyWinAdmin\InstalledKBs.csv'
$CSV2 = Import-Csv -Path 'C:\LazyWinAdmin\MatchMe.csv'

# Extract server names and installed KBs from CSV1
$Servers = $CSV1.ComputerName
$InstalledKBs = $CSV1.HotFixId

# Iterate through each server in CSV2
foreach ($row in $CSV2) {
    $ServerName = $row.Endpoint
    $KB = $row.CheckID

    # Check if the server exists in CSV1
    if ($Servers -contains $ServerName) {
        # Check if the installed KB is found
        if ($InstalledKBs -contains $KB) {
            Write-Output "Server $ServerName has KB $KB installed."
            # Add your actions here for when a match is found
        } else {
            Write-Output "Server $ServerName does not have KB $KB installed."
            # Add your actions here for when a match is not found
        }
    } else {
        Write-Output "Server $ServerName not found in InstalledKBs.csv."
        # Add your actions here for when the server is not found
    }
}
