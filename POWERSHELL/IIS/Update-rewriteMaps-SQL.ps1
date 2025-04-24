# Import the SQL Server module
Import-Module SqlServer

# Database connection parameters
$serverName = "localhost"
$databaseName = "RewriteDB"
$tableName = "RewriteTable"
$csvFilePath = "D:\RewriteMaps\Update-RewriteMaps.csv"

# SQL Server authentication
#$credential = Get-Credential

# Connect to the SQL Server
$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
$sqlConnection.ConnectionString = "Server=$serverName;Database=$databaseName;Integrated Security=true"
$sqlConnection.Open()

# Read CSV file and update entries in the database
$csvData = Import-Csv -Path $csvFilePath
foreach ($row in $csvData) {
    $originalUrl = $row.Source
    $newUrl = $row.Destination

    # Prepare SQL command to update data
    $sqlCommand = $sqlConnection.CreateCommand()
    $sqlCommand.CommandText = "UPDATE $tableName SET NewUrl = @NewUrl WHERE OriginalUrl = @OriginalUrl"

    # Prepare SQL parameters
    $sqlCommand.Parameters.AddWithValue("@OriginalUrl", $originalUrl)
    $sqlCommand.Parameters.AddWithValue("@NewUrl", $newUrl)

    # Execute the SQL command
    $sqlCommand.ExecuteNonQuery()
}

# Close the SQL connection
$sqlConnection.Close()

Write-Host "Data updated in SQL database."
