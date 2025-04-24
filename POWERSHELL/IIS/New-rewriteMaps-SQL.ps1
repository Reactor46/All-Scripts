# Import the SQL Server module
Import-Module SqlServer

# Database connection parameters
$serverName = "localhost"
$databaseName = "RewriteDB"
$tableName = "RewriteTable"
$csvFilePath = "D:\RewriteMaps\New-RewriteMaps.csv"

# SQL Server authentication
#$credential = Get-Credential

# Connect to the SQL Server
$sqlConnection = New-Object System.Data.SqlClient.SqlConnection
$sqlConnection.ConnectionString = "Server=$serverName;Database=$databaseName;Integrated Security=True"
$sqlConnection.Open()

# Prepare SQL command to insert data
$sqlCommand = $sqlConnection.CreateCommand()
$sqlCommand.CommandText = "INSERT INTO $tableName (OriginalUrl, NewUrl) VALUES (@OriginalUrl, @NewUrl)"

# Prepare SQL parameters
$sqlCommand.Parameters.Add("@OriginalUrl", [System.Data.SqlDbType]::VarChar, 255)
$sqlCommand.Parameters.Add("@NewUrl", [System.Data.SqlDbType]::VarChar, 255)

# Read CSV file and insert data into the database
$csvData = Import-Csv -Path $csvFilePath
foreach ($row in $csvData) {
    $originalUrl = $row.Source
    $newUrl = $row.Destination

    $sqlCommand.Parameters["@OriginalUrl"].Value = $originalUrl
    $sqlCommand.Parameters["@NewUrl"].Value = $newUrl

    $sqlCommand.ExecuteNonQuery()
}

# Close the SQL connection
$sqlConnection.Close()

Write-Host "Data imported from CSV to SQL database."
