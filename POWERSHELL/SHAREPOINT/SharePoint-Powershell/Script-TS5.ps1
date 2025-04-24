# This needs to be run as the farm admin account or another account that has full control on the Configuration database.

# Get the Config DB

Add-PSSnapin microsoft.sharepoint.powershell -ea SilentlyContinue

$conn = New-Object System.Data.SqlClient.SqlConnection

$cmd = New-Object System.Data.SqlClient.SqlCommand

$cmd2 = New-Object System.Data.SqlClient.SqlCommand

$configDb = Get-SPDatabase | ?{$_.TypeName -match "Configuration Database"}

$conn.ConnectionString = $configDb.DatabaseConnectionString

# Function to get row count for timer job history

Function GetRows

{

$conn.Open()

$cmd.connection = $conn

$cmd.CommandText = "select COUNT(*) from TimerJobHistory (nolock)"

$rows = $cmd.ExecuteReader()

if($rows.HasRows -eq $true)

{while($rows.Read())

{Write-host "Timer Job History Table contains: " $rows[0] " rows"}

“”}

$rows.Close()

$conn.Close()

}

# Function to truncate timer job history table

Function TruncateTable

{

$conn.Open()

$cmd2.connection = $conn

$cmd2.CommandText = "TRUNCATE table TimerJobHistory”

Write-Host -ForegroundColor Red "Truncating TimerJobHistory table on DB: " $configDb.Name

$cmd2.ExecuteReader()

$conn.Close()

}

GetRows

$answer = Read-Host "Would you like to Truncate the TimerJobHistory table? (all rows will be deleted) Enter Y or N "

If ($answer -ieq "Y")

{

# Truncate the table

TruncateTable

# Pause to let SQL do its thing

Start-Sleep -Seconds 10

# Get new row count. It should be close to 0.

GetRows

}

Else{Write-host -ForegroundColor green "You have chosen NOT to truncate the table. Nothing is changed."}

$conn.Close()