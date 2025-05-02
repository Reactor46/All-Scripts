## First type of connection
$SQLCred = New-Object System.Data.SqlClient.SqlCredential("username","password")
$Connection = New-Object System.Data.SqlClient.SqlConnection("SI-TEST",$SQLCred)
$Connection.Open()

## Second type of connection
$instance = "myInstance"
$userId = "myUserId"
$password = "myPassword"

$connectionString = "Data Source=$instance;Integrated Security=SSPI;Initial Catalog=master; User Id=$userId; Password=$password;"

$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString
$connection.Open()

## Third type of connection
$cred = New-Object System.Data.SqlClient.credential($userId, $password)
$connectionString = "Data Source=$Instance;Integrated Security=SSPI;Initial Catalog=master; User Id = $($cred.username); Password = $($cred.GetNetworkCredential().password);"
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString
$connection.Open()

## Secure Password Encrypting (Get-Creds)
$Account = "MyDomain\MyAccount"
$AccountPassword = Read-Host -AsSecureString
$DatabaseCredentials = New-Object System.Management.Automation.PSCredential($Account,$AccountPassword)


