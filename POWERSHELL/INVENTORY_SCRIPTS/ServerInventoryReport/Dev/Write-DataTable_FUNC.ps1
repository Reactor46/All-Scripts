Function Write-DataTable {
    [CmdletBinding()]
    param(
    [Parameter(Position=0, Mandatory=$true)] 
    [string]$Computername,
    [Parameter(Position=1, Mandatory=$true)] 
    [string]$Database,
    [Parameter(Position=2, Mandatory=$true)] 
    [string]$TableName,
    [Parameter(Position=3, Mandatory=$true)] 
    $Data,
	[Parameter(Position=4)] 
    [string]$Username,
    [Parameter(Position=5)] 
    [string]$Password,	
	[Parameter(Position=6)] 
    [Int32]$BatchSize=50000,
    [Parameter(Position=7)] 
    [Int32]$QueryTimeout=0,
    [Parameter(Position=8)] 
    [Int32]$ConnectionTimeout=15
    )
    
    $SQLConnection = New-Object System.Data.SqlClient.SQLConnection

    If ($Username) { 
        $ConnectionString = "Server={0};Database={1};UserID={2};Password{3};Trusted_Connection=False;Connect Timeout={4}" -f $Computername,$Database,$Username,$Password,$ConnectionTimeout 
    }
    Else { 
        $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $Computername,$Database,$ConnectionTimeout 
    }

    $SQLConnection.ConnectionString = $ConnectionString

    Try {
        $SQLConnection.Open()
        $bulkCopy = New-Object Data.SqlClient.SqlBulkCopy -ArgumentList $SQLConnection, ([System.Data.SqlClient.SqlBulkCopyOptions]::TableLock),$Null
        $bulkCopy.DestinationTableName = $tableName
        $bulkCopy.BatchSize = $BatchSize
        $bulkCopy.BulkCopyTimeout = $QueryTimeOut
        $bulkCopy.WriteToServer($Data)        
    }
    Catch {
        Write-Error "$($TableName): $($_)"
    }
    Finally {
        $SQLConnection.Close()
    }
}
