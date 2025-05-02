Function Invoke-SQLCmd {    
    [cmdletbinding(
        DefaultParameterSetName = 'NoCred',
        SupportsShouldProcess = $True,
        ConfirmImpact = 'Low'
    )]
    Param (
        [parameter()]
        [string]$Computername = $SqlServer,
        
        [parameter()]
        [string]$Database = $Database,    
        
        [parameter()]
        [string]$TSQL,

        [parameter()]
        [int]$ConnectionTimeout = 30,

        [parameter()]
        [int]$QueryTimeout = 120,

        [parameter()]
        [System.Collections.ICollection]$SQLParameter,

        [parameter(ParameterSetName='Cred')]
        [Alias('RunAs')]        
        [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty,

        [parameter()]
        [ValidateSet('Query','NonQuery')]
        [string]$CommandType = 'Query'
    )
    If ($PSBoundParameters.ContainsKey('Debug')) {
        $DebugPreference = 'Continue'
    }
    $PSBoundParameters.GetEnumerator() | ForEach {
        Write-Debug $_
    }
    #region Make Connection
    Write-Verbose "Building connection string"
    $Connection = New-Object System.Data.SqlClient.SQLConnection 
    If ($PSBoundParameters.ContainsKey('Verbose')) {
        $Handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] {
            Param($sender, $event) 
            Write-Verbose $event.Message -Verbose
        }
        $Connection.add_InfoMessage($Handler)
        $Connection.FireInfoMessageEventOnUserErrors=$True  
    }
    Switch ($PSCmdlet.ParameterSetName) {
        'Cred' {
            $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $Computername,
                                                                                        $Database,$Credential.Username,
                                                                                        $Credential.GetNetworkCredential().password,$ConnectionTimeout   
            Remove-Variable Credential
        }
        'NoCred' {
            $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $Computername,$Database,$ConnectionTimeout                 
        }
    }   
    $Connection.ConnectionString = $ConnectionString
    Write-Verbose "Opening connection to $($Computername)"
    $Connection.Open()
    #endregion Make Connection

    #region Initiate Query
    Write-Verbose "Initiating query -> $Tsql"
    $Command = New-Object system.Data.SqlClient.SqlCommand($Tsql,$Connection)
    If ($PSBoundParameters.ContainsKey('SQLParameter')) {
        $SqlParameter.GetEnumerator() | ForEach {
            Write-Verbose "Adding SQL Parameter: $($_.Key) with Value: $($_.Value)"
            If ($_.Value -ne $null) { 
                [void]$Command.Parameters.AddWithValue($_.Key, $_.Value) 
            }
            Else { 
                [void]$Command.Parameters.AddWithValue($_.Key, [DBNull]::Value) 
            }
        }
    }
    $Command.CommandTimeout = $QueryTimeout
    If ($PSCmdlet.ShouldProcess("Computername: $($Computername) - Database: $($Database)",'Run TSQL operation')) {
        Switch ($CommandType) {
            'Query' {
                Write-Verbose "Performing Query operation"
                $DataSet = New-Object system.Data.DataSet
                $DataAdapter = New-Object system.Data.SqlClient.SqlDataAdapter($Command)
                [void]$DataAdapter.fill($DataSet)
                $DataSet.Tables
            }
            'NonQuery' {
                Write-Verbose "Performing Non-Query operation"
                [void]$Command.ExecuteNonQuery()
            }
        }
    }
    #endregion Initiate Query    

    #region Close connection
    Write-Verbose "Closing connection"
    $Connection.Close()        
    #endregion Close connection
}