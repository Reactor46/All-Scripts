    Function ConnectTo-Database{
    # Begin Connection Config
    $SQLServer = "LASPSHOST.Contoso.CORP"
    $SQLDB = "SITEST"
    # Credentials for sa user
    $SQLCred = Import-Clixml -Path .\Creds\SQLCredsSA.xml
    # End Connection Config

    # Begin Region Connection
    Import-Module dbatools
    Write-Host "Building connection string" -ForegroundColor Black -BackgroundColor White
    $Connection = Connect-DbaInstance -SqlInstance $SqlServer -Credential $SQLCred
    Write-Host "Opening connection to $($SqlServer)"
    # End Region Connection

    }
