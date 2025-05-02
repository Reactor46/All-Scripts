    Function ConnectTo-Database{
    # Begin Connection Config
    $SqlServer = "LASPSHOST.FNBM.CORP"
    $Database = "SITEST"
    # Credentials for sa user
    $SQLCred = Import-Clixml -Path C:\Scripts\SIReport\Creds\SQLCredsSA.xml
    # End Connection Config

    # Begin Region Connection
    Import-Module C:\Scripts\SIReport\dbatools\dbatools.psm1
    Write-Host "Building connection string" -ForegroundColor Black -BackgroundColor White
    $Connection = Connect-DbaInstance -SqlInstance $SqlServer -Credential $SQLCred
    Write-Host "Opening connection to $($SqlServer)"
    # End Region Connection

    }
