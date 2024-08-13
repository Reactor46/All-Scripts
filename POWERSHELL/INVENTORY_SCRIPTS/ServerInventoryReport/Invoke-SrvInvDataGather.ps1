# Begin Connection Config
    $SqlServer = "LASPSHOST.FNBM.CORP"
    $Database = "SITEST"
    # Credentials for sa user
    $cred = Import-Clixml -Path .\Creds\SQLCredsSA.xml
    # End Connection Config

    # Begin Region Connection
    Import-Module dbatools
    Write-Host "Building connection string" -ForegroundColor Black -BackgroundColor White
    $Connection = Connect-DbaInstance -SqlInstance $SqlServer -Credential $cred
    Write-Host "Opening connection to $($SqlServer)"
    # End Region Connection

    $FNBM = "$PSScriptRoot\RESULTS\FNBM.CORP\FNBM.Alive.txt"
    $PHX = "$PSScriptRoot\RESULTS\PHX.FNBM.CORP\PHX.Alive.txt"
    $TST = "$PSScriptRoot\RESULTS\CREDITONEAPP.TST\C1A.TST.Alive.txt"
    $BIZ = "$PSScriptRoot\RESULTS\CREDITONEAPP.BIZ\C1A.BIZ.Alive.txt"