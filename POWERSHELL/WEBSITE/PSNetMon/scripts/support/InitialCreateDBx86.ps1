#######################################################################
#
#  Intial Database Creation in Powershellx86 
#
#  Run the script in x86 version of Powershell to create the DB used for
#     reporting features
#
#######################################################################

#######################################################################
#  Functions of creating db 
#######################################################################
Function Check-Path($Db)
{
 If(!(Test-Path -path (Split-Path -path $Db -parent)))
   { 
     Throw "$(Split-Path -path $Db -parent) Does not Exist" 
   }
  ELSE
  { 
   If(Test-Path -Path $Db)
     {
      Throw "$db already exists"
     }
  }
} #End Check-Path

Function Create-DataBase($Db)
{
 $application = New-Object -ComObject Access.Application
 $application.NewCurrentDataBase($Db,10)
 $application.CloseCurrentDataBase()
 $application.Quit()
} #End Create-DataBase

Function Invoke-ADOCommand($Db, $Command)
{
 $connection = New-Object -ComObject ADODB.Connection
 $connection.Open("Provider= Microsoft.Jet.OLEDB.4.0;Data Source=$Db" )
 $connection.Execute($command)
 $connection.Close()
} #End Invoke-ADOCommand

Function Invoke-ADOCommand($Db, $Command2)
{
 $connection = New-Object -ComObject ADODB.Connection
 $connection.Open("Provider= Microsoft.Jet.OLEDB.4.0;Data Source=$Db" )
 $connection.Execute($command2)
 $connection.Close()
} #End Invoke-ADOCommand

#######################################################################
#  Powershell variables
#######################################################################

$Db = "C:\TECH\Scripts\PowerShell\PSNetMon\db\ReportingDB.mdb"
$table = "Uptime"
$Fields = "F1Date Date, F2Resource Text, F3Status Text"
$command = "Create Table $table `($fields`)"
$table2 = "PortStatus"
$Fields2 = "F1Date Date, F2Ports Text, F3Status Text"
$command2 = "Create Table $table2 `($fields2`)"
$table3 = "ServiceStatus"
$Fields3 = "F1Date Date, F2Service Text, F3Status Text"
$command3 = "Create Table $table3 `($fields3`)"

#######################################################################
#  Process commands for script
#######################################################################

Check-Path -db $Db
Create-DataBase -db $Db
Invoke-ADOCommand -db $Db -command $command
Invoke-ADOCommand -db $Db -command $command2
Invoke-ADOCommand -db $Db -command $command3