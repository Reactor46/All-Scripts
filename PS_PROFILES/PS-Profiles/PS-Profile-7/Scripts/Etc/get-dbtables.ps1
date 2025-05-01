Function Check-Path($Db)
{
 If(!(Test-Path -path (Split-Path -path $Db -parent)))
   { 
     Throw "$(Split-Path -path $Db -parent) Does not Exist" 
   }
  ELSE
  { 
   If (!(Test-Path -Path $Db))
     {
      Throw "$db does not exist"
     }
  }
} #End Check-Path

Function New-Line (
                  $strIN,
                  $char = "=",
                  $sColor = "Yellow",
                  $uColor = "darkYellow",
                  [switch]$help
                 )
{
 if($help)
  {
    $local:helpText = `
@"
     New-Line accepts inputs: -strIN for input string and -char for seperator
     -sColor for the string color, and -uColor for the underline color. Only 
     the -strIn is required. The others have the following default values:
     -char: =, -sColor: Yellow, -uColor: darkYellow
     Example:
     New-Line -strIN "Hello world"
     New-Line -strIn "Morgen welt" -char "-" -sColor "blue" -uColor "yellow"
     New-Line -help
"@
   $local:helpText
   break
  } #end New-Line help
  
 $strLine= $char * $strIn.length
 Write-Host -ForegroundColor $sColor $strIN 
 Write-Host -ForegroundColor $uColor $strLine
} #end New-Line function

Function Get-DataType ($enum)
{
 [enum]::Parse("system.data.oledb.oledbtype",$enum)
} #end Get-DataType

Function Get-DataBaseSchema ($Db)
{
 New-Line("Obtaining Schema for $Db")
 $Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
 $DataSource = "Data Source ="+$Db
 $Connection = New-Object Data.OleDb.OleDbConnection("$Provider;$DataSource")
 $Connection.open()
 $Tables = $Connection.GetSchema("Tables") | 
   ForEach-Object { 
      if($_.Table_Type -eq "TABLE") 
        { $_.Table_Name } 
     }
   ForEach($table in $Tables)
    { 
     New-Line("Table Name: " + $table)
     $SchemaColumn = $Connection.GetSchema("Columns")
     $schemaColumn | 
     foreach-Object {
        If($_.Table_Name -eq $Table) 
         { "$($_.COLUMN_Name) $(Get-DataType($_.DATA_TYPE))" } 
      } #end Foreach-Object
    } #end Foreach table
$Connection.Close()
}

# *** Entry Point to Script ***

$Db = "\\Misfs1\svcreq$\Request for Services.mdb"
Check-Path($Db)
Get-DataBaseSchema($Db)