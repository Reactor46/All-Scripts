<#
.SYNOPSIS
    this script will get all Spiceworks Clients (Desktop, Laptop, Servers) which are running Windows
    and get the Monitor via WMI WMIMonitorID. Then adds this Monitor to Spiceworks Database as a 
    Manual Device.

    I run this Script LIVE, no Backup no Shutdown of Spiceworks!!! But I would recommend to
    do it OFFLINE! ;-)
    
    THIS SCRIPT IS PROVIDED AS IS, NO WARRANTY THAT IT WILL WORK.
    YOU RUN IT AT OWN RISK!
    DO a Backup of the Database, do not Run this when Spiceworks runs!

.NOTES
    File Name    : Get_Monitors_to_Spiceworks.ps1
    Author       : Thomas Voß
    Parameter    : DryRun
                   - Set to NO to do it real!

    Version      : 0.3
    CreateDate   : 18.02.2014
    UpdateDate   : 26.02.2014
    Prerequisites: PowerShell 3.0, SQL-Lite ODBC Driver, Spiceworks DB, Active Directory (RSAT) or REMOTING
    Links        : http://www.ch-werner.de/sqliteodbc/ ( SQL-ODBC-Driver )
                 : http://www.microsoft.com/en-us/download/details.aspx?id=34595 ( Windows 7 / Windows 2008 / Windows 2008 R2 )
#>

# Parameters
Param (
    [String]
    $DryRun       = "YES"
)

# Connection String to your Spiceworks Database
$Spiceworks_Connection_String = "Driver=SQLite3 ODBC Driver;Database=C:\Spiceworks\db\spiceworks_prod.db;"

# Maybe You have to change that!!!!!!!
$Site_id  = 9


# Maybe you have to ADD some MORE!
$MyManufacturerHash = @{ 
    "AAC" =	"AcerView";
    "ACR" = "Acer";
    "AOC" = "AOC";
    "AIC" = "AG Neovo";
    "APP" = "Apple Computer";
    "AST" = "AST Research";
    "AUO" = "Asus";
    "BNQ" = "BenQ";
    "CMO" = "Acer";
    "CPL" = "Compal";
    "CPQ" = "Compaq";
    "CPT" = "Chunghwa Pciture Tubes, Ltd.";
    "CTX" = "CTX";
    "DEC" = "DEC";
    "DEL" = "Dell";
    "DPC" = "Delta";
    "DWE" = "Daewoo";
    "EIZ" = "EIZO";
    "ELS" = "ELSA";
    "ENC" = "EIZO";
    "EPI" = "Envision";
    "FCM" = "Funai";
    "FUJ" = "Fujitsu";
    "FUS" = "Fujitsu-Siemens";
    "GSM" = "LG Electronics";
    "GWY" = "Gateway 2000";
    "HEI" = "Hyundai";
    "HIT" = "Hyundai";
    "HSL" = "Hansol";
    "HTC" = "Hitachi/Nissei";
    "HWP" = "Hewlett-Packard";
    "IBM" = "IBM";
    "ICL" = "Fujitsu ICL";
    "IVM" = "Iiyama";
    "KDS" = "Korea Data Systems";
    "LEN" = "Lenovo";
    "LGD" = "Asus";
    "LPL" = "Fujitsu";
    "MAX" = "Belinea"; 
    "MEI" = "Panasonic";
    "MEL" = "Mitsubishi Electronics";
    "MS_" = "Panasonic";
    "NAN" = "Nanao";
    "NEC" = "NEC";
    "NOK" = "Nokia Data";
    "NVD" = "Fujitsu";
    "OPT" = "Optoma";
    "PHL" = "Philips";
    "REL" = "Relisys";
    "SAN" = "Samsung";
    "SAM" = "Samsung";
    "SBI" = "Smarttech";
    "SGI" = "SGI";
    "SNY" = "Sony";
    "SRC" = "Shamrock";
    "SUN" = "Sun Microsystems";
    "SEC" = "Hewlett-Packard";
    "TAT" = "Tatung";
    "TOS" = "Toshiba";
    "TSB" = "Toshiba";
    "VSC" = "ViewSonic";
    "ZCM" = "Zenith";
    "UNK" = "Unknown";
    "_YV" = "Fujitsu";
}

# Warn the USER
If ( $DryRun.ToLower() -eq "yes" ) {
    echo "!This will be a DryRun, NOTHING WILL BE CHANGED."
} else {
    echo "!ATTENTION! Will now change the Spiceworks Database!"
}


###########################################################################################
#
# Function Get-ODBC-Data
#
###########################################################################################
# This function gets-OBDC-Data
# Parameter: MyQuery      = "SQL Statement
#          : MyConnection = "Connection String to DB"
# Sample:
# $connection = "Driver=SQLite3 ODBC Driver;Database=C:\Install\db\spiceworks_prod.db;"
# $query      = "select * from users"
# $result     = Get-ODBC-Data -MyQuery $query -MyConnection $connection
#
###########################################################################################

function Get-ODBC-Data {

    # Parameters
    param(
        [string]
        $MyQuery      = $( throw 'Query is required.' ),
        [string]
        $MyConnection = $( throw 'Connection is required.' )
    )

    # Create new Object
    $conn = New-Object System.Data.Odbc.OdbcConnection
    # Set Connection String
    $conn.connectionstring = $MyConnection
    # Open Connection
    $conn.open()
    # Create Command Object
    $cmd  = New-object System.Data.Odbc.OdbcCommand( $MyQuery, $conn )
    # Create Dataset Object
    $ds   = New-Object system.Data.DataSet
    # Fill that Dataset Object
    ( New-Object system.Data.odbc.odbcDataAdapter( $cmd ) ).fill( $ds ) | out-null
    # Close Connection
    $conn.close()
    # Return Data
    $ds.Tables[0]
}

###########################################################################################
#
# Function Set-ODBC-Data
#
###########################################################################################
# This function sets-OBDC-Data
# Parameter: MyQuery      = "SQL Statement
#          : MyConnection = "Connection String to DB"
# Sample:
# $connection = "Driver=SQLite3 ODBC Driver;Database=C:\Install\db\spiceworks_prod.db;"
# $query      = "Update users set email='test@test.de'"
# $result     = Set-ODBC-Data -MyQuery $query -MyConnection $connection
#
###########################################################################################


function Set-ODBC-Data {
    # Parameter
    param(
        [string]
        $MyQuery      = $( throw 'query is required.' ),
        [string]
        $MyConnection = $( throw 'Connection is required.' )
    )

    # Create ODBC Object
    $conn = New-Object System.Data.Odbc.OdbcConnection
    # Set Connection String
    $conn.connectionstring = $MyConnection
    # Create Command 
    $cmd  = new-object System.Data.Odbc.OdbcCommand( $MyQuery, $conn )
    # Open Connection
    $conn.open()
    # Execute 
    $cmd.ExecuteNonQuery()
    # Close Connection
    $conn.close()
}

# Credits goes to: Reinhard K
# http://social.technet.microsoft.com/profile/reinhard%20k/?ws=usercard-mini

$HKLM = 2147483650

function Enum-RegKey( [String]$computerName=$env:ComputerName, [Int64]$hive=$HKLM, [String]$subKeyName){
 $([wmiClass]"\\$computerName\root\default:StdRegProv").EnumKey( $hive, $subKeyName).sNames
}

function Get-RegBinaryValue( [String]$computerName=$env:ComputerName, [Int64]$hive=$HKLM, [String]$subKeyName, [String]$valueName ){
 $([wmiClass]"\\$computerName\root\default:StdRegProv").GetBinaryValue( $hive, $subKeyName, $valueName).uValue
}

function Get-EDID
{
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeLine = $true,
            ValueFromPipeLineByPropertyName = $true)]
        [Alias('Name','Computer','CName','cn')]
        [String[]]
        $ComputerName
    )
   
    BEGIN
    {
        $arr = @()
        $regKeyDisplay = "System\CurrentCOntrolSet\Enum\Display"
    } #End BEGIN
   
    PROCESS
    {
        $arr += $computerName
    } # End PROCESS
   
    END
    {
        foreach( $c in $arr )
        {
              Enum-Regkey -ComputerName $c -Hive $HKLM -Subkey $regKeyDisplay | %{
                  $prodID = $_
                  $prodID | %{
                      Enum-RegKey -ComputerName $c -Hive $HKLM -SubKey "$regKeyDisplay\$prodID" | %{
                          if( @(Enum-Regkey -ComputerName $c -Hive $HKLM -SubKey "$regKeyDisplay\$prodID\$_") -contains 'Control' )
                          {
                              $sn, $mdl = '',''
                              $edid = Get-RegBinaryValue -computerName $c -hive $HKLM -subKeyName "$regKeyDisplay\$prodID\$_\Device Parameters" -valueName EDID
                              if( $edid -ne $null -and $edid.Count -ge 0x7D )
                              {
                                  $manWeek, $manYear, $dtd1, $dtd2, $dtd3, $dtd4 = $edid[0x10], $($edid[0x11] + 1990), $edid[0x36 .. 0x47], $edid[0x48 .. 0x59], $edid[0x5A .. 0x6B], $edid[0x6C .. 0x7D]
                                  $edid[0x36 .. 0x47], $edid[0x48 .. 0x59], $edid[0x5A .. 0x6B], $edid[0x6C .. 0x7D] | %{
                                      $dtd = $_
                                      $tagIdentifier = $_[0..4]
                                      $tag = $tagIdentifier[3]
                                      $tagIdentifierTest = $tagIdentifier[0..2+4]
                                      if( @(diff $tagIdentifierTest @(0,0,0,0)).Count -eq 0)
                                      {
                                          switch ($tag)
                                          {
                                              255 # 0xFF, SerialNumber
                                              {
                                                  # $sn = ''
                                                  for( $i = 5; $i -le 15; $i++)
                                                  {
                                                      if( $dtd[$i] -eq 0x0A)
                                                      {
                                                          break
                                                      }
                                                      else
                                                      {
                                                          $sn += [Char]($dtd[$i])
                                                      }
                                                  }
                                              } # End 0xFF
                                              252 # 0xFC, monitor name descriptor
                                              {
                                                  # $mdl = ''
                                                  for( $i = 5; $i -le 15; $i++)
                                                  {
                                                      if( $dtd[$i] -eq 0x0A)
                                                      {
                                                          break
                                                      }
                                                      else
                                                      {
                                                          $mdl += [Char]($dtd[$i])
                                                      }
                                                  }
                                              } # End 0xFC
                                          } #End switch
                                      } #End If
                                  } # End ForEach-Object
                                  new-object PSObject -Prop @{ UserFriendlyName = $mdl; ManufacturerName = $prodID; SerialNumberID = $sn; ProductCodeID = $mdl; WeekOfManufacture = $manWeek; YearOfManufacture = $manYear }
                              } # End If  
                          } #End If
                      } # End ForEach-Object
                  } # End ForEach-Object
              } # End ForEach-Object
          } # End foreach

  } # End End
} # End Function

# Get All Devices which are Laptop/Desktop/Server and have OS like Windows ;-) from Spiceworks
$SpiceWorks_Devices = Get-ODBC-Data -MyQuery "select name, server_name,id, primary_owner_name, user_id, device_type from devices where manually_added='f' and ( device_type='Desktop' or device_type='Laptop' or device_type='Server') and operating_system like '%windows%'" -MyConnection $Spiceworks_Connection_String

# Now iterate through the Devices
foreach ( $Computer in $SpiceWorks_Devices ) {

    echo "Searching on: $( $Computer.server_name ) - $( $Computer.name )"
    $All     = Get-WmiObject -Namespace root\wmi -Class WmiMonitorID -ComputerName $Computer.server_name -ErrorAction SilentlyContinue
    $EDIDGet = 0

    # Nothing there we need to check with get-EDID
    if ( ! $All ) {
        echo "Searching on: $( $Computer.server_name ) - $( $Computer.name ) - Alternativ Method"
        $EDIDGet = 1
        $All = get-EDID -ComputerName $Computer.server_name
    }

    foreach ($monitor in $All)
    {
        $serial       = ""
        $product      = ""
        $manu         = ""
        $userfriendly = ""
        $description  = ""
        $Monitorname  = ""
        $server_name  = ""

        if ( $EDIDGet -eq 1 ) {
            $serial       = $monitor.SerialNumberID
            $product      = $monitor.ProductCodeID
            $manu         = $monitor.ManufacturerName.Substring(0,3)
            $userfriendly = $monitor.UserFriendlyName
        } else {
            $monitor.SerialNumberID   | foreach { if ( $_ -gt 0 ) { $serial       += [char]$_ } }
            $monitor.ProductCodeID    | foreach { if ( $_ -gt 0 ) { $product      += [char]$_ } }
            $monitor.ManufacturerName | foreach { if ( $_ -gt 0 ) { $manu         += [char]$_ } }
            $monitor.UserFriendlyName | foreach { if ( $_ -gt 0 ) { $userfriendly += [char]$_ } }
        }
        # Some checks if there are empty values
        if ( $serial -eq "" ) {
            $serial       = "0"
        }
        if ( $product -eq "" ) {
            $product      = "Unknown"
        }
        # Thats a guess only, but if not it is UNKNOWN
        # Serial is empty or 0 
        # userfriendly is empty
        # device_type is Laptop
        # Maybe Internal Display
        if ( $serial -eq "0" -AND $userfriendly -eq "" -AND $Computer.device_type -eq "Laptop" ) {
            $product      = "Internal Display"
            $userfriendly = $product
        }
        if ( $userfriendly -eq "" ) {
            $userfriendly = $product
        }
        if ( $manu -eq "" ){
            $manu         = "UNK"
        }
        # check if We have long Version of Name         
        if ( $MyManufacturerHash[$manu] ) {
            $manu = $MyManufacturerHash[$manu]
        }
        # No Product
        if ( $product.ToLower() -eq $userfriendly.ToLower() ) {
            $product     = "$( $userfriendly.ToUpper() )"
        } else {
            $product     = "$( $userfriendly.ToUpper() ) ( $product )"
        }

        $description = "Build on: $( $monitor.WeekOfManufacture )/$( $monitor.YearOfManufacture )"
        $Monitorname = "$( $userfriendly -replace $($manu),'' )"
        $Monitorname = $Monitorname.Trim().ToUpper()
        if ( $serial -eq "0" ) {
            # This is for Fixing 0 Serials on Monitors
            [String]$serial  = $Computer.id
        } 

        $serial  = $serial.ToUpper()
        $product = $product.ToUpper()
        
        # Wee need to replace some Chars
        $server_name = $MonitorName -replace " ","_" -replace "\(","" -replace "\)",""

        "$( $manu.trim() ) $( $product.trim() ) ( $( $userfriendly.trim() ) ) >>$( $serial.trim() )<< $( $monitor.WeekOfManufacture )/$( $monitor.YearOfManufacture )"
      
        # operating_system = deviceid of Computer for this Monitor
        $query = "select id, operating_system, primary_owner_name, user_id from devices where device_type='Monitor' and serial_number='$( $serial.tostring() )' and model like '$( $product.tostring() )'"
        $Display_in_Spiceworks = Get-ODBC-Data -MyQuery $query -MyConnection $Spiceworks_Connection_String
        $query
        if ( $Display_in_Spiceworks.operating_system ) {
            if ( $Display_in_Spiceworks.operating_system -ne [String]$Computer.id ) {
               echo "Monitor on different Computer, update Device ID, primary Owner, userid"
               $query = "update devices set user_id='$( $Computer.user_id )', primary_owner_name='$( $Computer.primary_owner_name )', operating_system='$( $computer.id )' where id='$( $Display_in_Spiceworks.id )'"
               # Warn the USER
               If ( $DryRun.ToLower() -eq "yes" ) {
                   echo "!This will be a DryRun, NOTHING WILL BE CHANGED."
                   $query
               } else {
                   echo "!ATTENTION! Will now change the Spiceworks Database!"
                   $query
                   Set-ODBC-Data -MyQuery $query -MyConnection $Spiceworks_Connection_String
               }
               
            } else {
               if ( $Computer.primary_owner_name -ne $Display_in_Spiceworks.primary_owner_name ) {
                   echo "Ok, Monitor on that Machine, but Owner Changed!"
                   $query =" update devices set user_id='$( $Computer.user_id )', primary_owner_name='$( $Computer.primary_owner_name )' where id='$( $Display_in_Spiceworks.id )'"
                   If ( $DryRun.ToLower() -eq "yes" ) {
                       echo "!This will be a DryRun, NOTHING WILL BE CHANGED."
                       $query
                   } else {
                       echo "!ATTENTION! Will now change the Spiceworks Database!"
                       $query
                       Set-ODBC-Data -MyQuery $query -MyConnection $Spiceworks_Connection_String
                   }
                } else {
                    echo "Ok, Nothing changed"
                }
            }
        } else {
            echo "New Monitor Found, not in Database, creating"
            $query = "insert into devices ( primary_owner_name,                  user_id,                  spice_version, mac_address, asset_tag, exclude_tag, user_tag, scan_state, created_on,  updated_on,  site_id,         auto_tag,                       name,                manually_added, serial_number,   model,          manufacturer, operating_system,    device_type, type,      description ) " +
                                 " values ( '$( $Computer.primary_owner_name )', '$( $Computer.user_id )', '1',           '',          '',        '||',        '||',     'offline',  date('now'), date('now'), '$( $Site_id )', '|devices|identified|monitor|', '$( $Monitorname )', 't',            '$( $serial )', '$( $product )', '$( $manu )', '$( $Computer.id )', 'Monitor',   'Unknown', '$( $description )' ) "
            # Warn the USER
            If ( $DryRun.ToLower() -eq "yes" ) {
                echo "!This will be a DryRun, NOTHING WILL BE CHANGED."
                $query
            } else {
                echo "!ATTENTION! Will now change the Spiceworks Database!"
                $query
                Set-ODBC-Data -MyQuery $query -MyConnection $Spiceworks_Connection_String
            }
        }
    }
}