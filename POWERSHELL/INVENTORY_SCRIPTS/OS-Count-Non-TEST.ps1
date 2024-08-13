
$tableOSName = "OperatingSystems"
$tableOS = New-Object system.Data.DataTable "$tableOSName"
$colOS = New-Object system.Data.DataColumn OperatingSystem,([string])
$colOSversion = New-Object system.Data.DataColumn OperatingSystemVersion,([string])
$colOSType = New-Object system.Data.DataColumn OperatingSystemType,([string])
$tableOS.columns.add($colOS)
$tableOS.columns.add($colOSversion)
$tableOS.columns.add($colOSType)

$rowtableOS = $tableOS.NewRow()
$rowtableOS.OperatingSystem = "Windows Server 2016 Standard"
$rowtableOS.OperatingSystemVersion = "10.0"
$rowtableOS.OperatingSystemType = "Server"
$tableOS.Rows.Add($rowtableOS)

$rowtableOS = $tableOS.NewRow()
$rowtableOS.OperatingSystem = "Windows Server 2016 Datacenter"
$rowtableOS.OperatingSystemVersion = "10.0"
$rowtableOS.OperatingSystemType = "Server"
$tableOS.Rows.Add($rowtableOS)

$rowtableOS = $tableOS.NewRow()
$rowtableOS.OperatingSystem = "Windows Server 2012 R2 Standard"
$rowtableOS.OperatingSystemVersion = "6.3"
$rowtableOS.OperatingSystemType = "Server"
$tableOS.Rows.Add($rowtableOS)

$rowtableOS = $tableOS.NewRow()
$rowtableOS.OperatingSystem = "Windows Server 2012 R2 Datacenter"
$rowtableOS.OperatingSystemVersion = "6.3"
$rowtableOS.OperatingSystemType = "Server"
$tableOS.Rows.Add($rowtableOS)

$rowtableOS = $tableOS.NewRow()
$rowtableOS.OperatingSystem = "Windows Server 2012 Standard"
$rowtableOS.OperatingSystemVersion = "6.2"
$rowtableOS.OperatingSystemType = "Server"
$tableOS.Rows.Add($rowtableOS)

$rowtableOS = $tableOS.NewRow()
$rowtableOS.OperatingSystem = "Windows Server 2012 Datacenter"
$rowtableOS.OperatingSystemVersion = "6.2"
$rowtableOS.OperatingSystemType = "Server"
$tableOS.Rows.Add($rowtableOS)

$rowtableOS = $tableOS.NewRow()
$rowtableOS.OperatingSystem = "Windows Server 2008 R2 Standard"
$rowtableOS.OperatingSystemVersion = "6.1"
$rowtableOS.OperatingSystemType = "Server"
$tableOS.Rows.Add($rowtableOS)

$rowtableOS = $tableOS.NewRow()
$rowtableOS.OperatingSystem = "Windows Server 2008 R2 Enterprise"
$rowtableOS.OperatingSystemVersion = "6.1"
$rowtableOS.OperatingSystemType = "Server"
$tableOS.Rows.Add($rowtableOS)


write-host ""

write-host "Server Operating Systems : " -foregroundcolor "Green"

$ServerCount = 0

foreach ($object in ($tableOS | where {$_.OperatingSystemType -eq 'Server'}))

{

      $LDAPFilter = "(&(operatingsystem=" + $object.OperatingSystem + "*)(operatingsystemversion=" + $object.OperatingSystemVersion + "*)(!cn=*TEST*)(!cn=*TST*))"

      $OSCount = (Get-ADComputer -LDAPFilter $LDAPFilter).Count

      if ($OSCount -ne $null)

      {

            "" + $object.OperatingSystem  + ": " + $OSCount + ""

      }

      else

      {

            "" + $object.OperatingSystem  + ": 0"

            $OSCount = 0

      }

      $ServerCount += $OSCount

}

$ServerTotalNumber = "Total Number : " + $ServerCount + ""

write-host $ServerTotalNumber -foregroundcolor "Yellow"

write-host ""

$LDAPFilter = "(&(operatingsystem=*)"

foreach ($object in $tableOS)

{

      $LDAPFilter += "(!(&(operatingsystem=" + $object.OperatingSystem + "*)(operatingsystemversion=" + $object.OperatingSystemVersion + "*)))"

}

$LDAPFilter += ")"