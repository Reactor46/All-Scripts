# ComputerAudit.ps1
# Written by BigTeddy January 13, 2011
# Last modified January 14, 2012
# Version 2.22
# Exports the results to a csv file.
# Creates a log file of unreachable computers.
# Change the root of the LDAP path to suit.
#
# Gets the following information for each computer listed in a text file:
# Network Name, DN, Location, Room, IP Address, MAC Address, Manufacturer,
# Model, Operating System, OS Service Pack version, OS Install Date,
# Processor type, Processor speed, Processor count, Installed RAM,
# BIOS date, C drive capacity, C drive space.

# Adjust the path of the output file to suit:
$outputFile = 'ComputerAudit.csv'
# Adjust the path of the input file to suit:
$inputFile = 'PCs.txt'
# Adjust the path of the log file to suit:
$logFile = 'AuditLog.txt'
# Adjust LDAP path to suit:
$root = 'LDAP://DC=contoso,DC=internal'

# Leave this part as is:
$a = [adsisearcher]"objectcategory=computer" 
$a.SearchRoot=$root
$a.PageSize = 200
$a.PropertiesToLoad.Add("name") | Out-Null
$a.PropertiesToLoad.Add("location") | Out-Null
$a.PropertiesToLoad.Add("description") | Out-Null
$a.PropertiesToLoad.Add("distinguishedname") | Out-Null
$a.propertiestoload.add("samaccountname") 
$results = $a.findall() 
$ADObjects = @()
foreach($result in $results) {
 [Array]$propertiesList = $result.Properties.PropertyNames
 $obj = New-Object PSObject
 foreach($property in $propertiesList) { 
    $obj | add-member -membertype noteproperty -name $property -value ([string]$result.Properties.Item($property))
    } # end foreach
 $ADObjects += $obj
} # end foreach

New-Item -ItemType File -Path $logFile -Force | Out-Null
$computers = Get-Content $inputFile
$results = @()
foreach ($comp in $computers) {
    $result = New-Object psObject 
    Write-Host "Querying $($comp)" -ForegroundColor green
    # Query AD Objects
    $networkName = ($ADObjects | Where-Object { $_.name -eq $comp } | Select-Object -ExpandProperty samaccountname) -replace '\$$'
    $result | Add-Member -MemberType NoteProperty -Name 'Network Name' -Value $networkName
    $location = $ADObjects | Where-Object { $_.name -eq $comp } | Select-Object -ExpandProperty location
    $result | Add-Member -MemberType NoteProperty -Name 'Location' -Value $location
    $description = $ADObjects | Where-Object { $_.name -eq $comp } | Select-Object -ExpandProperty description
    $result | Add-Member -MemberType NoteProperty -Name 'Room' -Value $description
    $distinguishedName = $ADObjects | Where-Object { $_.name -eq $comp } | Select-Object -ExpandProperty distinguishedname
    $result | Add-Member -MemberType NoteProperty -Name 'DN' -Value $distinguishedname
    try {
        # IP Query
        $IPAddr = Test-Connection -ComputerName $comp -Count 1 -ErrorAction stop  | select -ExpandProperty ipv4Address | `
            select -ExpandProperty ipAddressToString
        $result | Add-Member -MemberType NoteProperty -Name 'IP Address' -Value $IPAddr
        # MAC address query
        $mac = gwmi -Class win32_networkadapterconfiguration -ComputerName $comp -ea stop |  `
            Where-Object {$_.macaddress -and ($_.ipaddress -contains $IPAddr -or $_.ipaddress -eq $IPAddr)} |select -Expand macaddress
        $result | Add-Member -MemberType NoteProperty -Name 'MAC Address' -Value $mac
        # WMI queries
        $computerSystem = gwmi -Class win32_computersystem -ComputerName $comp 
        $OperatingSystem = gwmi -Class win32_operatingsystem -ComputerName $comp
        $Processor = @(gwmi -Class win32_processor -ComputerName $comp)
        $RAM = [math]::round(($computerSystem.TotalPhysicalMemory)/1mb,3)
        $bios = gwmi -Class win32_bios -ComputerName $comp
        $biosDate = [Management.ManagementDateTimeConverter]::ToDateTime($bios.ReleaseDate)
        $OS_InstallDate = [Management.ManagementDateTimeConverter]::ToDateTime($OperatingSystem.InstallDate)
        $C_Drive = gwmi -Class win32_logicaldisk -ComputerName $comp | ? { $_.DeviceID -eq 'C:' }
        # Add results of WMI queries
        $result | Add-Member -MemberType NoteProperty -Name 'Manufacturer' -Value $computerSystem.Manufacturer
        $result | Add-Member -MemberType NoteProperty -Name 'Product Description' -Value $computerSystem.Model
        $result | Add-Member -MemberType NoteProperty -Name 'Operating System' -Value ($OperatingSystem.Name -split '\|')[0] 
        $result | Add-Member -MemberType NoteProperty -Name 'OS_SP' -Value $OperatingSystem.csdVersion
        $result | Add-Member -MemberType NoteProperty -Name 'OS_InstallDate' -Value $OS_InstallDate
        $result | Add-Member -MemberType NoteProperty -Name 'CPU' -Value $Processor[0].name
        $result | Add-Member -MemberType NoteProperty -Name 'ProcSpeed' -Value $Processor[0].maxclockspeed
        $result | Add-Member -MemberType NoteProperty -Name 'ProcCount' -Value $Processor.count
        $result | Add-Member -MemberType NoteProperty -Name 'Memory' -Value $RAM
        $result | Add-Member -MemberType NoteProperty -Name 'BIOSDate' -Value $biosDate
        $result | Add-Member -MemberType NoteProperty -Name 'Hard Drive Capacity' -Value ([math]::Round(($C_Drive.size)/1gb,3))
        $result | Add-Member -MemberType NoteProperty -Name 'C_Space' -Value ([math]::Round(($C_Drive.freespace)/1gb,3))
        } # end try   
    catch {
        Write-Host "Could not contact $($comp)" -ForegroundColor Red
        Out-File -FilePath $logFile -InputObject $comp -Append -Force
        } # end catch
    
    $results += $result
    
    } # end foreach $comp

$results | Export-Csv -Path $outputFile -NoTypeInformation -Force
$results | Select-Object -First 20 | Format-Table
Write-Host "$($results.count) computers processed. The results are stored in $outputFile"
$offlineCount = @(Get-Content $logFile -ErrorAction silentlycontinue).count
Write-Host "$offlineCount computers were offline. The list is stored in $logFile"
