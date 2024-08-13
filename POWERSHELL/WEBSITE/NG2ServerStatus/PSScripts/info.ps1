
Add-PSSnapIn Quest.ActiveRoles.ADManagement -ErrorAction SilentlyContinue

$ScriptBlock = {
    param(
        $filter
    )

    function removeSurplusSpaces([string]$value) {
        return [regex]::Replace($value, "\s+", " ")
    }

    function toUptime([DateTime]$lastBootTime) {
        [TimeSpan]$_uptime = New-TimeSpan $lastBootTime $(get-date)
        return (removeSurplusSpaces("{0,3} days, {1,2} hrs and {2,2} min" -f $_uptime.days, $_uptime.hours, $_uptime.minutes)).Trim()
    }

    $_systemname = $env:COMPUTERNAME

    $_last7Days = (Get-Date).AddDays(-7)
    $_last24H = (Get-Date).AddHours(-24)

    # Disk

    $DiskInfo = Get-WmiObject -Class "Win32_LogicalDisk" -Filter "DriveType=3" -namespace "root\CIMV2" | Select-Object DeviceID, FileSystem, Size, FreeSpace, VolumeName, VolumeDirty

    # CPU

    $CPUInfo = Get-WmiObject -Class "Win32_Processor" -namespace "root\CIMV2" | Select-Object @{label = 'Name'; expression = {$_.Name -replace "\s+", " "}}, Manufacturer, @{label = 'MaxClockSpeed'; expression = {([math]::round($_.MaxClockSpeed / 1024, 2)).ToString() + " GHz"}}, @{label = 'LoadPercentage'; expression = {$_.LoadPercentage.ToString()}}, NumberOfCores, NumberOfLogicalProcessors

    #  Computer

    $ComputerInfo = Get-WmiObject -Class "Win32_ComputerSystem" -namespace "root\CIMV2" | Select-Object Caption, Domain, Model, Name, NumberOfProcessors

    # OS

    $OSInfo = Get-WmiObject -Class "Win32_OperatingSystem" -namespace "root\CIMV2" | Select-Object BuildNumber, Caption, CSDVersion, Version, OSArchitecture, TotalVisibleMemorySize, FreePhysicalMemory, TotalVirtualMemorySize, FreeVirtualMemory, OSType, @{label = 'LastBootUpTime'; expression = {$_.ConvertToDateTime($_.LastBootUpTime).ToString("dd-MMM-yyy HH:mm:ss")}}, @{label = 'Uptime'; expression = {toUptime($_.ConvertToDateTime($_.LastBootUpTime))}}

    # Network

    $NetworkInfo = Get-WmiObject -Class "Win32_NetworkAdapterConfiguration" -Namespace "root\CIMV2" | Where-Object {$_.IPEnabled -eq 'True'} | Select-Object Description, DNSDomain, DHCPServer, IPAddress, MACAddress

    # Memory

    $MemoryInfo = Get-WmiObject -Class "win32_PhysicalMemory" -namespace "root\CIMV2" | Select-Object Capacity, BankLabel, DeviceLocator, PartNumber, SerialNumber

    # Events

    $_filterScriptBlock = [scriptblock]::Create($filter)

    $Events = @()

    Get-EventLog -LogName System -After $_last7Days -Newest 2000 | Where-Object $_filterScriptBlock | ForEach-Object {
        $EventProperties = @{
            EventID       = $_.EventID
            EntryType     = $_.EntryType
            Message       = ((($_.Message -replace "\\r\\n", " ") -replace ",\n", " ") -replace '"', "'") -replace "\s+", " "
            Source        = $_.Source
            TimeGenerated = $_.TimeGenerated.ToString("dd-MMM-yyy HH:mm:ss")
            TimeWritten   = $_.TimeWritten.ToString("dd-MMM-yyy HH:mm:ss")
            Log           = "System"
        }
        $Events += New-Object psobject -Property $EventProperties
    }

    Get-EventLog -LogName Application -After $_last7Days -Newest 2000 | Where-Object $_filterScriptBlock | ForEach-Object {
        $EventProperties = @{
            EventID       = $_.EventID
            EntryType     = $_.EntryType
            Message       = ((($_.Message -replace "\\r\\n", " ") -replace ",\n", " ") -replace '"', "'") -replace "\s+", " "
            Source        = ($_.Source).ToString().Trim()
            TimeGenerated = $_.TimeGenerated.ToString("dd-MMM-yyy HH:mm:ss")
            TimeWritten   = $_.TimeWritten.ToString("dd-MMM-yyy HH:mm:ss")
            Log           = "Application"
        }
        $Events += New-Object psobject -Property $EventProperties
    }

    $_eventsErrors7Days = $Events | Where-Object { $_.EntryType -eq 'Error'}
    $_eventsWarnings7Days = $Events | Where-Object { $_.EntryType -eq 'Warning'}
    $_eventsReboots7Days = $Events | Where-Object { $_.EventID -eq 6009 }

    # Combined info

    $SystemInfo = @()

    $SysInfoProperties = @{
        SysName   = $_systemname
        Disk      = $DiskInfo
        CPU       = $CPUInfo
        Computer  = $ComputerInfo
        OS        = $OSInfo
        Network   = $NetworkInfo
        Memory    = $MemoryInfo
        Reboots   = $_eventsReboots7Days
        Errors    = $_eventsErrors7Days
        Warnings  = $_eventsWarnings7Days
        Generated = (Get-Date).ToString("dd-MMM-yyy HH:mm:ss")
    }

    $SystemInfo += New-Object psobject -Property $SysInfoProperties
    $SystemInfo
}

# EOS #################################################################################################################

function Test-Servers {
    param([string[]]$Computers)

    $alive = @()
    $dead = @()
    foreach ($computer in $Computers) {
        $ping = Test-Connection -ComputerName $computer -Quiet -BufferSize 16 -Count 1 -ErrorAction SilentlyContinue
        if ($ping) {
            $alive += $computer
        }
        if (!$ping) {
            $dead += $computer
        }
    }
    $result = New-Object -TypeName psobject -Property @{
        Alive  = $alive
        Dead   = $dead
        Domain = $env:USERDOMAIN
    }
    $result
}

function EventFilter {
    $filterFileName = "EventsFilter.csv"
    if (Test-Path -Path $filterFileName) {
        $fline = @()
        #Write-Output "Importing events filter ..."
        $filters = Import-CSV -Path $filterFileName
        $filters | ForEach-Object {
            $eID = $_.EventID
            $eType = $_.EntryType
            $source = $_.Source
            $fline += "((`$_.EventID -eq $eID) -and (`$_.EntryType -eq '$eType') -and (`$_.Source -eq '$source'))"
        }
        $filters = "((`$_.EntryType -eq 'Error') -or (`$_.EntryType -eq 'Warning') -or (`$_.EventID -eq 6009)) -and -not ({0})" -f (($fline) -join " -or ")
    }
    return $filters
}

$localName = $env:COMPUTERNAME

$Computers = @()
Get-QADComputer -OSName "Windows Server*" | Where-Object {-not ($_.Name -like $localName)} | ForEach-Object {$Computers += $_.Name}

$ComputersToPing = $Computers + $localName

$connectivity = Test-Servers -Computers $ComputersToPing

$connectivity | ConvertTo-Json -Depth 2 | Out-File -FilePath "assets\ping.json"

$_noinfo = $connectivity.Alive

# Get all not pingable

$missing_servers = $connectivity.Dead

$event_filters = (EventFilter)

$name = "assets\servers.json"

"[" | Out-File -FilePath $name

# Local computer JSON export

$localData = Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList (EventFilter) -ErrorAction Continue

$localData | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath $name -Append

"," | Out-File -FilePath $name -Append

#  Remote servers JSON export

$Job = Invoke-Command -ScriptBlock $ScriptBlock -AsJob -JobName ServerStatusInfo -ComputerName $Computers -HideComputerName -ArgumentList $event_filters
# Wait for the job to complete
# Write-Host "Job created, waiting to finish ..."
Wait-Job -Job $Job  | Out-Null
# Receive data from the job
$Data = Receive-Job $Job | Select-Object * -ExcludeProperty RunspaceId, PSComputerName, PSShowComputerName
# Once we have received the data, we can remove the job
Remove-Job -Job $Job

Write-Debug "Remote data collecting completed"

# Export to JSON

$i = ($Data | Measure-Object).Count
foreach ($item in $Data) {
    $item | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath $name -Append
    $i--
    if (-not $i -eq 0) {
        "," | Out-File -FilePath $name -Append
    }
}

# EOF JSON

"]" | Out-File -FilePath $name -Append

#  EOS EndOfScript
