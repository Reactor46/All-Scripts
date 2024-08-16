# Copyright (c) 2009 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
# This file contains Content Index Troubleshooter functions
#

# Include the global constants and types
#
$scriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
. $scriptPath\CITSConstants.ps1
. $scriptPath\CITSTypes.ps1

<#
   .DESCRIPTION
   Validate-Arguments is called by Troubleshoot-CI.ps1 script to
   perform additional validation of command-line arguments. 
   If validation fails, this function throws an ArgumentException
   with specific information.

   .PARAMETER Server
   The simple NETBIOS name of mailbox server on which troubleshooting
   should be atempted for CI catalogs. If this optional parameter is
   not specified, local server is assumed. 

   .PARAMETER Database
   The name of database to troubleshoot. If this optional parameter is
   not specified, catalogs for all databases on the server specified
   by the Server parameter are troubleshooted.
   
   .PARAMETER Symptom
   Specifies the symptom to detect. Possible values are:
   'Deadlock', 'Corruption', 'Stall', 'Backlog' and 'All'.
   When 'All' is specified, all the first four symptoms in
   the list are checked.
   
   If this optional parameter is not specified, 'All' is assumed.
   
    .PARAMETER Action
   Specifies the action to be performed to resolve a symptom. The
   possible values are 'Detect', 'DetectAndResolve', 'Resolve'.
   The default value is 'Detect'
     
    .PARAMETER MonitoringContext
   Specifies if the command is being run in a monitoring context.
   The possible values are $true and $false. Default is $false.
   If the value is $true, warning/failure events are logged to the
   application event log. Otherwise, they are not logged.
   
   .PARAMETER FailureCountBeforeAlert
   Specifies the number of failures the troubleshooter will allow
   before raising an Error in the event log, leading to a SCOM alert.
   The allowed range for this parameter is [1,100], both inclusive.
   
   .PARAMETER FailureTimeSpanMinutes
   Specifies the number of minutes in the time span during which
   the troubleshooter will check the history of failures to count
   the failures and alert. If the failure count during this time
   span exceeds the value for FailureCountBeforeAlert, an alert
   is raised. No alerts are rasised if MonitoringContext is $false.
   The default value for this parameter is 600, which is equivalent
   to 10 hours.
   
   .INPUTS
   None. You cannot pipe objects to Troubleshoot-CI.ps1.

   .OUTPUTS
   Returns an object of type Arguments   
#>
function Validate-Arguments
{
    Param(
        [String]
        $Server,
        [String]
        $Database,
        [String]
        $Symptom,
        [String]
        $Action,
        [Switch]
        $MonitoringContext,
        [Int32]
        $FailureCountBeforeAlert,
        [Int32]
        $FailureTimeSpanMinutes
    )

    # if resolution is requested, only a specific
    # symptom is allowed, 'All' is not allowed.
    #
    if ($Action -ieq "Resolve" -and $Symptom -ieq "All" )
    {
        $argError = new-object System.ArgumentException ($LocStrings.AllNotAllowedForResolve)
        throw $argError
    }
    
    $Arguments = new-object -typename Arguments
    
    # If server name wasn't supplied, default to
    # local server name
    #
    if ([System.String]::IsNullOrEmpty($Server))
    {
        $Arguments.Server = $env:computername
    }
    else
    {
        $Arguments.Server = $Server
    }
    
    $Arguments.Database = $Database
    $Arguments.Symptom = $Symptom
    $Arguments.Action = $Action
    $Arguments.InstanceName = $null
    $Arguments.MonitoringContext = $MonitoringContext
    $Arguments.WriteApplicationEvent = $MonitoringContext
    $Arguments.FailureCountBeforeAlert = $FailureCountBeforeAlert
    $Arguments.FailureTimeSpanMinutes = $FailureTimeSpanMinutes
    
    return $Arguments
}

<#
   .DESCRIPTION
   Detects problems with catalog copies specified by the $Server and
   $Database parameters

   .PARAMETER Server
   The simple NETBIOS name of mailbox server. 

   .PARAMETER Database
   The name of database.
   
   .PARAMETER After
   The start time from which to scan for issues. Specifically, this is 
   used for checking bad disk block issues in event log. Troubleshoot-CI.ps1
   uses a default of 30 min before the script is run. This parameter
   is added mostly for use by tests.
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   An instance of the ServerStatus object.   
#>
function Detect-Problems
{
    Param(
        [String]
        [ValidateNotNullOrEmpty()]
        $Server,
        [String]
        $Database,
        [DateTime]
        $After
    )
    
    $copyArray = Get-Copies -Server $Server -Database $Database
    
    $ciStatusArray = Get-CIStatus -copies $copyArray
    
    if ($After -eq $null -or $After -eq [DateTime]::MinValue)
    {
        $startTime = (get-date).AddMinutes(-1 * $badDiskBlockCheckIntervalInMinutes)
    }
    else
    {
        $startTime = $After
    }
    
    write-verbose "checking bad block issues only after $startTime"
    
    $ciStatusArray = Check-BadDiskBlocks -Server $Server -StartTime $startTime -CIStatusArray $ciStatusArray

    foreach ($status in $CIStatusArray)
    {
        if (IsCatalogStalled -CIStatus $status -StallThresholdInSeconds $stallThresholdInSeconds)
        {
            write-verbose ("Catalog " + $status.Name + " is stalled.")
            $status.IsStalled = $true
        }
        
        if (IsCatalogBacklogged -CIStatus $status -BacklogThresholdInSeconds $backlogThresholdInSeconds)
        {
            write-verbose ("Catalog " + $status.Name + " is backlogged.")
            $status.isBacklogged = $true
        }
        
        if (IsCatalogHealthStale -CIStatus $status -StaleThresholdInSeconds $staleThresholdInSeconds)
        {
            write-verbose ("Catalog health for " + $status.Name + " is stale.")
            $status.IsHealthStale = $true
        }
      
        if (IsCatalogCorrupted -CIStatus $status)
        {
            write-verbose ("Catalog " + $status.Name + " is corrupted.")
            $status.IsCorrupted = $true            
        }
        
        # If the catalog has a bad block on the disk, set the IsCorrupted to true
        # so that the corresponding resolution action (reseed) can be taken.
        #
        if ($status.HasBadDiskBlock)
        {
            write-verbose ("Catalog " + $status.Name + " has bad disk block.")
            $status.IsCorrupted = $true
        }
        
        # If the catalog associated with passive is in a crawling state, set the IsCorrupted to true
        # so that the corresponding resolution action (reseed) can be taken.
        #
        if(IsCrawling -Copies $copyArray -CIStatus $status)
        {
            write-verbose ("Catalog " + $status.Name + " is in crawling state.")
            $status.IsCorrupted = $true
        }
    }
    
    $serverStatus = new-object -typename ServerStatus
    $serverStatus.Name = $Server
    $serverStatus.IsDeadlocked = $false
    $serverStatus.CatalogStatusArray = $ciStatusArray
    
    # if a database was specified in the 
    # argument list, it doesn't make sense to check
    # if the entire set of catalogs is deadlocked. 
    # Otherwise, check for deadlock using all catalogs 
    # in the status array
    #
    if ([System.String]::IsNullOrEmpty($Database))
    {
        if ((IsDeadlocked $CIStatusArray))
        {
            $serverStatus.IsDeadlocked = $true  
        }    
    }
    
    return $serverStatus    
}

<#
   .DESCRIPTION
   Builds a fake ServerStatus object for a specified 
   Symptom. Used when a specific resolution action
   was requested, overriding any real status.

   .PARAMETER Server
   The simple NETBIOS name of mailbox server. 

   .PARAMETER Database
   The name of database.
   
   .PARAMETER Symptom
   The symptom to use in building the server status object.
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   An instance of the ServerStatus object.   
#>
function Build-ServerStatus
{
    Param(
        [String]
        [ValidateNotNullOrEmpty()]
        $Server,
        [String]
        [AllowNull()]
        $Database,
        [String] 
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Deadlock", "Corruption", "Stall")] 
        $Symptom
    )
  
    $serverStatus = new-object -typename ServerStatus
    $serverStatus.Name = $Server
    
    if ($Symptom -ieq "Deadlock")
    {
        $serverStatus.IsDeadlocked = $True
    }
    else
    {
        $copyArray = Get-Copies -Server $Server -Database $Database
        $ciStatusArray = Get-CIStatus -copies $copyArray
        $serverStatus.IsDeadlocked = $False
        $serverStatus.CatalogStatusArray = $ciStatusArray

        foreach ($status in $serverStatus.CatalogStatusArray)
        {
            if ($Symptom -ieq "Stall")
            {
                $status.IsStalled = $True
            }
            elseif ($Symptom -ieq "Corruption")
            {
                $status.IsCorrupted = $True
            }
        }
    }
        
    return $serverStatus    
}

<#
   .DESCRIPTION
   Logs events to application log based on detection
   results. 

   .PARAMETER Arguments
   Object of type Arguments, containing command-line
   arguments 

   .PARAMETER ServerStatus
   Object of type ServerStatus, containing the current
   status of catalogs as returned by Detect-Problems
   function.
   
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   None.
#>
function Log-DetectionResults
{
    Param(
        [Object]
        [ValidateNotNull()]
        $Arguments,
        [Object]
        [ValidateNotNull()]
        $ServerStatus
    )
    
    $issuesFound = $False
    
    # If the server status is deadlocked, log it.
    #
    if ($ServerStatus.IsDeadlocked)
    {
        $issuesFound = $True
        Log-Event -Arguments $Arguments -EventInfo $LogEntries.DetectedDeadlock
    }
    
    foreach ($catalog in $ServerStatus.CatalogStatusArray)
    {
        if ($catalog.IsStalled)
        {
            $issuesFound = $True
            Log-Event -Arguments $Arguments -EventInfo $LogEntries.DetectedIndexingStall -Parameters @($catalog.DatabaseName)
        }
        elseif ($catalog.IsCorrupted)
        {
            $issuesFound = $True
            Log-Event -Arguments $Arguments -EventInfo $LogEntries.DetectedCatalogCorruption -Parameters @($catalog.DatabaseName)
        }
        elseif ($catalog.IsBacklogged)
        {
            $issuesFound = $True
            $hoursBacklogged = $backlogThresholdInSeconds/3600
            [string[]]$parameters = ($Catalog.DatabaseName, $hoursBacklogged)
            Log-Event -Arguments $Arguments -EventInfo $LogEntries.DetectedIndexingBacklog -Parameters $parameters
        } 
        else
        {
            Log-Event -Arguments $Arguments -EventInfo $LogEntries.CatalogHasNoIssues -Parameters @($catalog.DatabaseName)
        }       
    }
    
    # If no issues are found, log the fact.
    # This will help turn previous alerts
    # green.
    #
    if (!($issuesFound))
    {
        Log-Event -Arguments $Arguments -EventInfo $LogEntries.DetectedNoIssues -Parameters @($catalog.DatabaseName)
    }
}

<#
   .DESCRIPTION
   Attempts resolution of problems with CI catalogs

   .PARAMETER Arguments
   Object of type Arguments, containing command-line
   arguments 

   .PARAMETER ServerStatus
   Object of type ServerStatus, indicating the current
   status of catalogs.
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   A modified server status object with resolution
   status added to each catalog status object.   
#>
function Resolve-Problems
{
    Param(
        [Object]
        [ValidateNotNullOrEmpty()]
        $Arguments,
        [Object]
        [ValidateNotNullOrEmpty()]
        $ServerStatus
    )

    # $todo$ if resolution is already
    # in progress on the target server
    # initiated by some other instance of
    # troubleshooter, log error and return
    #
    
    try
    {
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $LogEntries.TSResolutionStarted

        $stalled = $False
        # Are any catalogs stalled?
        #
        foreach ($catalog in $ServerStatus.CatalogStatusArray)
        {
            if ($catalog.IsStalled -eq $True)
            {
                write-verbose ($catalog.Name + " is stalled.")
                $stalled = $True
                break
            }
        }
        
        if ($ServerStatus.IsDeadlocked -or $stalled)
        {
            write-verbose ("Detected indexing stalls or a deadlock. Restarting search services on " + $Arguments.Server)
            Restart-SearchServices $Arguments $defaultRestartTimeout   
        }
            
        # now look for corruptions and start reseeding
        # each corrupted catalog
        #
        foreach ($catalog in $ServerStatus.CatalogStatusArray)
        {
            if ($catalog.IsCorrupted -eq $True)
            {
                write-verbose ($catalog.Name + " seems to be corrupted. Reseeding the catalog..")
                Reseed-Catalog -Arguments $Arguments -Catalog $catalog
            }
        }
        
        # Log success event
        #
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $LogEntries.TSResolutionFinished
    }
    catch [System.Exception]
    {
        $message=($error[0].Exception.ToString() + $error[0].InvocationInfo.PositionMessage)
        write-verbose ("Caught Exception: $message")
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $LogEntries.TSResolutionFailed `
            -Parameters @($message)
    }
}

<#
   .DESCRIPTION
   Gets all copies of given database on the given server.
   If server parameter is null or empty, local server is assumed.
   If database is non-null and non-empty, just that copy 
   on the given server or local server is returned.

   .PARAMETER Server
   The simple NETBIOS name of mailbox server.

   .PARAMETER Database
   The name of database.
      
   .INPUTS
   None. You cannot pipe objects to Troubleshoot-CI.ps1.

   .OUTPUTS
   An array of database copy objects.   
#>
function Get-Copies
{
    Param(
        [String] [AllowNull()]
        $Server,
        [String] [AllowNull()]
        $Database
    )
    
    return (Get-MatchingDatabaseCopyStatusObjects -Server $Server -Database $Database)
}

<#
   .DESCRIPTION
   Gets content index catalog status for a set of database copies

   .PARAMETER Copies
   An array of mailboxdatabasecopy objects
         
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   An array of CIStatus objects.   
#>
function Get-CIStatus
{
    Param(
        [Object[]]
        [ValidateNotNullOrEmpty()]
        $Copies
    )
    
    write-verbose "In function Get-CIStatus"

    $serversPopulated = new-object System.Collections.ArrayList
    foreach ($copy in $Copies)
    {
        $svr = $copy.MailboxServer.ToLower()
        if (-not ($serversPopulated.Contains($svr)))
        {
            write-verbose ("Now populating " + $svr)
            Populate-CounterTable -Server $svr
            write-verbose ("Adding counters for server " + $svr)
            $index = $serversPopulated.Add($svr)
        }
    }
    
    $statusList = @()
    foreach ($copy in $Copies)
    {
        $hashKeyPrefix = ("\\" + $copy.MailboxServer.toLower() + "\" + "msexchange search indices(" + $copy.DatabaseName + ")\")
        write-verbose ("hashKeyPrefix=" + $hashKeyPrefix)
        $aolni = Get-CachedCounter -Value ($hashKeyPrefix + $aolniCounterName)
        $tslni = Get-CachedCounter -Value ($hashKeyPrefix + $tslniCounterName)
        
        $ciStatus = new-object -typename CIStatus
        $ciStatus.Name = $copy.Name
        $ciStatus.DatabaseName = $copy.DatabaseName
        $ciStatus.BacklogCounter  = $aolni
        $ciStatus.StallCounter = $tslni
        $ciStatus.Health = $copy.ContentIndexState
                
        $ciHealth = Get-CatalogHealth -Server $copy.MailboxServer -Database $copy.DatabaseName
        $ciStatus.HealthReason = $ciHealth.ErrorCode
        $ciStatus.HealthTimestamp = $ciHealth.Timestamp
 
        # Initialize detection flags
        #
        $ciStatus.IsStalled = $false
        $ciStatus.IsBacklogged = $false
        $ciStatus.IsCorrupted = $false
        $ciStatus.IsHealthStale = $false
        $ciStatus.HasBadDiskBlock = $false
        
        $statusList += $ciStatus
    }
    
    return $statusList
}

<#
   .DESCRIPTION
   Checks the application event log for MSSearch crashes and
   then checks the System event log for any bad block errors.
   Then the function maps the disk in the error log to the catalog
   and sets the status.HasBadDiskBlock to true/false.
   
   .PARAMETER Server
   The simple NETBIOS name of mailbox server.

   .PARAMETER StartTime
   The time from which to check badk disk/msftesql crash events.
 
   .PARAMETER CIStatusArray
   An array of CIStatus objects
         
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   Modified (if necessary) array of CIStatus objects   
#>
function Check-BadDiskBlocks
{
    Param(
        [String] 
        [ValidateNotNullOrEmpty()]
        $Server,
        [DateTime]
        $StartTime,  
        [Object[]]
        [ValidateNotNullOrEmpty()]
        $CIStatusArray
    )
    
    # Initialize HasBadDiskBlock member of all status objects
    #
    foreach ($status in $CIStatusArray)
    {
        $status.HasBadDiskBlock = $false
    }
    
    # Check if we have any msftesql crashes in the past N minutes
    #
    try
    {
        $msftesqlCrashes = get-eventlog -computername $Server -after $StartTime -logname "Application" -source $msftesqlServiceName | where {$_.eventId -eq $msftesqlCrashEventId}
    }
    catch [System.Exception]
    {
        $msftesqlCrashes = $null
    }
    
    if (($msftesqlCrashes -eq $null) -or ($msftesqlCrashes.Count -eq 0))
    {
        write-verbose "Check-BadDiskBlocks did not find any msftesql crashes in event log since $startTime"
        return $CIStatusArray
    }
    
    # Check if we have any bad disk block errors after $startTime.
    # If yes, map each disk in the error log to a catalog and
    # set the HasBadDiskBlock member of the corresponding status object
    # to true.
    #
    try
    {
        $badDiskEvents = get-eventlog -computername $Server -after $StartTime -logname "System" -source $diskSourceName | where {$_.eventId -eq $badDiskEventId}
    }
    catch [System.Exception]
    {
        $badDiskEvents = $null
    }
    
    if (($badDiskEvents -eq $null) -or ($badDiskEvents.Count -eq 0))
    {
        write-verbose "Check-BadDiskBlocks did not find any bad disk block events in event log after $startTime"
        return $CIStatusArray
    }
    
    # Scan bad disk events, and get the unique bad disk names.
    #
    $badDiskNames=@{}
    $i = 0
    foreach ($event in $badDiskEvents)
    {
        if (($event.ReplacementStrings -eq $null) -or ($event.ReplacementStrings.Length -eq 0))
        {
            continue;
        }
        
        $diskName = ($event.ReplacementStrings[0]).ToLower()
        if (!$badDiskNames.Contains($diskName))
        {
            $badDiskNames.Add($diskName,$i)
            $i++
        }
    }
    
    if ($badDiskNames.Keys -eq $null -or $badDiskNames.Keys.Count -eq 0)
    {
        write-verbose "No bad disk names found in Disk event logs"
        return $CIStatusArray
    }
    
    foreach ($diskName in $badDiskNames.Keys)
    {
        # $diskName is in the format "\device\harddisk3\dr3"
        # we need to extract the disk number i.e., 3 from it.
        #
        $parts = $diskName.Split("\")
        $prefix = "harddisk"
        foreach ($part in $parts)
        {
            if (($part -eq $null) -or ($part.Length -eq 0))
            {
                continue
            }
            
            if ($part.StartsWith($prefix))
            {
                $number = $part.Substring($prefix.Length)
                $diskNumber = [int]$number
                
                write-verbose "Extracted disk number $diskNumber from $part"
                $databaseName = Map-DiskNumberToDatabase -DiskNumber $diskNumber
                
                if ($databaseName -eq $null)
                {
                    write-verbose "$diskNumber could not be mapped to any database"
                    continue
                }
                
                write-verbose "$diskNumber maps to $databaseName"
                
                foreach ($status in $CIStatusArray)
                {
                    if ($status.DatabaseName -ieq $databaseName)
                    {
                        write-verbose "match found: $databaseName"
                        $status.HasBadDiskBlock = $true
                    }
                }
            }
        }
    }
    
    return $CIStatusArray
}

<#
   .DESCRIPTION
   Returns the name of database, given a physical disk number

   .PARAMETER DiskNumber
   physical disk number  
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   name of database hosted on that disk.   
#>

function Map-DiskNumberToDatabase
{
    Param(
        [Int32] 
        $DiskNumber
    )

    # The block below to parse the DiskPart output was
    # provided by Daniel Joiner.
    # It tries to find the volume name
    #
    $diskdetails = "select disk $DiskNumber`ndetail disk" | DiskPart
    for($j=0; $j -lt $diskdetails.Count; $j++)
    {
        if($diskdetails[$j] -match "^  Volume [0-9]+ +")
        {
            if($diskdetails[$j+1] -match "^    [a-zA-Z0-9_-]")
            {
                $MountPoint = $diskdetails[$j+1] -replace "^    (.*)",'$1'
                break
            }
            else
            { 
                $MountPoint = $diskdetails[$j] -replace "^  Volume [0-9]+ +([A-Z]).*",'$1' 
            }
        }
    }
    
    write-verbose "Looking for database with the mount point: $MountPoint"

    if ($MountPoint -eq $null)
    {
        return $null
    }
    
    $databases = @(Get-MailboxDatabase | ?{$_.EdbFilePath -like "$MountPoint*"})
    
    if ($databases -eq $null -or $databases.Length -eq 0)
    {
        return $null
    }
    
    # Since we have only one database/catalog per disk, 
    # we only need to get one name.
    return $databases[0].Name
}

<#
   .DESCRIPTION
   Returns an entry from the perf counter cache hash.  
   The purpose of the function is
   to provide a point for injection during testing

   .PARAMETER Value
   The key of to look up  
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   the value of the array at the key.   
#>
function Get-CachedCounter
{
    Param(
        [String] 
        [ValidateNotNullOrEmpty()]
        $Value
    )

    return $counterHashTable[$Value]
}

<#
   .DESCRIPTION
   Gets one or more databasecopystatus objects

   .PARAMETER Server
   The simple NETBIOS name of mailbox server.

   .PARAMETER Database
   The name of database.
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   An array of database copy objects.   
#>
function Get-MatchingDatabaseCopyStatusObjects
{
    Param(
        [String] [AllowNull()]
        $Server,
        [String] [AllowNull()]
        $Database
    )
    
    if ([System.String]::IsNullOrEmpty($Database))
    {
        $dbs = @(get-mailboxdatabase -Server $Server)
    }
    else
    {
        $dbs = @($Database)
    }
    
    $dbCount = $dbs.Length
    write-verbose "Found $dbCount databases on $Server"
    
    $copyIdList = @()
    foreach ($db in $dbs)
    {
        $copyIdList += "$db\$Server"
    }
    
    $copies = @()
    foreach ($copyId in $copyIdList)
    {
        $copy = Get-MailboxDatabaseCopyStatus $copyId
        $copies += $copy

    }
    
    return $copies   
}

<#
   .DESCRIPTION
   Gets catalog performance counters for a server and stores them
   in the counter cache.

   .PARAMETER Server
   The name of mailbox server  
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   An array of database copy objects.   
#>
function Populate-CounterTable
{
    Param(
        [String] 
        [ValidateNotNullOrEmpty()]
        $Server
    )

    write-verbose ("In Populate-CounterTable " + $Server)
    
    $c = Get-CatalogCounters $Server
    
    foreach ($sample in $c.CounterSamples)
    {
        $path = $sample.Path
        write-verbose ("Adding value for " + $path)
        if (-not ([System.String]::IsNullOrEmpty($path)))
        {
            $counterHashTable[$path] = $sample.CookedValue
        }
    }
}

<#
   .DESCRIPTION
   Gets catalog performance counter samples from a server.
   Only the counters defined in the global counter array
   are obtained.

   .PARAMETER Server
   The name of mailbox server
   
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   Performance counter data.   
#>
function Get-CatalogCounters
{
    Param(
        [String] [ValidateNotNullOrEmpty()]
        $Server
    )
    
    return (Get-Counter $counters -MaxSamples 1 -ComputerName $Server -ErrorAction SilentlyContinue)
}

<#
   .DESCRIPTION
   Gets the catalog health object from registry

   .PARAMETER Server
   The name of mailbox server
   
   .PARAMETER Database
   The name of database.     
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   CatalogHealth object. 
#>
function Get-CatalogHealth
{
    Param(
        [String] [ValidateNotNullOrEmpty()]
        $Server,
        [String] [ValidateNotNullOrEmpty()]
        $Database
    )
 
    write-verbose "In function Get-CatalogHealthRegKey"
    
	$dbGuid = Get-DatabaseGuid $Database
    
    return (Get-HealthFromRegistry $Server $dbGuid)    
}

<#
   .DESCRIPTION
   Gets the GUID of a database
   
   .PARAMETER Database
   The name of database.     
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   Guid of given database.   
#>
function Get-DatabaseGuid
{
    Param(
        [String] [ValidateNotNullOrEmpty()]
        $Database
    )
    
    write-verbose ("Get-MailboxDatabase " + $Database)
 	$db = get-mailboxdatabase $Database
	$dbGuid = $db.Guid
   
    return $dbGuid
}

<#
   .DESCRIPTION
   Gets catalog health from remote/local registry
   
   .PARAMETER Server
   The name of mailbox server

   .PARAMETER DatabaseGuid
   The GUID of database.     
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   CatalogHealth object.   
#>
function Get-HealthFromRegistry
{
    Param(
        [String] [ValidateNotNullOrEmpty()]
        $Server,
        [String] [ValidateNotNullOrEmpty()]
        $DatabaseGuid
    )
    
	$baseKey= [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $Server)
	if($baseKey -eq $null)
	{
        $exception = new-object System.InvalidOperationException ($LocStrings.RegistryOpenError + $Server)
        throw $exception
	}
	$keyPath = "$copyHealthKeyPath{$databaseGuid}"
	$key = $baseKey.OpenSubKey($keyPath)
	if($Key -eq $null)
	{
        $exception = new-object System.InvalidOperationException ($LocStrings.RegistryReadError + $keyPath)
        throw $exception
	}
    
    $regHealth = new-object -typename CatalogHealth
    $regHealth.ErrorCode = $key.GetValue("ErrorCode")
    $regHealth.Timestamp = $key.GetValue("TimeStamp")
        
    return $regHealth
}

<#
    .DESCRIPTION
    Determines if search service(s) is/are deadlocked

    .PARAMETER CIStatusArray
    Array of CI Status objects
         
    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    $true if deadlock was detected, $false otherwise   
#>
function IsDeadlocked
{
    Param(
        [Object[]]
        [ValidateNotNullOrEmpty()]
        $CIStatusArray
    )

    # CI is deemed deadlocked if either of these 
    # conditions happen:
    #
    # 1. All catalog health timestamps are stale 
    # 2. All catalogs are stalled
    #
    
    $allTimestampsStale = $false
    $staleCatalogCount = 0

    $allCatalogsStalled = $false
    $stalledCatalogCount = 0
    
    foreach ($status in $CIStatusArray)
    {
        if ($status.IsHealthStale)
        {
            $staleCatalogCount = $staleCatalogCount + 1
        }
        
        if ($status.IsStalled)
        {
            $stalledCatalogCount = $stalledCatalogCount + 1
        }
    }
    
    if ($staleCatalogCount -ge $CIStatusArray.Length)
    {
        write-verbose ("Health status in registry is stale for all catalogs")
        $allTimestampsStale = $true
    } 
    
    if ($stalledCatalogCount -ge $CIStatusArray.Length)
    {
        write-verbose ("Indexing stalled for all catalogs")
        $allCatalogsStalled = $true
    } 
    
    return ($allTimestampsStale -or $allCatalogsStalled)
}

<#
    .DESCRIPTION
    Determines if indexing is stalled for a catalog 

    .PARAMETER CIStatus
    Catalog status object for the catalog
         
    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    $true if a stall was detected, $false otherwise   
#>
function IsCatalogStalled
{
    Param(
        [Object]
        [ValidateNotNull()]
        $CIStatus,
        [Int32]
        $StallThresholdInSeconds
    )
 
    if ($CIStatus.StallCounter -ge $StallThresholdInSeconds)
    {
        return $true
    }
    
    return $false
}

<#
    .DESCRIPTION
    Determines if status data indicates corruption for a catalog. 

    .PARAMETER CIStatus
    Catalog status object for the catalog
         
    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    $true if a corruption was indicated, $false otherwise   
#>
function IsCatalogCorrupted
{
    Param(
        [Object]
        [ValidateNotNull()]
        $CIStatus
    )
 
    if ($CIStatus.HealthReason -ieq $corruptionIndicator)
    {
        return $true
    }
    
    return $false
}

<#
    .DESCRIPTION
    Determines if indexing is backlogged for a catalog 

    .PARAMETER CIStatus
    Catalog status object for the catalog
         
    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    $true if a backlog was detected, $false otherwise   
#>
function IsCatalogBacklogged
{
    Param(
        [Object]
        [ValidateNotNull()]
        $CIStatus,
        [Int32]
        $BacklogThresholdInSeconds
    )
 
    if ($CIStatus.BacklogCounter -ge $BacklogThresholdInSeconds)
    {
        return $true
    }
    
    return $false
}

<#
    .DESCRIPTION
    Determines if health status in registry is stale for a catalog 
    Returns true if catalog health timestamp is older than the 
    threshold time span for the given catalog

    .PARAMETER CIStatus
    Catalog status object for the catalog
         
    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    $true if health status was stale, $false otherwise   
#>
function IsCatalogHealthStale
{
    Param(
        [Object]
        [ValidateNotNull()]
        $CIStatus,
        [Int32]
        $StaleThresholdInSeconds
    )
 
    write-verbose ("CIStatus.HealthTimestamp = " + $CIStatus.HealthTimestamp)
    
    $staleThreshold = New-TimeSpan -Seconds $StaleThresholdInSeconds
    $now = (Get-Date)
    write-verbose ("current time = " + $now)
    $lastModified = (Get-Date $CIStatus.HealthTimestamp)
    write-verbose ("Health status for " + $CIStatus.Name + " last modified at " + $lastModified)
    $timeSpan = New-TimeSpan -Start $lastModified -End $now
    
    if ($timeSpan -gt $staleThreshold)
    {
        write-verbose ("Health status in registry for " + $CIStatus.Name + " is stale")
        return $true;
    }

    return $false
}

<#
    .DESCRIPTION
    Returns true if catalog for a passive is crawling

    .PARAMETER Copies
    An array of mailboxdatabasecopy objects
         
    .PARAMETER CIStatus
    Catalog status object for the catalog
         
    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    $true if catalog for a passive is crawling, $false otherwise   
#>
function IsCrawling
{
    Param(
        [Object[]]
        [ValidateNotNullOrEmpty()]
        $Copies,

    [Object]
        [ValidateNotNull()]
        $CIStatus
    )

    foreach($copy in $Copies)
    {
        if(($copy.Name -eq $CIStatus.Name) -and ($copy.DatabaseName -eq $CIStatus.DatabaseName))
        {
                if(($CIStatus.Health -eq "Crawling") -and ($copy.Status -eq "Healthy"))
                {
                    return $true
                }
                return $false
        }
    }
}

<#
   .DESCRIPTION
   Restarts search services. This is a simpler
   overload for restart-searchservices, with
   default values for other parameters 

   .PARAMETER Server
   The Server on which to restart the services
   
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   None. Throws exception if not successful.   
#>
function Restart-SearchServices1
{
    Param(
        [string] 
        [ValidateNotNullOrEmpty()]
        $Server
    )
    
    $arguments = new-object -typename Arguments
    $arguments.Server = $Server
    
    Restart-SearchServices -Arguments $arguments -Timeout $defaultRestartTimeout
}

<#
   .DESCRIPTION
   Restarts search services 

   .PARAMETER Arguments
   The Arguments object constructed from script args
   
   .PARAMETER Timeout
   Time span to wait for stopping/terminating
   services. If processes could not be stopped within
   the timeout, a TimeoutException is raised.
   
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   None. Throws exception if not successful.   
#>
function Restart-SearchServices
{
    Param(
        [Object] 
        [ValidateNotNull()]
        $Arguments,
        [TimeSpan]
        $Timeout
    )
    
    # $TODO$ log RestartingSearchServices event
    #
    
    try
    {
        Validate-Timeout $Timeout "Restart-SearchServices"
        
        $deadline = (Get-Date) + $Timeout
        write-verbose ("Restart-SearchServices deadline = " + $deadline)
        
        # Stop the ExSearch service
        #
        Stop-SearchService -Server $Arguments.Server -ServiceName $exsearchServiceName -ProcessName $exsearchProcessName -Timeout $Timeout
        
        # Stop msftesql service
        #
        $newTimeout = $deadline - (Get-Date)
        Stop-SearchService -Server $Arguments.Server -ServiceName $msftesqlServiceName -ProcessName $msftesqlProcessName -Timeout $newTimeout
       
        $startTime = (Get-Date)
        # Now start ExSearch service.  
        #
        $newTimeout = $deadline - (Get-Date)
        Start-SearchService -Server $Arguments.Server -ServiceName $exsearchServiceName -ProcessName $exsearchProcessName -Timeout $newTimeout
        
        # At this point, we expect msftesql service 
        # is also started. Otherwise, start it.
        #
        $newTimeout = $deadline - (Get-Date)
        Start-SearchService -Server $Arguments.Server -ServiceName $msftesqlServiceName -ProcessName $msftesqlProcessName -Timeout $newTimeout
        
        # Wait until we see an event saying exsearch service started successfully
        #
        $newTimeout = $deadline - (Get-Date)
        Wait-ForEvent `
            -Server $Arguments.Server `
            -LogName "Application" `
            -EventSource $exsearchEventSource `
            -EventId 100 `
            -StartTime $startTime `
            -Timeout $newTimeout
            
        # Log Success
        #
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $LogEntries.RestartSuccess 
    }
    catch
    {
        $reason = $error[0].Exception.ToString() + $error[0].InvocationInfo.PositionMessage
        
        write-verbose ("Restart-SearchServices failed. Reason: " + $reason)
        # Log Failure
        #
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $LogEntries.RestartFailure `
            -Parameters ($reason)
    }
}

<#
   .DESCRIPTION
   Reseeds a catalog from the active instance 

   .PARAMETER Arguments
   The Arguments object constructed with script args
   
   .PARAMETER Catalog
   The name of mailbox database copy corresponding
   to the catalog that needs to be reseeded
   
   .OUTPUTS
   None. Throws exception if not successful.   
#>
function Reseed-Catalog
{
    Param(
        [Object] 
        [ValidateNotNull()]
        $Arguments,
        [Object]
        [ValidateNotNull()]
        $Catalog
    )

    $errorPref = $ErrorActionPreference
    # Change the error action preference so that any error during reseed
    # is reported as a failure. Without this, an error in 
    # update-mailboxdatabasecopy is not thrown as an exception.
    #
    $ErrorActionPreference="Stop"
    try
    {
        $problemdb = Get-MailboxDatabase -Identity $Catalog.DatabaseName -Status
        if ($problemdb.Mounted)
        {
            Update-MailboxDatabaseCopy $Catalog.Name -DeleteExistingFiles -Force -confirm:$false -CatalogOnly -ErrorAction "Stop"
            
            # Log success event
            #
            [string[]]$parameters = @($Catalog.DatabaseName)

            Log-Event `
                -Arguments $Arguments `
                -EventInfo $LogEntries.ReseedSuccess `
                -Parameters $parameters
        }
        else
        {
            # Log Failure
            #
            [string[]]$parameters = ($Catalog.DatabaseName, "Active database copy not mounted")

            Log-Event `
                -Arguments $Arguments `
                -EventInfo $LogEntries.ReseedFailure `
                -Parameters $parameters
        }
    }
    catch
    {
        $reason = $error[0].Exception.ToString() + $error[0].InvocationInfo.PositionMessage
        
        # Log Failure
        #
        [string[]]$parameters = ($Catalog.DatabaseName, $reason)

        Log-Event `
            -Arguments $Arguments `
            -EventInfo $LogEntries.ReseedFailure `
            -Parameters $parameters
    }
    finally
    {
        # Revert back to original error action preference
        #
        $ErrorActionPreference=$errorPref
    }
}


<#
   .DESCRIPTION
   Stops a service on a local/remote server.
   This function stops the service, makes sure
   the process has been terminated. If not, 
   it tries to kill the process. 

   .PARAMETER Server
   The name of mailbox server
   
   .PARAMETER ServiceName
   The name of the service
   
   .PARAMETER ProcessName
   The name of the process behind the service
   
   .PARAMETER TimeoutMinutes
   Time in minutes to wait for stopping/terminating
   services. If processes could not be stopped within
   the timeout, a TimeoutException is raised.
   
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   None. Throws exception if not successful.   
#>
function Stop-SearchService
{
    Param(
        [String] 
        [ValidateNotNullOrEmpty()]
        $Server,
        [String]
        $ServiceName,
        [String]
        $ProcessName,
        [TimeSpan]
        $Timeout
    )

    Validate-Timeout $Timeout "Stop-SearchService"

    $deadline = (Get-Date) + $Timeout
    $stopServiceTimeout = new-timespan -seconds 120
    $stopProcessTimeout = new-timespan -seconds 60
        
    write-verbose ("Stop-SearchService, deadline = " + $deadline)
    
    # Keep stopping and terminating the process
    # until the process dies or we reach timeout
    #
    while ($deadline -gt (Get-Date))
    {
        $serviceStopped = Stop-UsingServiceController -Server $Server -ServiceName $ServiceName -Timeout $stopServiceTimeout
        if (!($serviceStopped -eq $true))
        {
            $processStopped = Stop-LocalOrRemoteProcess -Server $Server -ProcessName $ProcessName -Timeout $stopProcessTimeout
            if (!($processStopped -eq $true))
            {
                # When there's a deadlock, killing msftefd
                # seems to stop services faster.
                # 
                $msftefdStopped = Stop-LocalOrRemoteProcess -Server $Server -ProcessName $msftefdProcessName -Timeout $stopProcessTimeout
            }
        }
        else
        {
            break
        }
    }
    
    $newTimeout = $deadline - (get-date)
    Wait-ForProcessToStop -ProcessName $ProcessName -Server $Server -Timeout $newtimeout
    
    # If we are here, there was no timeout exception.
    #
    write-verbose ("Successfully stopped service " + $ServiceName + " within the timeout period.")
}

<#
   .DESCRIPTION
   Starts a service on a local/remote server.
   This function starts the service, makes sure
   the process has been running. 

   .PARAMETER Server
   The name of mailbox server
   
   .PARAMETER ServiceName
   The name of the service
   
   .PARAMETER ProcessName
   The name of the process behind the service
   
   .PARAMETER TimeoutMinutes
   Time in minutes to wait for starting the
   service. If processes could not be started within
   the timeout, a TimeoutException is raised.
   
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   None. Throws exception if not successful.   
#>
function Start-SearchService
{
    Param(
        [String] 
        [ValidateNotNullOrEmpty()]
        $Server,
        [String]
        [ValidateNotNullOrEmpty()]
        $ServiceName,
        [String]
        [ValidateNotNullOrEmpty()]
        $ProcessName,
        [TimeSpan]
        [ValidateNotNullOrEmpty()]
        $Timeout
    )

    Validate-Timeout $Timeout "Start-SearchService"

    $deadline = (Get-Date) + $Timeout
    
    # Find out if the service is already started.
    #
    $p = Get-WMIProcess -ProcessName $ProcessName -Server $Server
    if (!($p -eq $null))
    {
        write-verbose ("Service " + $ServiceName + " is already started.")
        write-verbose ("process : $p")
        return
    }

    write-verbose ("Starting service " + $ServiceName)
    write-verbose ("Start-SearchService deadline = " + $deadline)

    # Start the service
    #
    $service = new-Object System.ServiceProcess.ServiceController($ServiceName, $Server)
    $service.Start()
    
    $service.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Running, $Timeout)
    
    write-verbose ("Successfully started service " + $ServiceName + " within the timeout period.")
}

<#
   .DESCRIPTION
   Stops a service using the service controller 

   .PARAMETER Server
   The name of mailbox server
   
   .PARAMETER ServiceName
   The name of the service
      
   .PARAMETER Timeout
   Time to wait for starting the
   service. If processes could not be started within
   the timeout, a TimeoutException is raised.
   
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   Returns $true if service was stopped, $false otherwise   
#>
function Stop-UsingServiceController
{
    Param(
        [String] 
        [ValidateNotNullOrEmpty()]
        $Server,
        [String]
        [ValidateNotNullOrEmpty()]
        $ServiceName,
        [TimeSpan]
        [ValidateNotNullOrEmpty()]
        $Timeout
    )
    
    $deadline = (Get-Date) + $Timeout
    $stautsCheckIntervalSeconds = 5

    # Save the original value for 
    # error preference
    #
    $errorPref = $ErrorActionPreference
    try
    {
        Validate-Timeout $Timeout "Stop-UsingServiceController"
        $serviceFilter = ("name='" + $ServiceName + "'")
        write-verbose ("Service filter used for get-wmiobject = " + $serviceFilter)
        $ErrorActionPreference = "Continue"
        $service = get-wmiobject win32_service -filter $serviceFilter -ComputerName $Server
        if (!($service -eq $null))
        {
            write-verbose ("Stopping service " + $ServiceName + " on " + $Server + " using service controller")
            $service.StopService()
            
            # Check service status often
            # until we timeout or service
            # is stopped.
            #
            while ((!($service.State -ieq "Stopped")) -and
                   ($deadline -gt (get-date)))
            {
                Start-Sleep -seconds $stautsCheckIntervalSeconds
                
                $service = get-wmiobject win32_service -filter $serviceFilter -ComputerName $Server
                write-verbose ("Service state of " + $ServiceName + " on " + $Server + ":" + $service.State)
            }
            
            if ($service.State -ieq "Stopped")
            {
                write-verbose ("Stopped service " + $ServiceName + " on " + $Server + " using service controller")
                return $true
            }
        }
        else
        {
            write-verbose ("Could not get service object for service " + $ServiceName + " on computer " + $Server)
        }
    }
    finally
    {
        $ErrorActionPreference = $errorPref
    }
 
    write-verbose ("Could not stop service " + $ServiceName + " on " + $Server + " using service controller")
    return $false
}

<#
   .DESCRIPTION
   Terminates a process on a local/remote computer  

   .PARAMETER Server
   The name of computer
   
   .PARAMETER ProcessName
   The name of the process
   
   .PARAMETER Timeout
   Time to wait for the task to complete.
   A timeout exceptionis thrown if it doesn't   

   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   Returns $true if process was stopped, $false otherwise   
#>
function Stop-LocalOrRemoteProcess
{
    Param(
        [String] 
        [ValidateNotNullOrEmpty()]
        $Server,
        [String]
        [ValidateNotNullOrEmpty()]
        $ProcessName,
        [TimeSpan]
        [ValidateNotNullOrEmpty()]
        $Timeout        
    )

    $deadline = (Get-Date) + $Timeout
    $stautsCheckIntervalSeconds = 5

    # Save the original value for 
    # error preference
    #
    $errorPref = $ErrorActionPreference
    try
    {
        $ErrorActionPreference = "Continue"
        $process = Get-WMIProcess -ProcessName $ProcessName -Server $Server
        if (!($process -eq $null))
        {
            write-verbose ("Stopping process " + $ProcessName + " on " + $Server + " using remote WMI object")
            $process.Terminate(1)
            
            # Keep checking the status of the process
            # until it is the expected value or we
            # timeout.
            #
            while ((!($process -eq $null)) -and
                   ($deadline -gt (get-date)))
            {
                Start-Sleep -seconds $stautsCheckIntervalSeconds
                $process = Get-WMIProcess -ProcessName $ProcessName -Server $Server
            }
        }
        
        if ($process -eq $null)
        {
            write-verbose ("Process " + $ProcessName + " on computer " + $Server + " has stopped.")
            return $true
        }
        else
        {
            write-verbose ("Could not stop process " + $ProcessName + " on computer " + $Server)
        }
    }
    finally
    {
        $ErrorActionPreference = $errorPref
    }
    
    return $false
}

<#
   .DESCRIPTION
   Gets a WMI process object on a local/remote computer  

   .PARAMETER ProcessName
   The name of the process without the file extension 
   
   .PARAMETER Server
   The name of the server

   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   Returns the first process object that matches given filter   
#>
function Get-WMIProcess
{
    Param(
        [String] 
        [ValidateNotNullOrEmpty()]
        $ProcessName,
        [String]
        [ValidateNotNullOrEmpty()]
        $Server
    )
    
    # Note: When using WMI objects, we need to append '.exe'
    # for process names
    #
    $filter = ("name='" + $ProcessName + ".exe'")
    
    write-verbose ("Get-WMIObject win32_process -Filter " + $filter + " -Server " + $Server)    
    $processList = get-wmiobject win32_process -Filter $filter -ComputerName $Server
    if ($processList -is [Array])
    {
        $process = $processList[0]
    }
    else
    {
        $process = $processList
    }
    
    return $process
}

<#
   .DESCRIPTION
   Translates an event type string into an event type enum.
   
   
   .PARAMETER Type
   String representing the event type to translate
 
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   Enum representing the event type to be used by a monitoring event.   
#>
function Get-EventTypeEnum
{
    Param(
        [String] 
        [ValidateSet("Error", "Warning", "Information")]
        $Type
    )
        
    if ($Type -eq "Error")
    {
        return $EVENT_TYPE_ERROR     
    }

    if ($Type -eq "Warning")
    {
        return $EVENT_TYPE_WARNING     
    }
    
    if ($Type -eq "Information")
    {
        return $EVENT_TYPE_INFORMATION
    }
}

<#
   .DESCRIPTION
   Logs an event in the crimson log. In addition,
   if the event is not crimson-specific and monitoring
   context is $True, the event is logged to the 
   application log.
   
   "CI Troubleshooter" as the event source for application
   log and "CITS Operational" used for crimson log.

   .PARAMETER Arguments
   The Arguments object constructed from script args

   .PARAMETER EventInfo
   An object containing event id, type and message
      
   .PARAMETER Parameters
   Array of strings that comprise the parameters for the event
   These parameters must match the order and count specified
   in the message in the EventInfo paremeter
 
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None.   
#>
function Log-Event
{
    Param(
        [Object] 
        [ValidateNotNull()]
        $Arguments,
        [Object[]]
        [ValidateNotNullOrEmpty()]
        $EventInfo,
        [String[]]
        $Parameters
    )
    
    $id = $EventInfo[0]
    $type = $EventInfo[1]
    $message = $EventInfo[2]

    write-verbose ("Log-Event called with id=" + $id + " type=" + $type + " message=" + $message)
    # Replace all parameter substrings (%1, %2, etc)
    # with the real parameters
    # Start backwards so that %10 is replaced with the 10th parameter and not the first one.
    #
    for ($i = $Parameters.Length + 1; $i -gt 0; $i--)
    {
        $substring = "%$i"
        if ($message.Contains($substring))
        {
            $message = $message.Replace($substring, $Parameters[$i - 1])
        }
    }
    
    write-verbose $message
    
    $category = 1
    if ($Parameters -eq $null)
    {
        [string[]]$messageAndParams= @($message)
    }
    else
    {
        [string[]]$messageAndParams= @($message) + @($Parameters)
    }
    
    Write-CrimsonLogEntry `
        -Server $Arguments.Server `
        -EventId $id `
        -Category $category `
        -Type $type `
        -MessageAndParams $messageAndParams 
        
    # Log the event to application log
    # only if the id is not a crimson-specific
    # event and monitoring context is $True
    #
    if ($id -lt 6000) 
    {
        if ($Arguments.MonitoringContext -eq $True)
        {
            # Add the event into the pipeline, at the end of the TS 
            # call Write-MonitoringEvents to flush the events.
            $eventType = Get-EventTypeEnum -Type $type
            Add-MonitoringEvent `
                    -Id $id `
                    -Type $eventType `
                    -Message $message `
                    -InstanceName $Arguments.InstanceName
        }

        if ($Arguments.WriteApplicationEvent -eq $True)
        {
                Write-AppLogEntry `
                    -Server $Arguments.Server `
                    -EventSource $AppLogSourceName `
                    -EventId $id `
                    -Category $category `
                    -Type $type `
                    -MessageAndParams $messageAndParams
        }
    }
    else
    {
        write-verbose "MonitoringContext was false or the event is crimson-only, so skipped logging event"
    }
}

<#
   .DESCRIPTION
   Writes an entry to application event log,
   using global constants $appLogName and
   $appLogSourceName

   .PARAMETER Server
   The name of the server
   
   .PARAMETER EventId
   The id of the event to create

   .PARAMETER Category
   The category of the event

   .PARAMETER Type
   The type of the event. Allowed values are "Information", "Warning" and "Error"
   
   .PARAMETER MessageAndParams
   Array of strings comprising the message and its parameters.
   The first item in the array must be the message string.
 
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None.   
#>
function Write-AppLogEntry
{
    Param(
        [String]
        [ValidateNotNullOrEmpty()]
        $Server,
        [Int]
        $EventId,
        [Int]
        $Category,
        [String] 
        [ValidateSet("Error", "Warning", "Information")] 
        $Type = "Information",
        [String[]]
        [ValidateNotNullOrEmpty()]
        $MessageAndParams
    )
    
    Write-EventLogEntry `
        -Server $Server `
        -LogName $appLogName `
        -EventSource $appLogSourceName `
        -EventId $id `
        -Category $category `
        -Type $type `
        -MessageAndParams $messageAndParams
}

<#
   .DESCRIPTION
   Writes an entry to windows event 
   (crimson) log, using global constants
   $crimsonLogName and $crimsonLogSourceName

   .PARAMETER Server
   The name of the server
   
   .PARAMETER EventId
   The id of the event to create

   .PARAMETER Category
   The category of the event

   .PARAMETER Type
   The type of the event. Allowed values are "Information", "Warning" and "Error"
   
   .PARAMETER MessageAndParams
   Array of strings comprising the message and its parameters.
   The first item in the array must be the message string.
 
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None.   
#>
function Write-CrimsonLogEntry
{
    Param(
        [String]
        [ValidateNotNullOrEmpty()]
        $Server,
        [Int]
        $EventId,
        [Int]
        $Category,
        [String] 
        [ValidateSet("Error", "Warning", "Information")] 
        $Type = "Information",
        [String[]]
        [ValidateNotNullOrEmpty()]
        $MessageAndParams
    )
    
    Write-EventLogEntry `
        -Server $Server `
        -LogName $crimsonLogName `
        -EventSource $crimsonLogSourceName `
        -EventId $id `
        -Category $category `
        -Type $type `
        -MessageAndParams $messageAndParams
}

<#
   .DESCRIPTION
   Writes an entry to event log  

   .PARAMETER Server
   The name of the server
   
   .PARAMETER LogName
   The name of event log (Application/other)

   .PARAMETER EventSource
   The event source under which this entry should be logged.
   The source is created if it doesn't exist already.

   .PARAMETER EventId
   The id of the event to create

   .PARAMETER Category
   The category of the event

   .PARAMETER Type
   The type of the event. Allowed values are "Information", "Warning" and "Error"
   
   .PARAMETER MessageAndParams
   Array of strings comprising the message and its parameters.
   The first item in the array must be the message string.
 
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None.   
#>
function Write-EventLogEntry
{
    Param(
        [String]
        [ValidateNotNullOrEmpty()]
        $Server,
        [String]
        [ValidateNotNullOrEmpty()]
        $LogName,
        [String]
        [ValidateNotNullOrEmpty()]
        $EventSource,
        [Int]
        $EventId,
        [Int]
        $Category,
        [String] 
        [ValidateSet("Error", "Warning", "Information")] 
        $Type = "Information",
        [String[]]
        [ValidateNotNullOrEmpty()]
        $MessageAndParams
    )

    # First, make sure the event source is registered
    #
    if (!([System.Diagnostics.EventLog]::SourceExists($EventSource, $Server)))
    {
        write-verbose "Event source $EventSource doesn't exist on $Server. Creating it."
        
        try
        {        
            $creationData = 
                new-object System.Diagnostics.EventSourceCreationData($EventSource, $LogName)
            $creationData.MachineName = $Server
            [System.Diagnostics.EventLog]::CreateEventSource($creationData)
        }
        catch [System.Management.Automation.MethodInvocationException]
        {
            # Due to race condition between multiple invocations, if event source 
            # gets registered by one invocation, the other may fail with this error
            # trying to re-register the event source. In this case, let's clear
            # the error and continue the execution.
            #
            if ($error[0].InvocationInfo.InvocationName -eq "CreateEventSource")
            {
                $exception = $error[0].Exception
                write-verbose ("Expected exception when event source already exists: $exception")
                $error.Clear()
            }
        }
    }
   
    # Create the framework object for EventLog
    #
    $log = new-object System.Diagnostics.EventLog($LogName, $Server)
    $log.Source = $EventSource
    
    $eventInstance = new-object System.Diagnostics.EventInstance($EventId, $Category, $Type)
    $log.WriteEvent($eventInstance, $MessageAndParams)
}

<#
   .DESCRIPTION
   Waits for a specific event  

   .PARAMETER Server
   The name of the server
   
   .PARAMETER LogName
   The name of the log e.t."Application"

   .PARAMETER EventSource
   The event source under which this entry should be logged.
   The source is created if it doesn't exist already.

   .PARAMETER EventId
   The id of the event to create
   
   .PARAMATER StartTime
   The time from which to look for service started event.
   
   .PARAMETER Timeout
   The timespan of the wait. If the event was not logged
   within the given timespan, a timeout exception is thrown.

   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None.   
#>
function Wait-ForEvent
{
    Param(
        [String]
        [ValidateNotNullOrEmpty()]
        $Server,
        [String]
        [ValidateNotNullOrEmpty()]
        $LogName,
        [String]
        [ValidateNotNullOrEmpty()]
        $EventSource,
        [Int]
        $EventId,
        [DateTime]
        $StartTime,
        [TimeSpan]
        $Timeout
    )
    
    Validate-Timeout $Timeout "Wait-ForEvent"

    $start = $StartTime.ToUniversalTime().ToString("o")
    $query = "*[System[TimeCreated[@SystemTime > '$start'] and EventID='$EventId' and Provider[@Name='$EventSource']]]"
    
    $deadline = (Get-Date) + $Timeout
        
    write-verbose ("Wait-ForEvent deadline: " + $deadline)
    write-verbose ("Wait-ForEvent is using query: $query")
    
    $found = $False

    # Keep checking for the event until at least
    # one is found or timeout expires
    while ($deadline -gt (Get-Date))
    {
        try
        {
            $events = get-winevent  -LogName $LogName -ComputerName $Server -FilterXPath $query -ErrorAction "Stop"
            
            # we won't reach here if the above
            # didn't find any events
            #
            $found = $True
            write-verbose "Found the following events:"
            if ($events -is [Array])
            {
                foreach ($event in $events)
                {
                    write-verbose $event.Message
                }
            }
            else
            {
                write-verbose $events.Message
            }
        }
        catch
        {
            $exception = $error[0].Exception
            write-verbose ("Error retrieving requested event from event log: $exception")
        }
        
        if (!($found))
        {            
            Start-Sleep -Seconds 5
            write-verbose "Did not find event. checking again.."
        }
        else
        {
            break
        }
    }
    
    if (!($found))
    {
        $msg = ($LocStrings.TimeoutWaitingForEvent + $EventId)
        write-verbose -Message $msg
        throw (new-object -typename System.TimeoutException($msg))
    }
}

<#
   .DESCRIPTION
   Waits for a process to completely stop  
   Used to wait for the process to go away
   after killing it or stopping the service

   .PARAMETER ProcessName
   The name of the process (without the file extension)

   .PARAMETER Server
   The name of the server

   .PARAMETER Timeout
   The timespan of the wait. If the process was not stopped
   within the given timespan, a timeout exception is thrown.
   
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None.   
#>
function Wait-ForProcessToStop
{
    Param(
        [String]
        [ValidateNotNullOrEmpty()]
        $ProcessName,
        [String]
        [ValidateNotNullOrEmpty()]
        $Server,
        [TimeSpan]
        $Timeout
    )
    
    Validate-Timeout $Timeout "Wait-ForProcessToStop"
    
    $deadline = (Get-Date) + $Timeout
        
    write-verbose ("Wait-ForProcessToStop deadline: " + $deadline)
    
    $stopped = $False

    while ($deadline -gt (get-date))
    {
        $p = Get-WMIProcess -ProcessName $ProcessName -Server $Server 
        if ($p -eq $null)
        {
            write-verbose "Process $ProcessName successfully stopped."
            $stopped = $True
            break
        }
        else
        {
            Start-Sleep -Seconds 5
            write-verbose "Process $ProcessName not stopped yet. Checking again.."
        }
    }
    
    if (!$stopped)
    {
        $msg = ($LocStrings.TimeoutWaitingForProcessToStop + $ProcessName)
        write-verbose -Message $msg
        throw (new-object -typename System.TimeoutException($msg))
    }
}

<#
   .DESCRIPTION
   Validates the timeout and throws
   TimeoutException if it is negative.  

   .PARAMETER $Timeout
   The timeout value
   
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None.   
#>
function Validate-Timeout
{
    param(
        [TimeSpan]
        $Timeout,
        [String]
        [ValidateNotNullOrEmpty()]
        $Caller
    )
    
    if ($Timeout.TotalMilliseconds -le 0.0)
    {
        $msg = ($LocStrings.TimeoutZeroOrNegative + $Caller)
        write-verbose -Message $msg
        throw (new-object -typename System.TimeoutException($msg))
    }
}

<#
   .DESCRIPTION
   Returns True if another instance of
   Troubleshooter is already running.  
   
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None.   
#>
function Is-AnotherInstanceRunning
{
    # Get the last event logged by Troubleshooter
    # in the crimson log
    #
    $eventId = $LogEntries.TSStarted[0]
    $filterXPath = "*[System[EventID='$eventId' and Provider[@Name='$crimsonLogSourceName']]]"
    write-verbose "Filter used for getting first start event is $filterXPath"
    $start = get-winevent -LogName $crimsonLogName -FilterXPath $filterXPath -MaxEvents 1 -ErrorAction:SilentlyContinue
    
    # if we have no TSStarted event, no
    # other instance is running
    #
    if ($start -eq $null)
    {
        return $False
    }
    
    # if we have a corresponding finished/failed event
    # there is no other instance running.
    # 
    $timeCreated = $start.TimeCreated
    $startTime = $timeCreated.ToUniversalTime().ToString("o")
    $filterXPath = "*[System[TimeCreated[@SystemTime >= '$startTime'] and Provider[@Name='$crimsonLogSourceName']]]"
    write-verbose "Filter used for getting matching finish/fail event is $filterXPath"

    $finish = get-winevent -LogName $crimsonLogName -FilterXPath $filterXPath | `
              where {($_.id -eq $LogEntries.TSSuccess[0]) -or ($_.id -eq $LogEntries.TSFailed[0])}
                  
    if ($finish -eq $null)
    {
        return $True
    }
}

<#
   .DESCRIPTION
   Loads Exchange Powershell Snapin,
   if not already loaded.  
   
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None.   
#>
function Load-ExchangeSnapin
{
    if (! (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:SilentlyContinue) )
    {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
    }
}

$exchangeInstallPath = (get-item -path env:ExchangeInstallPath).Value
$exchangeScriptPath = join-path $exchangeInstallPath "Scripts"

#
# Get the default English strings first
#
Import-LocalizedData -BindingVariable LocStrings -FileName CITSLibrary.strings.psd1

#
# If the localized strings are available, then use them
#
Import-LocalizedData -BindingVariable LocStrings -BaseDirectory $exchangeScriptPath -FileName CITSLibrary.strings.psd1 -ErrorAction:SilentlyContinue

# Event log entry dictionary
#
$LogEntries = @{
#
# Events logged to application log and windows event (crimson) log
# Information: 5000-5299; Warning: 5300-5599; Error: 5600-5999;
#
#   Informational events
#
    TSStarted=(5000,"Information",$LocStrings.TSStarted)
    TSSuccess=(5001,"Information", $LocStrings.TSSuccess)
    DetectedNoIssues=(5002,"Information", $LocStrings.DetectedNoIssues)
    CatalogHasNoIssues=(5003, "Information", $LocStrings.CatalogHasNoIssues)
    RestartSuccess=(5004,"Information", $LocStrings.RestartSuccess)
    ReseedSuccess=(5005,"Information", $LocStrings.ReseedSuccess)
#
#   Warning events
#
    DetectedDeadlock=(5300,"Warning",$LocStrings.DetectedDeadlock)
    DetectedCatalogCorruption=(5301,"Warning", $LocStrings.DetectedCatalogCorruption)
    DetectedIndexingStall=(5302,"Warning", $LocStrings.DetectedIndexingStall)
#
#   Error events
#
    TSFailed=(5600,"Error", $LocStrings.TSFailed)
    DetectedSameSymptomTooManyTimes=(5601,"Error", $LocStrings.DetectedSameSymptomTooManyTimes)
    RestartFailure=(5602,"Error", $LocStrings.RestartFailure)
    ReseedFailure=(5603,"Error", $LocStrings.ReseedFailure)
    DetectedIndexingBacklog=(5604,"Error", $LocStrings.DetectedIndexingBacklog)
    AnotherInstanceRunning=(5605,"Error",$LocStrings.AnotherInstanceRunning)
#
#   Events logged only to crimson (windows) event log
#   Information: 6000-6299; Warning: 6300-6599; Error: 6600-5999;
#
#
#   Informational events
#
    TSDetectionStarted=(6000, "Information", $LocStrings.TSDetectionStarted)
    TSDetectionFinished=(6001, "Information", $LocStrings.TSDetectionFinished)
    TSResolutionStarted=(6002, "Information", $LocStrings.TSResolutionStarted)
    TSResolutionFinished=(6003, "Information", $LocStrings.TSResolutionFinished)
#
#   Warning events
#

#
#   Error events
#
    TSDetectionFailed=(6600, "Error", $LocStrings.TSDetectionFailed)
    TSResolutionFailed=(6601, "Error", $LocStrings.TSResolutionFailed)
    }


# SIG # Begin signature block
# MIIdngYJKoZIhvcNAQcCoIIdjzCCHYsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUqIICdolcAImRPJc9E/6f0xNa
# n+egghhkMIIEwzCCA6ugAwIBAgITMwAAAJqamxbCg9rVwgAAAAAAmjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI5
# WhcNMTcwNjMwMTkyMTI5WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkIxQjctRjY3Ri1GRUMyMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEApkZzIcoArX4o
# w+UTmzOJxzgIkiUmrRH8nxQVgnNiYyXy7kx7X5moPKzmIIBX5ocSdQ/eegetpDxH
# sNeFhKBOl13fmCi+AFExanGCE0d7+8l79hdJSSTOF7ZNeUeETWOP47QlDKScLir2
# qLZ1xxx48MYAqbSO30y5xwb9cCr4jtAhHoOBZQycQKKUriomKVqMSp5bYUycVJ6w
# POqSJ3BeTuMnYuLgNkqc9eH9Wzfez10Bywp1zPze29i0g1TLe4MphlEQI0fBK3HM
# r5bOXHzKmsVcAMGPasrUkqfYr+u+FZu0qB3Ea4R8WHSwNmSP0oIs+Ay5LApWeh/o
# CYepBt8c1QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFCaaBu+RdPA6CKfbWxTt3QcK
# IC8JMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAIl6HAYUhsO/7lN8D/8YoxYAbFTD0plm82rFs1Mff9WBX1Hz
# /PouqK/RjREf2rdEo3ACEE2whPaeNVeTg94mrJvjzziyQ4gry+VXS9ZSa1xtMBEC
# 76lRlsHigr9nq5oQIIQUqfL86uiYglJ1fAPe3FEkrW6ZeyG6oSos9WPEATTX5aAM
# SdQK3W4BC7EvaXFT8Y8Rw+XbDQt9LJSGTWcXedgoeuWg7lS8N3LxmovUdzhgU6+D
# ZJwyXr5XLp2l5nvx6Xo0d5EedEyqx0vn3GrheVrJWiDRM5vl9+OjuXrudZhSj9WI
# 4qu3Kqx+ioEpG9FwqQ8Ps2alWrWOvVy891W8+RAwggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TCCBhAwggP4
# oAMCAQICEzMAAABkR4SUhttBGTgAAAAAAGQwDQYJKoZIhvcNAQELBQAwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xNTEwMjgyMDMxNDZaFw0xNzAx
# MjgyMDMxNDZaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29ycG9yYXRpb24w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCTLtrY5j6Y2RsPZF9NqFhN
# FDv3eoT8PBExOu+JwkotQaVIXd0Snu+rZig01X0qVXtMTYrywPGy01IVi7azCLiL
# UAvdf/tqCaDcZwTE8d+8dRggQL54LJlW3e71Lt0+QvlaHzCuARSKsIK1UaDibWX+
# 9xgKjTBtTTqnxfM2Le5fLKCSALEcTOLL9/8kJX/Xj8Ddl27Oshe2xxxEpyTKfoHm
# 5jG5FtldPtFo7r7NSNCGLK7cDiHBwIrD7huTWRP2xjuAchiIU/urvzA+oHe9Uoi/
# etjosJOtoRuM1H6mEFAQvuHIHGT6hy77xEdmFsCEezavX7qFRGwCDy3gsA4boj4l
# AgMBAAGjggF/MIIBezAfBgNVHSUEGDAWBggrBgEFBQcDAwYKKwYBBAGCN0wIATAd
# BgNVHQ4EFgQUWFZxBPC9uzP1g2jM54BG91ev0iIwUQYDVR0RBEowSKRGMEQxDTAL
# BgNVBAsTBE1PUFIxMzAxBgNVBAUTKjMxNjQyKzQ5ZThjM2YzLTIzNTktNDdmNi1h
# M2JlLTZjOGM0NzUxYzRiNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzcitW2oynUC
# lTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# b3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEGCCsGAQUF
# BwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0MAwGA1Ud
# EwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjiDGRDHd1crow7hSS1nUDWvWas
# W1c12fToOsBFmRBN27SQ5Mt2UYEJ8LOTTfT1EuS9SCcUqm8t12uD1ManefzTJRtG
# ynYCiDKuUFT6A/mCAcWLs2MYSmPlsf4UOwzD0/KAuDwl6WCy8FW53DVKBS3rbmdj
# vDW+vCT5wN3nxO8DIlAUBbXMn7TJKAH2W7a/CDQ0p607Ivt3F7cqhEtrO1Rypehh
# bkKQj4y/ebwc56qWHJ8VNjE8HlhfJAk8pAliHzML1v3QlctPutozuZD3jKAO4WaV
# qJn5BJRHddW6l0SeCuZmBQHmNfXcz4+XZW/s88VTfGWjdSGPXC26k0LzV6mjEaEn
# S1G4t0RqMP90JnTEieJ6xFcIpILgcIvcEydLBVe0iiP9AXKYVjAPn6wBm69FKCQr
# IPWsMDsw9wQjaL8GHk4wCj0CmnixHQanTj2hKRc2G9GL9q7tAbo0kFNIFs0EYkbx
# Cn7lBOEqhBSTyaPS6CvjJZGwD0lNuapXDu72y4Hk4pgExQ3iEv/Ij5oVWwT8okie
# +fFLNcnVgeRrjkANgwoAyX58t0iqbefHqsg3RGSgMBu9MABcZ6FQKwih3Tj0DVPc
# gnJQle3c6xN3dZpuEgFcgJh/EyDXSdppZzJR4+Bbf5XA/Rcsq7g7X7xl4bJoNKLf
# cafOabJhpxfcFOowMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkqhkiG9w0B
# AQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAG
# A1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTEw
# HhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
# aWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
# q/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03a8YS2Avw
# OMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akrrnoJr9eW
# WcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0RrrgOGSsbmQ1
# eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy4BI6t0le
# 2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9sbKvkjh+
# 0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAhdCVfGCi2
# zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8kA/DRelsv
# 1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTBw3J64HLn
# JN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmnEyimp31n
# gOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90lfdu+Hgg
# WCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0wggHpMBAG
# CSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2oynUClTAZ
# BgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/
# BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBaBgNVHR8E
# UzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9k
# dWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsGAQUFBwEB
# BFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNVHSAEgZcw
# gZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsGAQUFBwIC
# MDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABlAG0AZQBu
# AHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKbC5YR4WOS
# mUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11lhJB9i0ZQ
# VdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6I/MTfaaQ
# dION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0wI/zRive
# /DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560STkKxgrC
# xq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQamASooPoI/
# E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGaJ+HNpZfQ
# 7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ahXJbYANah
# Rr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA9Z74v2u3
# S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33VtY5E90Z1W
# Tk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr/Xmfwb1t
# bWrJUnMTDXpQzTGCBKQwggSgAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBuDAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUpWU6bDjYLqMb0/7XADKA8qIfG8AwWAYKKwYB
# BAGCNwIBDDFKMEigIIAeAEMASQBUAFMATABpAGIAcgBhAHIAeQAuAHAAcwAxoSSA
# Imh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEB
# BQAEggEAjsXIgCvNxhQpjqJd5lVkN0mtUFFw+BWLV3i2vD2OBB9EONPVudJ+vFc8
# SbHKq7oM690qiEz1rggD5rAS+UlRknpzQZu+aPcctpDZRceifn0takAZbKD1jCeb
# xPxA90sZ7Kh7u3LTy0ebwkwLUAd75QXMV78VqjlR3RtPk8vyLETBtjjLEPsAIqQH
# DOGT1wiRYRtJlatJ8/Yk/C4mpHrlUlnVQ/vSKBvzv1ty0DB6HC/FvgYdx4b2iymz
# eztjqMcE7qtQllAkk/TeJtO2qUi7uGRNP5jO1QwKCiNT5TzN9T+VasgQ+Q9sJdd5
# L13lxNtH0oqXjgKOeYUlQBRrjZ+wv6GCAigwggIkBgkqhkiG9w0BCQYxggIVMIIC
# EQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAACampsWwoPa1cIA
# AAAAAJowCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJ
# KoZIhvcNAQkFMQ8XDTE2MDkwMzE4NDQ1MlowIwYJKoZIhvcNAQkEMRYEFANNpwds
# YQPnN9IHcQ9uKv+QsBJsMA0GCSqGSIb3DQEBBQUABIIBAGKHYePi+DS+52X04WKH
# EkJ5ZLR0AslAXWshLa5nMxYN25br53XY1Q7aAosFDqJmi9kd6C6vT0jItCa/YQLt
# C3vFvnGzxOLqIzNOZGbSufKDGf4wjiVsleUGd1f+UBSQQ/G/AMbavSo2tJX3GlIe
# sDwB65bTtj+a1z6qsqhDKcniQ37jdfymDMYFIbLXsTQZEPibgDxXb/IKjYfHcGeO
# k+W5MKq5vxBiRCPbq6frgM/sfHoDZKrSiccBJVV7ZkdiKhYQXM2WOeJrxWzicAvI
# GK3+s673JaRpy15v2vbas0b1NQ60rD0t68kcz9CUWbDhk7+kDUSCFGkmoXv3GONo
# GNA=
# SIG # End signature block
