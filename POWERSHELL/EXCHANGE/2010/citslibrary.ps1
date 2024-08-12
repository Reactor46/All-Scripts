# Copyright (c) 2009 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
# This file contains Content Index Troubleshooter functions
#

#####################################################################################
#
#
# THIS FILE EXISTS IN TWO LOCATIONS. MAKE SURE TO BOTH COPIES OF THE FILE ARE UPDATED WHEN 
# WHEN EITHER COPY IS CHANGED
# <DEPOT>\Sources\dev\management\src\management\scripts\troubleshooter\CITSLibrary.ps1
# <DEPOT>\Sources\dev\mgmtpack\src\HealthMainfests\scripts\troubleshooter\CITSLibrary.ps1
# The management version of the library gets deployed during exchange setup and the 
# mgmtpack version of the library only gets deployed when the management pack is installed
#
######################################################################################

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
        $FailureTimeSpanMinutes,
        [bool]
        $CanTakeProcessCrashDumps
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
    $Arguments.CanTakeProcessCrashDumps = $CanTakeProcessCrashDumps
    $value = Get-RegKeyValue -Path $troubleshooterRegKey -Name 'TroubleshooterDisabled' -DefaultValue 0
    if ($value -gt 0)
    {
        $Arguments.TroubleshooterDisabled = $true
    }
    
    $Arguments.CanTakeProcessCrashDumps = -not [bool]::Parse((Get-RegKeyValue -Server $Arguments.Server -Path $troubleshooterRegKey -Name 'DisableCrashDump' -DefaultValue 'False'))
    
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
   
   .PARAMETER Symptom
   Symptom to detect.
      
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
        $After,
        [string]
        [ValidateNotNullOrEmpty()]
        $Symptom
    )
    
    $serverStatus = new-object -typename ServerStatus
    $serverStatus.Name = $Server
    $serverStatus.IsDeadlocked = $false
    $serverStatus.IsRTMServer = Get-IsRTMServer -Server $server
    
    # Run the memory check only if 'All' or 'MsftefdHealth' is specified as a symptom
    #
    if ($Symptom -ieq "MsftefdHealth" -or $Symptom -ieq "All")
    {
        $serverStatus.CumulativeMsftefdMemoryConsumption = Get-MsftefdMemoryUsage $server
        $serverStatus.CatalogStatusArray = @()
        $serverStatus.IsMsftefdMemoryConsumptionOverLimit = ($MaxMsftefdMemoryConsumption -gt 0) -and ($serverStatus.CumulativeMsftefdMemoryConsumption -gt $MaxMsftefdMemoryConsumption)
    }
    
    $serverStatus.BadIFilters = @(Get-BadIFilters -Server $server -IsRTMServer $serverStatus.IsRTMServer)
    $serverStatus.IFiltersToEnable = @(Get-IFiltersToEnable -Server $server -IsRTMServer $serverStatus.IsRTMServer)
    
    # Run these checks when we dont want to check msftefd health only
    if ($Symptom -ine "MsftefdHealth")
    {
        $copyArray = @(Get-Copies -Server $Server -Database $Database)
        
        $ciStatusArray = @(Get-CIStatus -copies $copyArray)
        
        if ($After -eq $null -or $After -eq [DateTime]::MinValue)
        {
            $startTime = (get-date).AddMinutes(-1 * $badDiskBlockCheckIntervalInMinutes)
        }
        else
        {
            $startTime = $After
        }
        
        write-verbose "checking bad block issues only after $startTime"
        
        $ciStatusArray = @(Check-BadDiskBlocks -Server $Server -StartTime $startTime -CIStatusArray $ciStatusArray)

        foreach ($status in $CIStatusArray)
        {
            if (IsCatalogStalled -CIStatus $status -StallThresholdInSeconds $stallThresholdInSeconds)
            {
                write-verbose ("Catalog " + $status.Name + " is stalled.")
                $status.IsStalled = $true
                if ($status.StallCounter -gt $extendedStallThresholdInSeconds)
                {
                    write-verbose ("Catalog " + $status.Name + " is stalled for an extended period of " + $extendedStallThresholdInSeconds/(60*60*24) + " days")
                    $status.IsStalledExtendedPeriod = $true
                }
            }
            
            if (IsCatalogBacklogged -CIStatus $status -BacklogThresholdInSeconds $backlogThresholdInSeconds -RetryItemsThreshold $retryItemsThreshold)
            {
                write-verbose ("Catalog " + $status.Name + " is backlogged.")
                $status.isBacklogged = $true
            }
            
            if (IsCatalogHealthStale -CIStatus $status -StaleThresholdInSeconds $staleThresholdInSeconds)
            {
                write-verbose ("Catalog health for " + $status.Name + " is stale.")
                $status.IsHealthStale = $true
            }
            
            if (IsLargeCatalog -CIStatus $status -PercentThreshold $maxPercentageCatalogSize)
            {
                write-verbose ("Percentage Catalog size for " + $status.Name + " is greater than the allowed threshold.")
                $status.IsLargeCatalog = $true
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
            
            # If the catalog has a bad block on the disk, set the IsCorrupted to true
            # so that the corresponding resolution action (reseed) can be taken.
            #
            if ($status.BadDiskBlockMasterMerge)
            {
                write-verbose ("Catalog " + $status.Name + " has bad disk block during master merge.")
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
                
        $serverStatus.CatalogStatusArray = $ciStatusArray
        
        Check-ForRetryQueueIssues -CatalogStatusArray $serverStatus.CatalogStatusArray -RetryItemsThreshold $retryItemsThreshold
        
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
        [ValidateSet("Deadlock", "Corruption", "Stall", "MsftefdHealth")] 
        $Symptom
    )
  
    $serverStatus = new-object -typename ServerStatus
    $serverStatus.Name = $Server
    $serverStatus.CatalogStatusArray = @()
    
    if ($Symptom -ieq "Deadlock")
    {
        $serverStatus.IsDeadlocked = $True
    }
    elseif($Symptom -ieq "MsftefdHealth")
    {
        $serverStatus.IsMsftefdMemoryConsumptionOverLimit = $True
    }
    else
    {
        $copyArray = Get-Copies -Server $Server -Database $Database
        $ciStatusArray = Get-CIStatus -copies $copyArray
        $serverStatus.IsDeadlocked = $False
        $serverStatus.CatalogStatusArray = $ciStatusArray
        $serverStatus.BadIFilters = @()

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
        Log-Event -Arguments $Arguments -EventInfo $LogEntries.DetectedDeadlock -Parameters @("SomeString")
    }
    
    
    if ($ServerStatus.IsMsftefdMemoryConsumptionOverLimit)
    {
        $ProcessInstances = ''
        foreach($process in @(Get-Process -name $msftefdProcessName -ComputerName $Arguments.Server -ErrorAction:SilentlyContinue))
        {
            if ($process -eq $null)
            {
                continue
            }

            $ProcessInstances += [String]::Format("Id: {0}, MemoryUsage(MB): {1}`r`n", $process.Id , $process.PrivateMemorySize64/(1024*1024))
        }
        
        Log-Event -Arguments $Arguments `
                -EventInfo $LogEntries.MsftefdMemoryUsageHigh `
                -Parameters @($Arguments.MaxCumulativeMsftefdMemoryConsumption, $ServerStatus.CumulativeMsftefdMemoryConsumption, $ProcessInstances)
        $issuesFound = $true
    }
    
    if ($serverStatus.BadIFilters -ne $null -and $serverStatus.BadIFilters.Length -gt 0)
    {
        [string]$filterNames = ""
        foreach($filterName in $serverStatus.BadIFilters)
        {
            $filterNames += "'$filterName',"
        }
        
        # Remove the additional ',' from the end of the filterNames string
        #
        $filterNames = $filterNames.SubString(0, $filterNames.Length - 1)
        $issuesFound = $true
        Log-Event -Arguments $Arguments -EventInfo $LogEntries.FoundBadIFiltersEnabled -Parameters @($filterNames)
    }
    
    if ($serverStatus.IFiltersToEnable -ne $null -and $serverStatus.IFiltersToEnable.Length -gt 0)
    {
        [string]$filterNames = ""
        foreach($filterName in $serverStatus.IFiltersToEnable)
        {
            $filterNames += "'$filterName',"
        }
        
        # Remove the additional ',' from the end of the filterNames string
        #
        $filterNames = $filterNames.SubString(0, $filterNames.Length - 1)
        $issuesFound = $true
        Log-Event -Arguments $Arguments -EventInfo $LogEntries.IFiltersToEnable -Parameters @($filterNames)
    }
    
    $backLoggedCatalogs = ""
    $catalogWithNoIssues = ""
    $catalogsWithRetryQueueIssues = ""
    $backloggedThresholdHours = $backlogThresholdInSeconds/3600
    foreach ($catalog in $ServerStatus.CatalogStatusArray)
    {
        if ($catalog -eq $null)
        {
            continue
        }
        
        if ($catalog.IsStalled)
        {
            $issuesFound = $True
            Log-Event -Arguments $Arguments -EventInfo $LogEntries.DetectedIndexingStall -Parameters @($catalog.DatabaseName, $catalog.StallCounter, $stallThresholdInSeconds)
            if ($catalog.IsStalledExtendedPeriod)
            {
                Log-Event `
                    -Arguments $Arguments `
                    -EventInfo $LogEntries.DetectedIndexingStallExtendedPeriod `
                    -Parameters @($catalog.DatabaseName, $catalog.StallCounter, $extendedStallThresholdInSeconds)
            }
        }
        elseif ($catalog.IsCorrupted)
        {
            $issuesFound = $True
            Log-Event -Arguments $Arguments -EventInfo $LogEntries.DetectedCatalogCorruption -Parameters @($catalog.DatabaseName)
        }
        elseif ($catalog.IsBacklogged)
        {
            $issuesFound = $True
            [string[]]$parameters = ($Catalog.DatabaseName, $backloggedThresholdHours, $retryItemsThreshold)
            Log-Event -Arguments $Arguments -EventInfo $LogEntries.DetectedIndexingBacklog -Parameters $parameters
            $catalogStatus = [String]::Format("{0} ({1}, {2}, {3})", $Catalog.DatabaseName, $Catalog.BacklogCounter, $Catalog.NumberOfItemsInRetryQueue, $Catalog.NumberOfRetryItemsProcessed)
            $backLoggedCatalogs = $backLoggedCatalogs + $catalogStatus

            if ($catalog.HasRetryQueueIssues)
            {
                # Skipping logging of events per catalog. Will be logging one event for all the catalogs on the server
                # [string[]]$parameters = ($Catalog.DatabaseName, $catalog.NumberOfRetryItemsProcessed, 0)
                # Log-Event -Arguments $Arguments -EventInfo $LogEntries.ItemsStuckInRetryQueue -Parameters $parameters
                $catalogsWithRetryQueueIssues = $catalogsWithRetryQueueIssues + $catalogStatus
            }
        }
        elseif ($catalog.IsLargeCatalog)
        {
            $issuesFound = $True
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $LogEntries.CatalogSizeGreaterThanExpectedDBLimit `
                -Parameters @($catalog.DatabaseName, $maxPercentageCatalogSize, $catalog.PercentageCatalogSize)
        }
        else
        {
            if ($catalog.IsCorrupted -eq $false)
            {
                # This particular copy was found healthy in this troubleshooter run. Clear off any state stored by the TS associated with catalog
                #
                Reset-EventRetryCounter -Server $Arguments.Server -EventId $LogEntries.ActiveCatalogCopyCorrupt[0] -OptionalComponent $Catalog.DatabaseName
                Reset-EventRetryCounter -Server $Arguments.Server -EventId $LogEntries.CatalogReseedLoop[0] -OptionalComponent $Catalog.DatabaseName
                Reset-EventRetryCounter -Server $Arguments.Server -EventId $LogEntries.ReseedFailure[0] -OptionalComponent $Catalog.DatabaseName
            }

            $catalogWithNoIssues = $catalogWithNoIssues + $Catalog.DatabaseName + [System.Environment]::NewLine
        }       
    }
    
    if ([System.String]::IsNullOrEmpty($catalogWithNoIssues) -eq $false)
    {
        Log-Event -Arguments $Arguments -EventInfo $LogEntries.CatalogHasNoIssues -Parameters @($catalogWithNoIssues)
    }
    
    if ([System.String]::IsNullOrEmpty($catalogsWithRetryQueueIssues) -eq $false)
    {
        Log-Event -Arguments $Arguments -EventInfo $LogEntries.RetryQueuesStagnant -Parameters @($catalogsWithRetryQueueIssues)
    }
    
    if ([System.String]::IsNullOrEmpty($backLoggedCatalogs) -eq $false)
    {
        [string[]]$parameters = ($backLoggedCatalogs, $backloggedThresholdHours, $retryItemsThreshold)
        Log-Event -Arguments $Arguments -EventInfo $LogEntries.DetectedIndexingBacklogOrLargeRetryQueuesOnMultipleDatabases -Parameters $parameters
    }
    
    # If no issues are found, log the fact.
    # This will help turn previous alerts
    # green.
    #
    if (!($issuesFound))
    {
        Log-Event -Arguments $Arguments -EventInfo $LogEntries.DetectedNoIssues -Parameters @("SomeString")
    }
}

<#
   .DESCRIPTION
   Gets the memory usage of all the filter MSDTED processes

   .PARAMETER Server
   Name of the server to monitor the process 
         
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   None.
#>

function Get-MsftefdMemoryUsage
{
    Param(
    [string]
    $Server = $env:ComputerName
    )
    
    if ($Server -ieq "localhost")
    {
        $Server = $env:ComputerName    
    }
    
    [long]$privateBytes = (@(Get-Process -name $msftefdProcessName -ComputerName $Server -ErrorAction:SilentlyContinue ) | measure -Property PrivateMemorySize64 -Sum).Sum
    [long]$privateMbs = $privateBytes/(1024*1024)
    return $privateMbs
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
            -EventInfo $LogEntries.TSResolutionStarted `
            -Parameters @("SomeString")

        $restartServices = $false
        $stalled = $False
        $includeCrashDump = $false
        $crashDumpProcessNames = @()
        $crashDumpProcessId = -1
        $additionalRestartContext = ""
        # Are any catalogs stalled?
        #
        foreach ($catalog in $ServerStatus.CatalogStatusArray)
        {
            if ($catalog.IsStalled -eq $True)
            {
                write-verbose ($catalog.Name + " is stalled.")
                $stalled = $True
                $restartServices = $true
                $additionalRestartContext += $catalog.Name + " is stalled."
                break
            }
            
            if ($catalog.IsBacklogged -and $catalog.HasRetryQueueIssues)
            {
                write-verbose ($catalog.Name + " retry queue is stalled.")
                $stalled = $True
                $restartServices = $true
                $additionalRestartContext += $catalog.Name + " retry queue is stalled."
                break
            }
        }
        
        if ($ServerStatus.IsDeadlocked -or $stalled)
        {
            write-verbose ("Detected indexing stalls or a deadlock. Restarting search services on " + $Arguments.Server)
            $crashDumpProcessNames += $msftesqlProcessName
            $crashDumpProcessNames += $exsearchServiceName
            $additionalRestartContext += " Detected indexing stalls or a deadlock."
            $restartServices = $true
            $includeCrashDump = $true
        }
        
        if ($ServerStatus.IsMsftefdMemoryConsumptionOverLimit)
        {
            write-verbose ("Msftefd memory consumption over limit. Restarting search services on " + $Arguments.Server)
            $additionalRestartContext += " Msftefd memory consumption over limit."
            $msftefdProcesses = @(Get-Process -name $msftefdProcessName -ComputerName $Arguments.Server -ErrorAction:SilentlyContinue)
            $highestMemoryUsage = 512
            foreach($msftefdProcess in $msftefdProcesses)
            {
                $memoryUsageInMb = $msftefdProcess.PrivateMemorySize64/(1024*1024)
                if ($memoryUsageInMb -gt $highestMemoryUsage)
                {
                    $crashDumpProcessId = $msftefdProcess.Id
                }
            }
            
            $restartServices = $true
            $includeCrashDump = $true
        }
        
        if ($disableBadIFilters -gt 0)
        {
            if ($serverStatus.BadIFilters -ne $null -and $serverStatus.BadIFilters.Length -gt 0)
            {
                foreach($filterName in $serverStatus.BadIFilters)
                {
                    Disable-BadIFilter -Server $Arguments.Server -FilterName $filterName
                }
            
                $restartServices = $true
            }
        
            if ($serverStatus.IFiltersToEnable -ne $null -and $serverStatus.IFiltersToEnable.Length -gt 0)
            {
                foreach($filterName in $serverStatus.IFiltersToEnable)
                {
                    if ((Enable-IFilter -Server $Arguments.Server -FilterName $filterName) -eq $true)
                    {
                        $restartServices = $true
                        # Once the disabled IFilters have been enabled we should reset the Retry counter
                        # This will ensure that the TS does not enable an IFilter for the next 6 runs (6 Hours)
                        # If enabling any one of the IFilters fails in this pass the TS will try again in the next 
                        # scheduled window (6 hours later) and not in the next run to avoid restarting services
                        # frequently in Datacenter. (Enabling an IFilter is low priority compared to other operations)
                        #
                        Reset-EventRetryCounter -Server $Arguments.Server -EventId $msftesqlBadIFilterEventIdEventId
                    }
                }
            }
        }
        
        if ($restartServices)
        {
            if ($includeCrashDump)
            {
                Get-ProcessDump -Arguments $Arguments -crashDumpProcessId $crashDumpProcessId -crashDumpProcessNames $crashDumpProcessNames
                
                # Only log this error if we took a crash dump for the FD memory consumption issue
                #
                if ($ServerStatus.IsMsftefdMemoryConsumptionOverLimit)
                {
                    Log-Event -Arguments $Arguments `
                        -EventInfo $LogEntries.MsftefdMemoryUsageHighWithCrashDump `
                        -Parameters @($Arguments.MaxCumulativeMsftefdMemoryConsumption, $ServerStatus.CumulativeMsftefdMemoryConsumption)
                }
            }
            
            $restartCount = [int](Get-RegKeyValue `
                                -Server $server `
                                -Path ($troubleshooterRegKey + $LogEntries.ServiceRestartAttempt[0].ToString()) `
                                -Name 'CurrentCount' `
                                -DefaultValue 0)
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $LogEntries.ServiceRestartAttempt `
                -Parameters @($restartCount.ToString())
            $additionalRestartContext += Get-ServerStatusString $ServerStatus
            Restart-SearchServices -Arguments $Arguments -Timeout $defaultRestartTimeout -AdditionalContext $additionalRestartContext
        }
        else
        {
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $LogEntries.ServiceRestartNotNeeded `
                -Parameters @("SomeString")
            Reset-EventRetryCounter -Server $Arguments.Server -EventId $LogEntries.ServiceRestartAttempt[0]
        }
            
        # now look for corruptions and start reseeding
        # each corrupted catalog
        #
        foreach ($catalog in $ServerStatus.CatalogStatusArray)
        {
            if ($catalog -eq $null)
            {
                continue
            }
            
            if ((ShouldIgnoreRecovery -DatabaseName $catalog.DatabaseName))
            {
                Log-Event `
                    -Arguments $Arguments `
                    -EventInfo $LogEntries.CatalogRecoveryDisabled `
                    -Parameters @($catalog.DatabaseName)
                continue
            }
        
            if ($catalog.IsCorrupted -eq $True)
            {
                write-verbose ($catalog.Name + " seems to be corrupted. Reseeding the catalog..")
                Reseed-Catalog -Arguments $Arguments -Catalog $catalog
            }
        }
        
        # To avoid the MSFTESQL Process from consuming a lot of CPU we limit processor affinity.
        # The processor affinity count will be read from a registry. 
        # Make sure that this line appears after the restart services call otherwise the changes would be lost once the process restarts
        #
        [void](Set-ProcessorAffinity -ProcessName $msftesqlProcessName -NumberOfCPU $affinityValue)
        
        # Log success event
        #
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $LogEntries.TSResolutionFinished `
            -Parameters @("SomeString")
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
   Checks if Recovery actions for the catalog of a mailbox database should be ignored on this server

   .PARAMETER DatabaseName
   Name of the database being checked
         
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   $true if the recovery should be ignored. $False otherwise
#>
function ShouldIgnoreRecovery
{
    Param($DatabaseName)    
    if ([String]::IsNullOrEmpty($DatabaseName) -eq $false -and $disableRecoveryForDatabasesList -ne $null)
    {
        foreach($databaseToIgnore in $disableRecoveryForDatabasesList)
        {
            if ($databaseToIgnore -ne $null -and $databaseToIgnore -ieq $DatabaseName)
            {
                return $true
            }
        }
    }
    
    return $false
}

<#
   .DESCRIPTION
   Sets the affinity of a process to the last N number of CPU's in the system.
   The reason why we pick "last" is - usually the first a few CPUs are heavily used already.

   .PARAMETER ProcessName
   Name of the process
   
   .PARAMETER NumberOfCPU
   Total Number of CPU's
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   $true if the value was set successfully or $false otherwise
#>
function Set-ProcessorAffinity
{
    Param(
        [Object]
        [ValidateNotNullOrEmpty()]
        $ProcessName,
        [int]
        $NumberOfCPU
    )
    
    $process = Get-Process -Name $ProcessName -ErrorAction:SilentlyContinue
    if ($process -eq $null)
    {
        return $false
    }

    # Get the number of logical processor count
    $logicalProcessorCount = 0
    $processors = @(Get-WmiObject Win32_Processor)
    foreach ($processor in $processors)
    {
        $logicalProcessorCount += $processor.NumberOfLogicalProcessors
        $processor.Dispose()
    }
    
    # If we have only 1 processor, there is no point setting affinity.
    if ($logicalProcessorCount -le 1)
    {
        return $false
    }
    
    # If we don't have a meaningful value from caller for $NumberOfCPU, we will use 1/3 of total logical processors
    # If there are only 2 logical processors, we will use 1.
    if ($NumberOfCPU -lt 1)
    {
        $NumberOfCPU = [Math]::Floor($logicalProcessorCount / 3)
        if ($NumberOfCPU -eq 0)
        {
            $NumberOfCPU = 1
        }
    }
    
    $processorAffinityValue = 0
    for($power = 0; $power -lt $NumberOfCPU; $power++)
    {
        # Set the mask to use the last N processor.
        $processorAffinityValue += [Math]::Pow(2, $logicalProcessorCount - 1 - $Power)
    }
    
    if ($processorAffinityValue -gt 0)
    {
        try
        {
            if ([int]$process.ProcessorAffinity -ne $processorAffinityValue)
            {
                $process.ProcessorAffinity = new-object IntPtr $processorAffinityValue
            }

            return $true
        }
        catch [System.Exception]
        {
            # Do nothing 
            $message=($error[0].Exception.ToString() + $error[0].InvocationInfo.PositionMessage)
            write-verbose ("Caught Exception: $message")
        }
    }

    return $false
}


<#
   .DESCRIPTION
   Attempts to take a crash dump of a process

   .PARAMETER Arguments
   Object of type Arguments, containing command-line
   arguments 

   .PARAMETER crashDumpProcessId
   Unique PID of the process.
   
   .PARAMETER crashDumpProcessNames
   List containing the process names
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   None   
#>
function Get-ProcessDump
{
    Param(
        [Object]
        [ValidateNotNullOrEmpty()]
        $Arguments,
        [int]
        $crashDumpProcessId,
        $crashDumpProcessNames
    )
    
    try
    {
        if ((Check-CanTakeCrashDump -Server $Arguments.Server) -eq $true)
        {
            if ($crashDumpProcessId -gt 0)
            {
                Update-LastCrashDumpTime -Server $Arguments.Server
                .\dump-process.ps1 -uniquePid $crashDumpProcessId -Alias exsearch@microsoft.com -numDumps 1 -dfs -Full
            }
            else
            {
                foreach($processName in $crashDumpProcessNames)
                {
                    Update-LastCrashDumpTime -Server $Arguments.Server
                    .\dump-process.ps1 -processname $processName -Alias exsearch@microsoft.com -numDumps 1 -dfs 
                }
            }
        }
    }
    catch [System.Exception]
    {
        # Catch any exceptions thrown by the dump process script
        # and ignore them. Process dump collection is not a critical piece
        #
        $message=($error[0].Exception.ToString() + $error[0].InvocationInfo.PositionMessage)
        write-verbose ("Caught Exception: $message")            
    }
}

<#
   .DESCRIPTION
   Checks if the TS can take a crash dump of a process. Function will return true only if $minTimeBetweenCrashDumps 
   time criteria is met

   .PARAMETER Server
   The simple NETBIOS name of mailbox server.

   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   None.   
#>
function Check-CanTakeCrashDump
{
    Param(
        [String] [AllowNull()]
        $Server
    )
    
    [DateTime]$crashDumpTime = [System.DateTime]::MinValue
    $value = Get-RegKeyValue -Server $Server -Path $troubleshooterRegKey -Name "LastCrashDumpTime"
    if ($value -ne $null)
    {
        $crashDumpTime = [DateTime]$value
    }
    
    if (($crashDumpTime + $minTimeBetweenCrashDumps) -lt (Get-Date))
    {
        return $true
    }
    
    return $false
}

<#
   .DESCRIPTION
   Updates the time when a crash dump was taken by the TS to the current time

   .PARAMETER Server
   The simple NETBIOS name of mailbox server.
     
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None.   
#>
function Update-LastCrashDumpTime
{
    Param(
        [String] [AllowNull()]
        $Server
    )
    
    Set-RegKeyValue -Server $server -Path $troubleshooterRegKey -Name "LastCrashDumpTime" -Value (Get-Date)
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
        [ValidateNotNull()]
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
        $noirt = Get-CachedCounter -Value ($hashKeyPrefix + $noirtCounterName)
        $nomltc = Get-CachedCounter -Value ($hashKeyPrefix + $nomltcCounterName)
        $noirfrt = Get-CachedCounter -Value ($hashKeyPrefix + $noirfrtCounterName)

        $ciStatus = new-object -typename CIStatus
        $ciStatus.Name = $copy.Name
        $ciStatus.DatabaseName = $copy.DatabaseName
        $ciStatus.DatabaseGuid = Get-DatabaseGuid -Database $copy.DatabaseName
        $ciStatus.BacklogCounter  = $aolni
        $ciStatus.NumberOfItemsInRetryQueue  = $noirt
        $ciStatus.NumberOfMailboxesLeftToCrawl = $nomltc
        $ciStatus.NumberOfRetryItemsProcessed = Get-RetryDocumentsProcessedSinceLastRun `
            -CurrentRetryTableItemsProcessed $noirfrt `
            -DatabaseCopy $copy
        $ciStatus.StallCounter = $tslni
        $ciStatus.Health = $copy.ContentIndexState
                
        $ciHealth = Get-CatalogHealth -Server $copy.MailboxServer -Database $ciStatus.DatabaseName
        $ciStatus.HealthReason = $ciHealth.ErrorCode
        $ciStatus.HealthTimestamp = $ciHealth.Timestamp

        $ciStatus.PercentageCatalogSize = Get-PercentageCatalogSize -DatabaseCopy $copy
 
        # Initialize detection flags
        #
        $ciStatus.IsStalled = $false
        $ciStatus.IsBacklogged = $false
        $ciStatus.IsCorrupted = $false
        $ciStatus.IsHealthStale = $false
        $ciStatus.HasBadDiskBlock = $false
        $ciStatus.BadDiskBlockMasterMerge = $false
        $statusList += $ciStatus
    }
    
    return $statusList
}

<#
    .DESCRIPTION
    Gets the number of retry documents processed since the last troubleshooter run

    .PARAMETER databaseCopy
    Database copy 

    .PARAMETER CurrentRetryTableItemsProcessed
    Database copy 
         
    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    Number of documents processed since last run
#>
function Get-RetryDocumentsProcessedSinceLastRun
{
    Param(
        [Object]
        [ValidateNotNull()]
        $DatabaseCopy,
        [int]
        $CurrentRetryTableItemsProcessed
    )
    
    $returnValue = $CurrentRetryTableItemsProcessed
    $lastRetryTableDocumentProcessedCountRegkeyPath = $troubleshooterRegKey + 'NumberOfRetryQueueItemsProcessed'
    $lastRetryTableDocumentProcessedCount = [int](Get-RegKeyValue `
                                                    -Server $server `
                                                    -Path $lastRetryTableDocumentProcessedCountRegkeyPath `
                                                    -Name $DatabaseCopy.DatabaseName `
                                                    -DefaultValue 0)
    if ($lastRetryTableDocumentProcessedCount -le $CurrentRetryTableItemsProcessed)
    {
        $returnValue = $CurrentRetryTableItemsProcessed - $lastRetryTableDocumentProcessedCount
    }
    
    Set-RegKeyValue `
                -Server $server `
                -Path $lastRetryTableDocumentProcessedCountRegkeyPath `
                -Name $DatabaseCopy.DatabaseName `
                -Value $CurrentRetryTableItemsProcessed
    
    return $returnValue
}

<#
    .DESCRIPTION
    Tries to gets the management pack version deployed on the current server

    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    Management pack version deployed on the server
#>
function Get-ManagementPackVersion
{
    $location = Get-Location
    $version = "MP version not found"
    try
    {
        # Check if the Operations Manager Client snapin is present on the server
        #
        $operationsManagerClientSnapin = Get-PSSnapin -Registered | ?{$_.Name -ieq 'Microsoft.EnterpriseManagement.OperationsManager.Client'}
        if ($operationsManagerClientSnapin -ne $null)
        {
            # TODO the ideal way would be to get the current server from the OpsMgrConnector config file.
            #

            $managementGroupServerList = @()
            # Get the Name of the management group server that this computer is a part of
            #
            $OperationsManagerRegistryPath = "SOFTWARE\Microsoft\Microsoft Operations Manager\3.0"

            $managementGroups = Get-RegKeySubKeyNames -Server $Env:ComputerName -Path "$OperationsManagerRegistryPath\Agent Management Groups"
            foreach($managementGroup in $managementGroups)
            {
                if ($managementGroup -ne $null)
                {
                    $managementGroupServer = Get-RegKeyValue -Server $Env:ComputerName -Path "$OperationsManagerRegistryPath\Agent Management Groups\$managementGroup\Parent Health Services\0" -Name "NetworkName" -DefaultValue $null
                    if ($managementGroupServer -ne $null)
                    {
                        $managementGroupServerList += $managementGroupServer
                    }
                }
            }

            $managementGroupServer = Get-RegKeyValue -Server $Env:ComputerName -Path "$OperationsManagerRegistryPath\Machine Settings" -Name "DefaultSDKServiceMachine" -DefaultValue $null
            if ($managementGroupServer -ne $null)
            {
                $managementGroupServerList += $managementGroupServer
            }

            foreach($managementGroupServer in $managementGroupServerList)
            {
                # Try connecting to the Management group server to get the exchange management pack version
                #
                Add-PSSnapin 'Microsoft.EnterpriseManagement.OperationsManager.Client' -ErrorAction:SilentlyContinue
                Set-Location "OperationsManagerMonitoring::"            
                $Script:MG = New-ManagementGroupConnection -ConnectionString $managementGroupServer
                if ($Script:MG -eq $null)
                {
                    continue
                }

                $exchangeManagementPack = Get-ManagementPack | ?{$_.Name -ieq 'Microsoft.Exchange.2010'}
                if ($exchangeManagementPack -ne $null)
                {
                    return $exchangeManagementPack.Version.ToString()
                }
            }
        }
    }
    catch
    {
        write-verbose ("Failed to query the management pack verion." + $Error[0].ToString())
    }
    finally
    {
        if ($location -ne $null)
        {
            Set-Location $location
        }
    }
    
    return $version
}


<#
    .DESCRIPTION
    Gets the catalog size as a percentage of the overall database size

    .PARAMETER databaseCopy
    Database copy 
         
    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    Catalog size as a percentage of the overall database size   
#>
function Get-PercentageCatalogSize
{
    Param(
        [Object]
        [ValidateNotNull()]
        $DatabaseCopy
    )

    $percentCatalogSize = 0

    # Percentage catalog size is a best effort calculation. The troubleshooter will fail if it encounters any problems getting the catalog data
    try
    {
        $mailboxDatabase = Get-CachedMailboxDatabase -DatabaseName $DatabaseCopy.DatabaseName
        $catalogDirectory = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($mailboxDatabase.EdbFilePath), "CatalogData-" + $mailboxDatabase.Guid.ToString() + "-*");
        $catalogSizeMb = ((gci $catalogDirectory -recurse | measure-object Length -sum).Sum / (1024 * 1204))
        $edbSizeMb = ((gci $mailboxDatabase.EdbFilePath).Length / (1024 * 1204))

        # Don't bother with size checks if the MDB is really small
        if ($edbSizeMb -ge 1024)
        {
            $percentCatalogSize = ($catalogSize * 100) / $edbSize
        }
    }
    catch [System.Exception]
    {
        $percentCatalogSize = -1
    }

    return $percentCatalogSize
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
        [ValidateNotNull()]
        $CIStatusArray
    )
    
    Check-MasterMergeDiskCorruptions -Server $server -StartTime $startTime -CIStatusArray $CIStatusArray
    
    if ($CIStatusArray.Length -eq 0)
    {
        return $CIStatusArray
    }
    
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
        $msftesqlCrashes = @(get-eventlog -computername $Server -after $StartTime -logname "Application" -source $msftesqlServiceName -ErrorAction:SilentlyContinue) | where {$_.eventId -eq $msftesqlCrashEventId}
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
        $badDiskEvents = @(get-eventlog -computername $Server -after $StartTime -logname "System" -source $diskSourceName -ErrorAction:SilentlyContinue) | where {$_.eventId -eq $badDiskEventId}
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
                $databaseName = Map-DiskNumberToDatabase -DiskNumber $diskNumber -Server $server
                
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

   .PARAMETER Server
   Name of the server to lookup. This function does not work for a remote server yet because it uses the
   diskpart utility which works only on local computers
   
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
        [String] 
        [ValidateNotNullOrEmpty()]
        $Server,
        [Int32] 
        $DiskNumber
    )
    
    if ($server -ine $env:ComputerName)
    {
        write-verbose "Map-DiskNumberToDatabase does not work for remote computers"
        return $null
    }

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
    
    $databases = @(Get-CachedMailboxDatabase -Server $server | ?{$_.EdbFilePath -like "$MountPoint*"})
    
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
   Checks the application event log for MSSearch master merge failures because
   of disk errors and reports those catalogs as corrupt.
   
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
function Check-MasterMergeDiskCorruptions
{
    Param(
        [String] 
        [ValidateNotNullOrEmpty()]
        $Server,
        [DateTime]
        $StartTime,  
        [Object[]]
        [ValidateNotNull()]
        $CIStatusArray
    )
    
    try
    {
        $masterMergeErrors = @(get-eventlog `
                -computername $Server `
                -after $StartTime `
                -logname "Application" `
                -source $msftesqlServiceName `
                -ErrorAction:SilentlyContinue) | where {$_.eventId -eq $msftesqlMasterMergeFailedBadDiskEventId}
    }
    catch [System.Exception]
    {
        $masterMergeErrors = $null
    }
                
    if ($masterMergeErrors -eq $null -or $masterMergeErrors.Length -eq 0)
    {
        return
    }
    
    # Scan bad disk events, and get the unique bad disk names.
    #
    $badCatalogDatabaseGuids=@()
    foreach ($event in $masterMergeErrors)
    {
        $mdbGuid = Get-DatabaseGuidFromCatalogName -EventLogMessage $event.Message
        
        # If the mdb guid is null then skip this event message
        #
        if ($mdbGuid -eq $null)
        {
            continue
        }
        
        $found = $false
        foreach($badCatalogDbGuid in $badCatalogDatabaseGuids)
        {
            if ($badCatalogDbGuid -ieq $mdbGuid)
            {
                $found = $true
            }
        }
        if ($found -eq $false)
        {
            $badCatalogDatabaseGuids += $mdbGuid
        }
    }
    
    if ($badCatalogDatabaseGuids.Length -eq 0)
    {
        return
    }
    
    foreach($catalog in $CIStatusArray)
    {
        foreach($badCatalogDbGuid in $badCatalogDatabaseGuids)
        {
            if($badCatalogDbGuid -ieq $catalog.DatabaseGuid.ToString())
            {
                # Mark the Catalog for that database corrupt
                #
                $catalog.BadDiskBlockMasterMerge = $true
            }
        }
    }
}

function Get-DatabaseGuidFromCatalogName
{
    Param(
        [String] 
        [ValidateNotNullOrEmpty()]
        $EventLogMessage
    )
    
    
    # Event message extracted from bug# 349577:
    # A master merge has been paused for catalog ExSearch-ac4031a0-5982-4ddf-9967-d34b5bc1fb75-c907a6bd-8553-4d39-93c3-c97cdb23605a due to error The request could not be performed because of an I/O device error.   0x8007045d. 
    # It will be rescheduled later.
    
    # Use the regex below to extract the mdb Guid from the event message. ExSearch-{MDBGUID}-{InstanceGuid}

    
    $mdbguidExtractorRegexString = "A master merge has been paused for catalog ExSearch-(?<MdbGuid>[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}).*The request could not be performed because of an I/O device error..*0x8007045d."
    $mdbguidExtractorRegex = new-object System.Text.RegularExpressions.RegEx($mdbguidExtractorRegexString, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    
    $matches = $mdbguidExtractorRegex.Match($EventLogMessage)
    if ($matches.Success)
    {
        return $matches.Groups["MdbGuid"].Value 
    }
    
    return $null
    
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
    
    Populate-MailboxDatabaseObjects -Server $server
    if ([System.String]::IsNullOrEmpty($Database))
    {
        $dbs = @(Get-CachedMailboxDatabase -Server $Server)
    }
    else
    {
        $dbs = @(Get-CachedMailboxDatabase -DatabaseName $Database)
    }
    
    $dbCount = $dbs.Length
    write-verbose "Verifying $dbCount database(s) on $Server"
    
    $copies = @()
    foreach($db in $dbs)
    {
        $mailboxDatabaseCopyStatus = @(Get-MailboxDatabaseCopyStatus $db)
        $includeMdbForAnalysis = $false
        
        $copyName = "$db\$Server"
        # Check if the database has a mounted copy on any server in the DAG
        foreach($copyId in $mailboxDatabaseCopyStatus)
        {
            if ($copyId.Status -ieq 'Mounted')
            {
                $includeMdbForAnalysis = $true
                break;
            }
        }
        
        if ($includeMdbForAnalysis)
        {
            foreach($copyId in $mailboxDatabaseCopyStatus)
            {
                # Check if the copy of the database on the server is either Healthy or Mounted
                if ($copyId.Name -ieq $copyName -and ($copyId.Status -ieq 'Mounted' -or $copyId.Status -ieq 'Healthy'))
                {
                    $copies += $copyId
                }
            }
        }
        else
        {
            write-verbose "The troubleshooter will ignore $db because does not have any Active mounted copies"
        }
    }
    
    return $copies
}


<#
    .DESCRIPTION
    Gets the cached copy of the mailbox database object

    .PARAMETER DatabaseName
    Mailbox database Name

    .Parameter ServerName
    NetBiosName of the Mailbox server
         
    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    Cached copy of the mailbox database
#>
function Get-CachedMailboxDatabase
{
    [CmdletBinding(DefaultParameterSetName='Default')]
    Param(
        [String]
        [Parameter(ParameterSetName='Default')]
        $DatabaseName,
        
        [String]
        $Server
    )

    if ([System.String]::IsNullOrEmpty($DatabaseName))
    {
        $databaseList = @(Get-MailboxDatabasesOnServer -Server $Server)

        if ($databaseList.Count -eq 0)
        {
            Populate-MailboxDatabaseObjects -Server $Server
            $databaseList = @(Get-MailboxDatabasesOnServer -Server $Server)
        }

        return $databaseList
    }
    else
    {
        if ($mailboxDatabaseHashTable.ContainsKey($DatabaseName) -eq $false)
        {
            $mailboxDatabaseHashTable[$DatabaseName] = Get-MailboxDatabase $DatabaseName -Status
        }

        return $mailboxDatabaseHashTable[$DatabaseName]
    }
}

<#
    .DESCRIPTION
    Gets the cached copy of all the mailbox database objects on a server

    .Parameter Server
    NetBiosName of the Mailbox server
         
    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    Cached copy of the mailbox database
#>
function Get-MailboxDatabasesOnServer
{
    Param(
        [String]
        $Server
    )

    $databaseList = @()
    foreach($db in $mailboxDatabaseHashTable.Values)
    {
        if ($db -ne $null -and $db.Servers -ne $null)
        {
            foreach($serverObject in $db.Servers)
            {
                if ($serverObject.Name -ieq $server)
                {
                    $databaseList += $db
                    break;
                }
            }
        }
    }

    return $databaseList
}

<#
   .DESCRIPTION
   Gets mailbox database objects on a server and stores them
   in the database object cache.

   .PARAMETER Server
   The simple NETBIOS name of mailbox server.

   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   An array of database copy objects.
#>
function Populate-MailboxDatabaseObjects
{
    Param(
        [String]
        $Server
    )

    if ([System.String]::IsNullOrEmpty($Server))
    {
        return;
    }

    $dbs = get-mailboxdatabase -Server $Server -Status
    foreach($db in $dbs)
    {
        if ($db -ne $null)
        {
            $mailboxDatabaseHashTable[$db.Name] = $db 
        }
    }
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
   
   .PARAMETER DatabaseName
   Guid of the database.
      
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
        $DatabaseName
    )
 
    write-verbose "In function Get-CatalogHealthRegKey"
    $DatabaseGuid = Get-DatabaseGuid $DatabaseName
    return (Get-HealthFromRegistry -Server $Server -DatabaseGuid $DatabaseGuid)    
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
    
    write-verbose ("Get-CachedMailboxDatabase " + $Database)
    $db = Get-CachedMailboxDatabase -DatabaseName $Database
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
        [ValidateNotNull()]
        $CIStatusArray
    )

    # CI is deemed deadlocked if either of these 
    # conditions happen:
    #
    # 1. All catalog health timestamps are stale 
    # 2. All catalogs are stalled
    #
    
    if ($CIStatusArray.Count -eq 0)
    {
        return $false
    }
    
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
    
    $isStalled = $false
    if ($CIStatus.StallCounter -ge $StallThresholdInSeconds)
    {
        $isStalled = $true
    }

    $numberOfMailboxesLeftToCrawlRegkeyPath = $troubleshooterRegKey + 'MailboxesLeftToCrawl'
    # If the catalog is in the crawling state check if crawl has not stalled
    #
    if ($CIStatus.NumberOfMailboxesLeftToCrawl -gt 0)
    {
        # Get the last value of the Mailboxes left to crawl counter
        #
        $lastCrawlMailboxCount = Get-RegKeyValue -Server $server -Path $numberOfMailboxesLeftToCrawlRegkeyPath -Name $CIStatus.DatabaseName
        
        # Check if the stored value is greater than 0 and is not equal to the current value of NumberOfMailboxesLeftToCrawl.
        # We should only do the equality check, because its possible for the service to start re-crawling all the mailboxes on the server between two runs 
        # in which case the new counter value will be greater than the saved registry counter
        #
        if ($lastCrawlMailboxCount -gt 0 -and $CIStatus.NumberOfMailboxesLeftToCrawl -eq $lastCrawlMailboxCount)
        {
            [REF]$outMessage = ""
            $alertCritical = Check-EventThresholdReached -Server $Arguments.Server -EventId $stallDuringCrawlThreshold -Parameters @($CIStatus.DatabaseName) -Message ([REF]$outMessage)
            if ($alertCritical)
            {
                # The number of mailboxes left to crawl has not changed for $stallDuringCrawlThreshold TS runs. Assume that the Service has stalled and restart it
                #
                write-verbose ("Indexing stalled detected. The number of mailboxes left to crawl counter has not reduced between two consequetive TS runs")
                $isStalled = $true
            }
        }
        else
        {
            # The mailboxes left to crawl counters are different reset the StallDuringCrawlThreshold counter for that database
            #
            Reset-EventRetryCounter -EventId $stallDuringCrawlThreshold -OptionalComponent $CIStatus.DatabaseName
        }
    }
    
    # Save the NumberOfMailboxesLeftToCrawl counter value to the registry for future comparision
    #    
    Set-RegKeyValue -Server $server -Path $numberOfMailboxesLeftToCrawlRegkeyPath -Name $CIStatus.DatabaseName -Value $CIStatus.NumberOfMailboxesLeftToCrawl
    return $isStalled
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
        $BacklogThresholdInSeconds,
        [Int32]
        $RetryItemsThreshold
    )
 
    $isBacklogged = $false 
    if ($CIStatus.BacklogCounter -ge $BacklogThresholdInSeconds)
    {
        $isBacklogged = $true
    }

    if ($CIStatus.NumberOfItemsInRetryQueue -ge $RetryItemsThreshold)
    {
        $isBacklogged = $true
    }
    
    return $isBacklogged
}

<#
    .DESCRIPTION
    Determines if the retry queues for a catalog are draining over time. 

    .PARAMETER CatalogStatusArray
    Catalog status objects for all the catalog
    
    .PARAMETER RetryItemsThreshold
    The minimum number of items that should be present in the retry queue before the troubleshooter assumes there are issues with the retry queue    
    
    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    $true if a backlog was detected, $false otherwise   
#>
function Check-ForRetryQueueIssues
{
    Param($CatalogStatusArray,
        [Int32]
        $RetryItemsThreshold)
    
    write-verbose ("Check-ForRetryQueueIssues")
    $totalRetryDocumentsProcessed = 0
    $hasCatalogsWithLargeRetryQueues = $false
    
    # Get the total number of retry documents processed on the server
    foreach($catalogStatus in $CatalogStatusArray)
    {
        $totalRetryDocumentsProcessed = $totalRetryDocumentsProcessed + $catalogStatus.NumberOfRetryItemsProcessed
        
        # If any one of the catalogs has retry queues greater than the threshold value we should check if the service is processing events in the retry table
        # or ignore this check completly
        if ($catalogStatus.NumberOfItemsInRetryQueue -ge $RetryItemsThreshold)
        {
            $hasCatalogsWithLargeRetryQueues = $true
        }
    }
    
    $RetryIssuesConsecutiveRuns = 0
    $RetryIssuesCatalogsRegKeyPath = $troubleshooterRegKey + 'BackloggedCatalogs'
    $RetryIssuesCatalogsRegValueName = 'RetryQueueIssuesCount'
   
    if ($hasCatalogsWithLargeRetryQueues -and ($totalRetryDocumentsProcessed -le 0))
    {
        # If the number of documents processed by the retry feeder on the server is 0, check if the feeder 
        # has been stalled consequetively for $minRetryTableIssueThreshold before assuming issues with the retry queues
        $lastRetryIssuesCount = [int](Get-RegKeyValue `
                                            -Server $server `
                                            -Path $RetryIssuesCatalogsRegKeyPath `
                                            -Name $RetryIssuesCatalogsRegValueName `
                                            -DefaultValue -1)
        $anyCatalogHasRetryQueueIssues = $false
        $minRetryTableThresholdReached = $lastRetryIssuesCount -ge $minRetryTableIssueThreshold
        if ($minRetryTableThresholdReached)
        {
            foreach($catalogStatus in $CatalogStatusArray)
            {
                $catalogStatus.HasRetryQueueIssues = $catalogStatus.isBacklogged
            }
        }
        
        $RetryIssuesConsecutiveRuns = $lastRetryIssuesCount + 1
    }
    
    Set-RegKeyValue `
                -Server $server `
                -Path $RetryIssuesCatalogsRegKeyPath `
                -Name $RetryIssuesCatalogsRegValueName `
                -Value $RetryIssuesConsecutiveRuns
}

<#
    .DESCRIPTION
    Determines if the size of a catalog is greater than the allowed threshold

    .PARAMETER CIStatus
    Catalog status object for the catalog
         
    .INPUTS
    None. You cannot pipe objects to this function.

    .OUTPUTS
    $true if health status was stale, $false otherwise   
#>
function IsLargeCatalog
{
    Param(
        [Object]
        [ValidateNotNull()]
        $CIStatus,
        [Int32]
        $PercentThreshold
    )
    
    write-verbose ("$CIStatus.PercentageCatalogSize = " + $CIStatus.PercentageCatalogSize)
    return $CIStatus.PercentageCatalogSize -ge $PercentThreshold
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
   Method to get a formatted string containg the important properties of the DatabaseCopyStatusEntry object
   
   .PARAMETER DatabaseCopies
   Array of DatabaseCopyStatusEntry objects
   
   .OUTPUTS
   Formatted string containg the important object properties.   
#>
function Get-DatabaseCopyStatusString
{
    Param(
        [Object[]] 
        [ValidateNotNull()]
        $DatabaseCopies
    )
    
    $DatabaseCopies = @($DatabaseCopies)
    $catalogStatusString = "Name, Status, ReplayQueueLength, CopyQueueLength, ContentIndexState, ContentIndexErrorMessage" + [System.Environment]::NewLine 
    foreach($databaseCopy in $DatabaseCopies)
    {
        $catalogStatusString += [System.String]::Format( `
            "'{0}', {1}, {2}, {3}, {4}, '{5}'", `
            $databaseCopy.Name, `
            $databaseCopy.Status, `
            $databaseCopy.ReplayQueueLength, `
            $databaseCopy.CopyQueueLength, `
            $databaseCopy.ContentIndexState, `
            $databaseCopy.ContentIndexErrorMessage) + [System.Environment]::NewLine
    }
    
    return $catalogStatusString
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
        $problemdb = Get-CachedMailboxDatabase -DatabaseName $Catalog.DatabaseName
        if ($problemdb.Mounted)
        {
            $sourceServer = $null
            $allCopiesCorrupt = $true
            # Before attempting to reseed check if the catalog is not mounted
            #
            $databaseCopies = @(Get-MailboxDatabaseCopyStatus $Catalog.DatabaseName)
            foreach($databaseCopy in $databaseCopies)
            {
                if (($databaseCopy.Status -ieq 'Healthy' -or $databaseCopy.Status -ieq 'Mounted') -and $databaseCopy.ContentIndexState -ieq 'Healthy')
                {
                    # Found at least one healthy catalog. Use that as the source of the reseed
                    #
                    $allCopiesCorrupt = $false
                    if ($databaseCopy.MailboxServer -ine $Arguments.Server)
                    {
                        $sourceServer = $databaseCopy.MailboxServer
                    }
                }
                
                if ($databaseCopy.MailboxServer -ieq $Arguments.Server)
                {
                    $catalogCopy = $databaseCopy
                }
            }
            
            $isPassiveCopy = Try-FailoverCorruptCatalog -DatabaseCopy $catalogCopy -Catalog $Catalog
            
            if ($isPassiveCopy)
            {
                if ($allCopiesCorrupt -and (IsCatalogCorrupted -CIStatus $Catalog) -eq $false)
                {
                    # If we do not have any healthy copies and this copy is crawling then do not bother with the reseed. Let the passive catalog crawl
                    # But make sure the TS logs a Reseed Failure Error log with an Exception explaining that all catalogs are corrupt. 
                    #
                    throw (new-object -typename System.InvalidOperationException("Cannot reseed a catalog that does not have any healthy copies"))
                }

                # If the TS reached here then at this point we have a catalog that needs to be deleted and reseeded. If all the copies are corrupt 
                # this operation will force the catalog to start recrawling.
                #
                Update-CatalogCopy -CatalogName $Catalog.Name -SourceServer $sourceServer
                
                # Log success event
                #
                [string[]]$parameters = @($Catalog.DatabaseName)
                Reset-EventRetryCounter -Server $Arguments.Server -EventId $LogEntries.ReseedFailure[0] -OptionalComponent $Catalog.DatabaseName
                Log-Event `
                    -Arguments $Arguments `
                    -EventInfo $LogEntries.ReseedSuccess `
                    -Parameters $parameters
                
                # Reseed succeeded Check if this particular catalog was found to be corrupted in the last troubleshooter run
                #
                [REF]$outMessage = ""
                $Catalog.InReseedLoop = Check-EventThresholdReached `
                        -Server $Server `
                        -EventId $TSRetrySettings.CatalogReseedLoop[0] `
                        -Parameters @($Catalog.DatabaseName) `
                        -Message ([REF]$outMessage)

                if ($Catalog.InReseedLoop)
                {
                    Log-Event `
                        -Arguments $Arguments `
                        -EventInfo $LogEntries.CatalogReseedLoop `
                        -Parameters @($Catalog.DatabaseName, $TSRetrySettings.CatalogReseedLoop[1])
                }
             }
        }
        else
        {
            # Log Failure
            #
            [string[]]$parameters = ($Catalog.DatabaseName, "Active database copy not mounted", "Database dismounted")

            Log-Event `
                -Arguments $Arguments `
                -EventInfo $LogEntries.ReseedFailure `
                -Parameters $parameters
        }
    }
    catch
    {
        $reason = $error[0].Exception.ToString() + $error[0].InvocationInfo.PositionMessage
        if ($databaseCopies -ne $null)
        {
            $catalogStatusString = Get-DatabaseCopyStatusString $databaseCopies
        }
        else
        {
            $catalogStatusString = "Error getting database copy state"
        }
        
        Handle-ReseedFailureError -ErrorMessage $reason -Arguments $arguments -DatabaseName $Catalog.DatabaseName -AdditionalContextString $catalogStatusString
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
   Calls the Update-MailboxDatabaseCopy command for a catalog to initiate a reseed

   .PARAMETER DatabaseCopy
   Database Object for that catalog

   .PARAMETER Catalog
   Catalog status object
      
   .OUTPUTS
   True. if Failover was successful.   
#>
function Try-FailoverCorruptCatalog
{
    Param
    (
        [ValidateNotNull()]
        $DatabaseCopy,
        [Object]
        [ValidateNotNull()]
        $Catalog
    )
    
    if ($catalogCopy.Status -ieq 'Mounted')
    {
        $returnValue = Post-FailureItem -mdbName $Catalog.DatabaseName -mdbGuid $Catalog.DatabaseGuid
        # HA guarantees that a corrupt copy will be failed over within 30 seconds
        # Wait 1 minute for HA to failover the corrupt copy and then attempt to reseed
        #
        Start-Sleep -Seconds 60
        $catalogCopy = Get-MailboxDatabaseCopyStatus $catalogCopy.Name
        if ($catalogCopy.Status -ieq 'Mounted')
        {
            $isPassiveCopy = $false
            $catalogStatusString = Get-DatabaseCopyStatusString $databaseCopies
            $catalogStatusString += ". PublishFailureItemEx resultCode: $returnValue"
            [string[]]$parameters = @($Catalog.DatabaseName, $catalogStatusString)
                    # Log Failure
                    #
                    Log-Event `
                        -Arguments $Arguments `
                        -EventInfo $LogEntries.ActiveCatalogCopyCorrupt `
                        -Parameters $parameters
             return $false
        }
     }

     return $true
}


<#
   .DESCRIPTION
   Calls the Update-MailboxDatabaseCopy command for a catalog to initiate a reseed

   .PARAMETER CatalogName
   Catalog that needs to be reseeded
   
   .PARAMETER sourceServer
   Source server to reseed from
      
   .OUTPUTS
   None. Throws exception if not successful.   
#>
function Update-CatalogCopy
{
    Param
    (
        [ValidateNotNull()]
        $CatalogName,
        $SourceServer
    )
    
    Update-MailboxDatabaseCopy $CatalogName -SourceServer $SourceServer -DeleteExistingFiles -Force -confirm:$false -CatalogOnly -ErrorAction "Stop"
}

<#
   .DESCRIPTION
   Handles exceptions thrown at the time of reseed and logs appropriate error messages

   .PARAMETER Arguments
   The Arguments object constructed with script args
   
   .PARAMETER Arguments
   Actual errror message
   
   .PARAMETER DataBaseName
   The name of mailbox database copy corresponding
   to the catalog that failed to reseeded
   
   .PARAMETER AdditionalContextString
   Additional context information that would get logged in the 
   reseed failed event
   
   .OUTPUTS
   None. Throws exception if not successful.   
#>
function Handle-ReseedFailureError
{
    Param
    (
        [Object] 
        [ValidateNotNull()]
        $Arguments,
        [string]
        [ValidateNotNullOrEmpty()]
        $ErrorMessage,
        [string]
        [ValidateNotNullOrEmpty()]
        $DatabaseName,
        [string]
        [ValidateNotNull()]
        $AdditionalContextString
    )
    
    if ($ErrorMessage.ToLower().Contains("microsoft.exchange.cluster.replay.seedinprogressexception"))
    {
        # Do nothing a reseed is in progress already
        return
    }        
    
    # Log Failure
    #
    [string[]]$parameters = ($DatabaseName, $ErrorMessage, $AdditionalContextString)

    Log-Event `
            -Arguments $Arguments `
            -EventInfo $LogEntries.ReseedFailure `
            -Parameters $parameters
    Check-ReseedErrorIsSevere -ErrorMessage $ErrorMessage -DatabaseName $DatabaseName -catalogCopyStatusString $catalogStatusString -Arguments $Arguments
}


function Check-ReseedErrorIsSevere
{
    Param
    (
        [string]
        $ErrorMessage,
        [string]
        $DatabaseName,
        [string]
        [ValidateNotNull()]
        $catalogCopyStatusString,
        [Object] 
        [ValidateNotNull()]
        $Arguments
    )
    
    if ($errorMessage -eq $null)
    {
        return
    }
    
    $regexExpression = "Microsoft.Exchange.Cluster.Replay.SeederServerTransientException:.+on source server (?<sourceServer>.?[a-z,A-Z,\d]*)\. Error: Log file \'(?<LogFileName>.*)\' is corrupt\. Error: Data error \(cyclic redundancy check\)\."
    $reseedErrorRegex = new-object System.Text.RegularExpressions.RegEx($regexExpression, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $matches = $reseedErrorRegex.Match($errorMessage)
    if ($matches.Success)
    {
            $catalogCopyStatusString += " Error when trying to reseed passive copy from source server: " + $matches.Groups["sourceServer"].Value + ", Error message:" + $errorMessage
            
            [string[]]$parameters = @($DatabaseName, $catalogCopyStatusString)
            # Log Failure
            #
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $LogEntries.ActiveCatalogCopyCorrupt `
                -Parameters $parameters
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
   
   .PARAMETER AdditionalContext
   Additional context to be logged if the method fails to restart the service
   
   .PARAMETER ShouldRetry
   Bool indicating if the service restart should be retried
   
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
        $Timeout,
        [string]
        $AdditionalContext,
        [bool]
        $ShouldRetry=$true
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
        Set-SearchServiceStartMode -Server $Arguments.Server -ServiceName $exsearchServiceName -StartMode 'Disabled'
        Stop-SearchService -Server $Arguments.Server -ServiceName $exsearchServiceName -ProcessName $exsearchProcessName -Timeout $Timeout

        # Stop msftesql service
        #
        $newTimeout = $deadline - (Get-Date)
        Stop-SearchService -Server $Arguments.Server -ServiceName $msftesqlServiceName -ProcessName $msftesqlProcessName -Timeout $newTimeout

        Reset-NewFilterMonitorSettings

        Set-SearchServiceStartMode -Server $Arguments.Server -ServiceName $exsearchServiceName -StartMode 'Automatic'
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
            -EventInfo $LogEntries.RestartSuccess `
            -Parameters @("SomeString")
    }
    catch
    {
        # In case something goes wrong, always attempt to set the service start mode back.
        Set-SearchServiceStartMode -Server $Arguments.Server -ServiceName $exsearchServiceName -StartMode 'Automatic'
        $reason = $error[0].Exception.ToString() + $error[0].InvocationInfo.PositionMessage
        write-verbose ("Restart-SearchServices failed. Reason: " + $reason)
            
        if ($shouldRetry -eq $true)
        {
            # Sleep for 2 Minutes and retry restarting the services again
            #
            Start-Sleep -Seconds 120
            Restart-SearchServices `
                -Arguments $arguments `
                -Timeout $Timeout `
                -AdditionalContext $AdditionalContext `
                -ShouldRetry $false
        }
        else
        {   
            # Log Failure
            #
            Log-Event `
                -Arguments $Arguments `
                -EventInfo $LogEntries.RestartFailure `
                -Parameters ($reason, $AdditionalContext)
        }
    }
}


<#
   .DESCRIPTION
   Gets a string formatted server status object 
   
   .PARAMETER ServerStatus
   The server status object
   
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   String formatted server status object
#>
function Get-ServerStatusString
{
    Param(
        [Object]
        $ServerStatus)
    if ($ServerStatus -eq $null)
    {
        $serverStatusString = ""
    }
    else
    {
        $serverStatusString += $serverStatus | Out-String
        foreach($catalogStatus in $serverStatus.CatalogStatusArray)
        {
            $serverStatusString += $catalogStatus | Out-String
        }
    }
    
    return $serverStatusString
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
    $stopProcessTimeout = new-timespan -seconds 30
    
    write-verbose ("Stop-SearchService, deadline = " + $deadline)
    $process = Get-WMIProcess -ProcessName $ProcessName -Server $Server
    $processId = $process.ProcessId
    while ($deadline -gt (Get-Date))
    {
        # Keep stopping and terminating the process
        # until the process dies or we reach timeout
        #
        Stop-UsingServiceController -Server $Server -ServiceName $ServiceName -Timeout $stopServiceTimeout

        # Once we call stop using the service controller check if the process has really exited
        # and if not kill the process
        $processStopped = VerifyAndStop-LocalOrRemoteProcess -Server $Server -ProcessName $ProcessName -Timeout $stopProcessTimeout -ProcessId $processId
        write-Verbose ("Process stopped " + $processStopped)
        if ($processStopped)
        {
            return
        }
        else
        {
            # E14 585609 - ExSearch could be waiting on msftesql process - Terminate it first
            VerifyAndStop-LocalOrRemoteProcess -Server $Server -ProcessName $msftesqlProcessName -Timeout $stopProcessTimeout

            # When there's a deadlock, killing msftefd
            # seems to stop services faster.
            # 
            $msftefdStopped = VerifyAndStop-LocalOrRemoteProcess -Server $Server -ProcessName $msftefdProcessName -Timeout $stopProcessTimeout
        }
    }    
    
    $currentTime = Get-Date
    if ($currentTime.AddSeconds(30) -gt $deadline)
    {
        # No point in calling Wait-ForProcessToStop with a timeout value less than 30 seconds
        #
        $newTimeout = $currentTime.AddSeconds(30) - $currentTime
    }
    else
    {
        $newTimeout = $deadline - $currentTime
    }
    
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
   Change the service start mode of a service.

   .PARAMETER Server
   The name of mailbox server
   
   .PARAMETER ServiceName
   The name of the service
   
   .PARAMETER StartMode
   The start mode for the service.
   Refer to Win32_Service.ChangeStartMode for valid values.
   
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   None. Throws exception if not successful.   
#>
function Set-SearchServiceStartMode
{
    Param(
        [String] 
        [ValidateNotNullOrEmpty()]
        $Server,
        [String]
        [ValidateNotNullOrEmpty()]
        $ServiceName,
        [String]
        [ValidateSet("Automatic", "Manual", "Disabled")]
        $StartMode
    )

    $errorPref = $ErrorActionPreference
    try
    {
        $serviceFilter = "(name='$ServiceName')"
        Write-Verbose ("Service filter used for get-wmiobject = " + $serviceFilter)
        $ErrorActionPreference = "Continue"
        $service = Get-WmiObject Win32_Service -Filter $serviceFilter -ComputerName $Server
        if ($service -ne $null)
        {
            [void]$service.ChangeStartMode($StartMode)
            $service.Dispose()
        }
        else
        {
            Write-Verbose ("Could not get service object for service " + $ServiceName + " on computer " + $Server)
        }
    }
    finally
    {
        $ErrorActionPreference = $errorPref
    }
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
   None
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
    $statusCheckIntervalSeconds = 15

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
        if ($service -ne $null)
        {
            write-verbose ("Stopping service " + $ServiceName + " on " + $Server + " using service controller")
            [void]$service.StopService()
            
            # Check service status often
            # until we timeout or service
            # is stopped.
            #
            while (($service.State -ine "Stopped") -and
                   ($deadline -gt (get-date)))
            {
                $service.Dispose()
                Start-Sleep -seconds $statusCheckIntervalSeconds
                $service = get-wmiobject win32_service -filter $serviceFilter -ComputerName $Server
                write-verbose ("Service state of " + $ServiceName + " on " + $Server + ":" + $service.State)
                [void]$service.StopService()
            }
            
            if ($service.State -ieq "Stopped")
            {
                write-verbose ("Stopped service " + $ServiceName + " on " + $Server + " using service controller")
                $service.Dispose()
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
}

<#
   .DESCRIPTION
   Checks if a remote process is running and terminates the process

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
function VerifyAndStop-LocalOrRemoteProcess
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
        $Timeout,
        [int]
        $processId = 0
    )

    $deadline = (Get-Date) + $Timeout
    $statusCheckIntervalSeconds = 10

    # Save the original value for 
    # error preference
    #
    $errorPref = $ErrorActionPreference
    try
    {
        $ErrorActionPreference = "Continue"
        $process = Get-WMIProcess -ProcessName $ProcessName -Server $Server
        if ($process -ne $null)
        {
            write-verbose ("Stopping process " + $ProcessName + " on " + $Server + " using remote WMI object")
            # If the calling function has not passed in the process ID then use the ID retured by the first call to Get-WMIProcess
            # The reason we check the processID is that if a service is terminated unexpectdly the service control manager might start the 
            # service automatically based on the recovery action set for the service.
            if ($processId -eq 0 )
            {
                $processId = $process.ProcessId
            }
            
            if ($process.ProcessId -eq $processID -and $process.ProcessId -ne 0)
            {
                # Keep checking the status of the process
                # until it is the expected value or we
                # timeout.
                #
                while (($process -ne $null) -and
                        ($process.ProcessId -eq $processId) -and
                       ($deadline -gt (get-date)))
                {
                    [void]$process.Terminate(1)
                    $process.Dispose()
                    Start-Sleep -seconds $statusCheckIntervalSeconds
                    $process = Get-WMIProcess -ProcessName $ProcessName -Server $Server
                }
            }
        }
        
        if ($process -eq $null -or $process.ProcessId -eq 0 -or $process.ProcessId -ne $processId)
        {
            write-verbose ("Process " + $ProcessName + " on computer " + $Server + " has stopped.")
            return $true
        }
        else
        {
           write-verbose ("Could not stop process " + $ProcessName + " on computer " + $Server)
           $process.Dispose()
           return $false
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
        for($index = 1; $index -lt $processList.Length; $index++)
        {
            # Dispose handles to any other wmi process objects
            #
            $processList[$index].Dispose()
        }
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
    
    [REF]$outMessage = ""
    $alertCritical = Check-EventThresholdReached -Server $Arguments.Server -EventId $id -Parameters $parameters -Message ([REF]$outMessage)
    
    if ($alertCritical -eq $false)
    {
        # Instead of suppressing the event completely we just change the EventId to be a warning type
        # Warning messages logged by the TS start in the range of 5300 to 5599. All transformed events 
        # will start from 5400
        $type = "Warning"
        $id = 5400 + ($id % 100)
        $message = $message + $outMessage.Value
    }

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
    
    Write-CrimsonLogEntry `
        -Server $Arguments.Server `
        -EventId $id `
        -Category $category `
        -Type $type `
        -MessageAndParams $messageAndParams 
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
            $p.Dispose()
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
   Checks if a particular event theshold is reached.
   
   .PARAMETER EventInfo
   An object containing event id, type and message
      
   .PARAMETER Parameters
   Array of strings that comprise the parameters for the event
   These parameters must match the order and count specified
   in the message in the EventInfo paremeter
   
   .PARAMETER Message
   Reference parameter which is used to get more details about the actual and expected count
 
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   Returns true if the Event is critical. False if its not critical and Null if its not to be considered
#>
function Check-EventThresholdReached
{
    Param(
        [string]
        $Server = $env:ComputerName,
        [int]
        [ValidateNotNull()]
        $EventId,
        [String[]]
        $Parameters,
        [REF]
        $Message
    )
    
    foreach($event in $TSRetrySettings.Keys)
    {
        # Check if the use defined the events correctly in the CiTSLibrary.ps1 or ignore that entry
        #
        if ($TSRetrySettings[$event].Length -ge 2)
        {
            # Find the a matching entry
            #
            if ($TSRetrySettings[$event][0] -ne $EventId)
            {
                continue
            }
            
            [int]$retryCount = $TSRetrySettings[$event][1]
            [string]$optionalComponent = 'CurrentCount'
            if ($TSRetrySettings[$event].Length -gt 2)
            {
                # Read the optional component if specified from the replacement strings
                #
                [string]$optionalComponent = $parameters[[int]$TSRetrySettings[$event][3]]
            }
            
            try
            {
                # Check if the Parent regkey is present. If not create it
                #
                $defaultValue = Get-RegKeyValue -Server $server -Path $troubleshooterRegKey -Name ""
                if ($defaultValue -eq $null)
                {
                    Set-RegKeyValue -Server $Server -Path $troubleshooterRegKey -Name "" -Value "."
                }
                                
                # Get the retry count of the alert
                #
                $regkeyPath = $troubleshooterRegKey + $EventId.ToString()
                [int]$currentCountValue = 1
                $currentCountValueKey = Get-RegKeyValue -Server $server -Path $regKeyPath -Name $optionalComponent
                if ($currentCountValueKey -ne $null)
                {
                    # Increment the counter value by 1
                    #
                    [int]$currentCountValue = $currentCountValueKey + 1
                }
                
                # Set the current counter value in the registry 
                #
                Set-RegKeyValue -Server $Server -Path $regKeyPath -Name $optionalComponent -Value ([int]$currentCountValue)
                                
                # Verify if the counter was set correctly.
                #
                $currentCountValueKey = Get-RegKeyValue -Server $server -Path $regKeyPath -Name $optionalComponent
                if (($currentCountValueKey -eq $null) -or ([int]$currentCountValueKey -ne $currentCountValue))
                {
                    # Could not save the registry setting.
                    # We should fail safe so we return the alert as critical
                    #
                    return $true;
                }
                
                $isCritical = $currentCountValue -ge $retryCount
                if ($isCritical -eq $false)
                {
                    # Message containing additional details why the event was suppressed
                    #
                    $Message.Value = " Converting error to warning because the current error count '$currentCountValue' is below the critical retry count of '$retryCount'."
                }
                
                return $isCritical
            }
            catch
            {
                # Handle the error and do nothing. But we fail safe. So return the alert as critical (Existing behavior)
                Write-Verbose $error[0]
                $message.Value = $error[0].ToString()
                return $true
            }
        }
    }
    
    return $null
}


<#
   .DESCRIPTION
   Resets the error count of for a particular event

   .PARAMETER EventId
   An object containing event id, type and message
      
   .PARAMETER OptionalComponent
   Optional component associated with the event
 
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   Returns true if the Event is not critical.   
#>
function Reset-EventRetryCounter
{
    Param(
        [string]
        $Server = $env:ComputerName,
        [int]
        $EventId,
        [String]
        $OptionalComponent
    )
            
    try
    {   
        write-verbose ("Reset-EventRetryCounter $EventId $OptionalComponent" )
        $componentsToReset = @()
        $regkeyPath = $troubleshooterRegKey + $EventId.ToString()
        
        #If the optional component specified in null or string.Empty we update all the values in the reg key
        if ([System.String]::IsNullOrEmpty($OptionalComponent))
        {
            foreach($keyValueName in (Get-RegKeyValueNames -Server $server -Path $regkeyPath))
            {
                $componentsToReset += $keyValueName
            }
        }
        else
        {
            $componentsToReset += $OptionalComponent;
        }
        
        foreach($component in $componentsToReset)
        {
            Set-RegKeyValue -Server $Server -Path $regKeyPath -Name $component -Value ([int]0) 
        }
    }
    catch
    {
        # Do nothing. Just write to the debug trace
        Write-Verbose $error[0].ToString()
    }
    
}


<#
   .DESCRIPTION
   Gets the list of Bad IFilters that caused problems on the servers. Reads the event log to extract
   the Filter GUID

   .PARAMETER Server
   The simple NETBIOS name of mailbox server. 
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   List of Bad Ifilters that are currently enabled
#>
function Get-BadIfilterGuidsFromEventLog
{
    Param(
    [string]
    $Server = $env:ComputerName
    )
    
    write-verbose "Get-BadIfilterGuidsFromEventLog $Server"
    $iFilterExtractorString = "{(?<FilterGuid>[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12})}"
    $iFilterExtractorRegex = new-object System.Text.RegularExpressions.RegEx($iFilterExtractorString, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    
    $badIfilterGuids = @{}
    $startTime = (Get-Date).AddMinutes(-1 * $badIFilterCheckIntervalInMinutes) 
    try
    {
        $iFilterErrors = @(get-eventlog `
                    -computername $Server `
                    -after $startTime `
                    -logname "Application" `
                    -source $exsearchEventSource `
                    -ErrorAction:SilentlyContinue) | where {$_.eventId -eq $msftesqlBadIFilterEventIdEventId}
    }
    catch
    {
        $iFilterErrors = @()
    }
    
    if ($iFilterErrors.Count -gt 0)
    {
        foreach($badIfilterEvent in $iFilterErrors)
        {
            $replacementString = $badIfilterEvent.ReplacementStrings[0]
            write-verbose ("Found event " + $badIfilterEvent.Message + " with replacement string: " + $replacementString)
            $matches = $iFilterExtractorRegex.Match($replacementString)
            if ($matches.Success)
            {
                $extractedFilterGuid = "{" + $matches.Groups["FilterGuid"].Value +"}"
                write-verbose "Extracted Filter GUID: $extractedFilterGuid"
                if (!$badIfilterGuids.Contains($extractedFilterGuid))
                {
                    $badIfilterGuids.Add($extractedFilterGuid, 1);
                }
                else
                {
                    $badIfilterGuids[$extractedFilterGuid] = $badIfilterGuids[$extractedFilterGuid] + 1
                }                
            }
        }
    }
    
    return $badIfilterGuids
}

<#
   .DESCRIPTION
   Gets the list of IFilters that should be enabled by the TS, that are currently disabled

   .PARAMETER Server
   The simple NETBIOS name of mailbox server. 
   
   .PARAMETER IsRTMServer
   Bool indicating if the server is an RTM machine.
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   List of Bad Ifilters that are currently enabled
#>
function Get-IFiltersToEnable
{
    Param(
    [string]
    $Server = $env:ComputerName,
    [bool]
    $IsRTMServer = $false
    )
    
    write-verbose "Get-IFiltersToEnable $Server"
    
    $returnValue = @()
    # For E14RTM builds we disable the default Office 2007 IFilters because of a known bug and do not enable them
    #
    if ($IsRTMServer -eq $false)
    {
        $outMessage = ""
        $subKeyNames = @(Get-RegKeySubKeyNames -Server $server -Path $MsSearchFilterPath)
        foreach($subKey in $subkeyNames)
        {
            if ($subKey.ToLower().Contains($disabledIFilterSuffix.ToLower()))
            {
                Write-Verbose "Found Ifilter $subKey that should be enabled"
                $returnValue += $subKey
            }
        }
        
        if ($returnValue.Count -gt 0)
        {
            # Found Disabled IFilters. Now check if we should enable them
            # If not then return an empty array
            #
            if ((Check-EventThresholdReached -Server $server -EventId $msftesqlBadIFilterEventIdEventId -Message ([REF]$outMessage)) -eq $false)
            {
                $returnValue = @()
            }
        }
    }
    
    return $returnValue    
}

<#
   .DESCRIPTION
   Gets the list of Bad IFilters that are currently enabled on a server

   .PARAMETER Server
   The simple NETBIOS name of mailbox server. 
   
   .PARAMETER IsRTMServer
   Bool indicating if the server is an RTM machine 
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   List of Bad Ifilters that are currently enabled
#>
function Get-BadIFilters
{
    Param(
    [string]
    $Server = $env:ComputerName,
    [bool]
    $IsRTMServer = $false
    )
    
    write-verbose "Get-BadIFilters $Server"
    
    $returnValue = @()
    # For E14RTM builds we disable the default Office 2007 IFilters because of a known bug
    #
    if ($IsRTMServer -eq $false)
    {
        $badIFilterGuids = Get-BadIfilterGuidsFromEventLog -Server $server
        if ($badIFilterGuids.Count -gt 0)
        {
            $subKeyNames = @(Get-RegKeySubKeyNames -Server $server -Path $MsSearchFilterPath)
            foreach($subKey in $subkeyNames)
            {
                $currentFilterGuid = Get-RegKeyValue -Server $server -Path "$MsSearchFilterPath\$subKey" -Name $null
                foreach($badIfilterguid in $badIFilterGuids.Keys)
                {
                    if ($badIFilterGuids[$badIfilterguid] -gt $msftesqlBadIFilterEventThreshold)
                    {
                        if ($badIfilterguid -ieq $currentFilterGuid)
                        {
                            Write-Verbose "Found bad IFilter: $subKey"
                            $returnValue += $subKey
                        }
                    }
                }
            }
        }
    }
    else
    {
        $subKeyNames = @(Get-RegKeySubKeyNames -Server $server -Path $MsSearchFilterPath)
        foreach($subKey in $subkeyNames)
        {
            foreach($badIfilterName in $badIfilterNames)
            {
                if ($badIfilterName -ieq $subKey)
                {
                    Write-Verbose "Found bad IFilter: $subKey"
                    $returnValue += $subKey
                }
            }
        }
    }
    
    return $returnValue
}

<#
   .DESCRIPTION
   Gets if the current server build is an RTM server

   .PARAMETER Server
   The simple NETBIOS name of mailbox server. 
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   True if the current server build is an RTM server
#>
function Get-IsRTMServer
{

    Param(
    [string]
    $Server = $env:ComputerName
    )
    
    write-verbose "Get-IsRTMServer $server"
    $adminDisplayVersion = (Get-ExchangeServer -Identity $server).AdminDisplayVersion
    return $adminDisplayVersion.Major -eq 14 -and $adminDisplayVersion.Minor -eq 0
}

<#
   .DESCRIPTION
   Disables a IFilter on a server

   .PARAMETER Server
   The simple NETBIOS name of mailbox server. 
   
   .PARAMETER FilterName
   Name of the Filter to disabled
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   None.
#>
function Disable-BadIFilter
{
    Param(
    [string]
    $Server = $env:ComputerName,
    [string]
    [ValidateNotNullOrEmpty()]
    $FilterName
    )
    
    write-verbose "Disable-BadIFilter $Server $FilterName"
    $newFilterName = $FilterName + $disabledIFilterSuffix
    Rename-RegKey -Server $server -Path $MsSearchFilterPath -OldName $FilterName -NewName $newFilterName
}

<#
   .DESCRIPTION
   Enabled a IFilter on a server that was previously disabled by the TS

   .PARAMETER Server
   The simple NETBIOS name of mailbox server. 
   
   .PARAMETER FilterName
   Name of the Filter to disabled
      
   .INPUTS
   None. You cannot pipe objects to this function.

   .OUTPUTS
   True if the function succeeded. False otherwise.
#>
function Enable-IFilter
{
    Param(
    [string]
    $Server = $env:ComputerName,
    [string]
    [ValidateNotNullOrEmpty()]
    $FilterName
    )
    
    write-verbose "Enable-IFilter $Server $FilterName"
    try
    {
        $indexOfTsSuffix = $FilterName.ToLower().IndexOf($disabledIFilterSuffix.ToLower())
        if ($indexOfTsSuffix -gt 0)
        {
            $newFilterName = $FilterName.Substring(0, $indexOfTsSuffix);
            Rename-RegKey -Server $server -Path $MsSearchFilterPath -OldName $FilterName -NewName $newFilterName
        }
        
        return $true
    }
    catch
    {
        # Do nothing. Enabling an IFilter is not a critical operation. So if it fails we just log the error message and continue
        #
        $message=($error[0].Exception.ToString() + $error[0].InvocationInfo.PositionMessage)
        write-verbose ("Caught Exception: $message")
        Log-Event `
            -Arguments $Arguments `
            -EventInfo $LogEntries.EnablingIFilterFailed `
            -Parameters @($filterName, $message)
        return $false
    }
}

<#
   .DESCRIPTION
   Once an IFilter is Enabled/Disabled by the troubleshooter. The exchange search service will attempt to re-process every document in the retry table
   to pick up the IFilter changes. To avoid this situation we delete the NewIFilterMonitor state from the registry

   .PARAMETER Server
   The simple NETBIOS name of mailbox server. 

   .OUTPUTS
   None.
#>
function Reset-NewFilterMonitorSettings
{
    Param (
        [string]
        $Server = $env:ComputerName
    )

    RecursiveDelete-RegSubKeys -Server $Server -Path $lastKnownFiltersKeyName
    RecursiveDelete-RegSubKeys -Server $Server -Path $needReindexDatabasesKeyName    
}


<#
   .DESCRIPTION
   Recursive Deletes the subkeys of a registry key

   .PARAMETER $Server
   Remote or local server name 
   
   .PARAMETER $Path
   Registry path
   
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None
#>
function RecursiveDelete-RegSubKeys
{
    Param (
        [string]
        $Server = $env:ComputerName,
        [string]
        [ValidateNotNullOrEmpty()]
        $Path
    )

    foreach($subKey in @(Get-RegKeySubKeyNames -Server $Server -Path $Path))
    {
        if ($subKey -ne $null)
        {
            RecursiveDelete-RegSubKeys -Server $Server -Path ("$Path\$subKey")
            Delete-RegKey -Server $Server -Path $Path -SubKey $subKey
        }
    }
}

<#
   .DESCRIPTION
    Renames a key to a new value. All names under the key are copied
    but subkeys are lost

   .PARAMETER $Server
   Remote or local server name 
   
   .PARAMETER $Path
   Registry path
      
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   Registry Key value names
#>
function Rename-RegKey
{
    Param (
        [string]
        $Server = $env:ComputerName,
        [string]
        [ValidateNotNullOrEmpty()]
        $Path,
        [string]
        [ValidateNotNullOrEmpty()]
        $OldName,
        [string]
        [ValidateNotNullOrEmpty()]
        $NewName
    )
    
    write-verbose "Rename-RegKey s:$Server p:$path o:$OldName n:$NewName"

    Copy-RegKey -Server $server -Path $Path -OldName $OldName -NewName $NewName
    Delete-RegKey -Server $Server -Path $Path -SubKey $OldName
}


<#
   .DESCRIPTION
    Copies a key to a new value. All names under the key are copied
    but subkeys are lost

   .PARAMETER $Server
   Remote or local server name 
   
   .PARAMETER $Path
   Registry path
      
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   Registry Key value names
#>
function Copy-RegKey
{
    Param (
        [string]
        $Server = $env:ComputerName,
        [string]
        [ValidateNotNullOrEmpty()]
        $Path,
        [string]
        [ValidateNotNullOrEmpty()]
        $OldName,
        [string]
        [ValidateNotNullOrEmpty()]
        $NewName
    )
    
    write-verbose "Copy-RegKey s:$Server p:$path o:$OldName n:$NewName"

    $baseKey= [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $Server)

    $oldPath= "$Path\$OldName"
    $names= Get-RegKeyValueNames -Server $Server -Path $oldPath
    if ($names -eq $null)
    {
        return
    }
    
    $newPath = "$path\$NewName"
    foreach( $name in $names )
    {
        $value= Get-RegKeyValue -Server $server -Path $oldPath -Name $name
        Set-RegKeyValue -Server $server -Path $newPath -Name $name -Value $value
    }
}


<#
   .DESCRIPTION
   Gets a registry key value names

   .PARAMETER $Server
   Remote or local server name 
   
   .PARAMETER $Path
   Registry path
      
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   Registry Key value names
#>
function Get-RegKeyValueNames
{
    Param (
        [string]
        $Server = $env:ComputerName,
        [string]
        [ValidateNotNullOrEmpty()]
        $Path
    )
    write-verbose "GetRegKeyValueNames s:$Server p:$path \n ->$v"

    $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $Server)

    $key = $baseKey.OpenSubKey($path,$true)
    $values = @()
    if ($key -eq $null)
    {
        write-verbose "GetRegKey key null $Server, $path"
    }
    else
    {
        $values= @($key.GetValueNames())
        $key.Close()
    }
    
    $baseKey.Close()
    return $values
}

<#
   .DESCRIPTION
   Gets a registry key value names

   .PARAMETER $Server
   Remote or local server name 
   
   .PARAMETER $Path
   Registry path
      
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   Registry Key sub key names
#>
function Get-RegKeySubKeyNames
{
    Param (
        [string]
        $Server = $env:ComputerName,
        [string]
        [ValidateNotNullOrEmpty()]
        $Path
    )
    write-verbose "Get-RegKeySubKeyNames s:$Server p:$path \n ->$v"

    $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $Server)

    $key = $baseKey.OpenSubKey($path,$true)
    $values = @()
    if ($key -eq $null)
    {
        write-verbose "GetRegKey key null $Server, $path"
    }
    else
    {
        $values= @($key.GetSubKeyNames())
        $key.Close()
    }
    
    $baseKey.Close()
    return $values
}

<#
   .DESCRIPTION
   Gets a registry key value

   .PARAMETER $Server
   Remote or local server name 
   
   .PARAMETER $Path
   Registry path
   
   .PARAMETER $Name
   Name of the registry parameter that needs to be set
   
   .PARAMETER $DefaultValue
   Default value to be return if Registry parameter not found
   
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   Registry value
#>
function Get-RegKeyValue
{
    Param (
        [string]
        $Server = $env:ComputerName,
        [string]
        [ValidateNotNullOrEmpty()]
        $Path,
        [string]
        $Name,
        [object]
        $DefaultValue
    )
    
    Write-Verbose "Get-RegKey s:$Server p:$path n:$Name\n ->$v"

    try
    {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $Server)

        $key = $baseKey.OpenSubKey($path,$true)

        if ($key -eq $null)
        {
            write-verbose "Get-RegKey key null $Server, $path, $Name"           
        }
        else
        {
            $DefaultValue = $key.GetValue($Name, $DefaultValue)
            $key.Close()
        }
        
        $baseKey.Close()
    }
    catch
    {
        # Handle the error and do nothing
        #
        Write-Verbose $error[0]
    }
    
    return $DefaultValue
}


<#
   .DESCRIPTION
   Sets a registry key value

   .PARAMETER $Server
   Remote or local server name 
   
   .PARAMETER $Path
   Registry path
   
   .PARAMETER $Name
   Name of the registry parameter that needs to be set
   
   .PARAMETER $Value
   Value to be set
   
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None
#>
function Set-RegKeyValue
{
    Param (
        [string]
        $Server = $env:ComputerName,
        [string]
        [ValidateNotNullOrEmpty()]
        $Path,
        [string]
        $Name,
        $Value
    )
    
    write-verbose "Set-RegKeyValue $Server, $path, $Name, $Value"

    $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $Server)

    $key = $baseKey.OpenSubKey($path,$true)

    if ($key -eq $null)
    {
        $key = $baseKey.CreateSubKey($path)                
    }
    
    $key.SetValue($Name, $Value) 
    $key.Close()
    $baseKey.Close()
}


<#
   .DESCRIPTION
   Deletes a registry key  

   .PARAMETER $Server
   Remote or local server name 
   
   .PARAMETER $Path
   Registry path
   
   .PARAMETER $Subkey
   Subkey name in the path that is to be deleted
   
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   None
#>
function Delete-RegKey
{
    Param (
        [string]
        $Server = $env:ComputerName,
        [string]
        [ValidateNotNullOrEmpty()]
        $Path,
        [string]
        [ValidateNotNullOrEmpty()]
        $SubKey
    )
    
    write-verbose  "Delete-RegKey s:$Server p:$Path sk:$SubKey"

    $baseKey= [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $Server)

    $key = $baseKey.OpenSubKey($path,$true)

    if ($key -eq $null)
    {
        write-verbose "Delete-RegKey key null $Server, $path, $SubKey"              
    }
    else
    {
        $key.DeleteSubKey($SubKey)
        $key.Close()
    }
    
    $baseKey.Close()
}
    

<#
   .DESCRIPTION
   Posts a failure item in the HA log to trigger a failover 
   
   .INPUTS
   None. You cannot pipe objects to this function

   .OUTPUTS
   Return code of the PublishFailureItemEx operation
#>
function Post-FailureItem
{
    param(
        [String]
        [ValidateNotNullOrEmpty()]
        $mdbName,
        [Guid]
        [ValidateNotNullOrEmpty()]
        $mdbGuid
    )
    
    try
    {
        
        return [HaDbFailureItemHelper]::PublishFailureItem($mdbGuid, $mdbName)
    }
    catch
    {
        return $error[0].Exception.ToString() + $error[0].InvocationInfo.PositionMessage
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


#
# Get the default EN strings only. 
# This should also avoid the problem of missmatched versions between the MP and the install exchange build.
# In those cases the TS would pick up the old localized strings from the exchange directory versus the latest strings from
# the deployed MP version.
#
Import-LocalizedData -BindingVariable LocStrings -FileName CITSLibrary.strings.psd1


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
    IFiltersToEnable=(5006,"Information", $LocStrings.IFiltersToEnable)
    ServiceRestartNotNeeded=(5007,"Information", $LocStrings.ServiceRestartNotNeeded)
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
    ActiveCatalogCopyCorrupt=(5606,"Error",$LocStrings.ActiveCatalogCopyCorrupt)
    MsftefdMemoryUsageHigh=(5607,"Error",$LocStrings.MsftefdMemoryUsageHigh)
    FoundBadIFiltersEnabled=(5608,"Error",$LocStrings.FoundBadIFiltersEnabled)
    EnablingIFilterFailed=(5609,"Error",$LocStrings.EnablingIFilterFailed)
    MsftefdMemoryUsageHighWithCrashDump=(5610,"Error",$LocStrings.MsftefdMemoryUsageHighWithCrashDump)
    DetectedIndexingBacklogOrLargeRetryQueuesOnMultipleDatabases=(5611,"Error", $LocStrings.DetectedIndexingBacklogOrLargeRetryQueuesOnMultipleDatabases)
    DetectedIndexingStallExtendedPeriod=(5612,"Error", $LocStrings.DetectedIndexingStallExtendedPeriod)
    CatalogSizeGreaterThanExpectedDBLimit=(5613,"Error", $LocStrings.CatalogSizeGreaterThanExpectedDBLimit)
    CatalogReseedLoop=(5614,"Error", $LocStrings.CatalogReseedLoop)
    TroubleshooterDisabled=(5615,"Error", $LocStrings.TroubleshooterDisabled)
    TSResolutionFailed=(5616, "Error", $LocStrings.TSResolutionFailed)
    ServiceRestartAttempt=(5617,"Error", $LocStrings.ServiceRestartAttempt)
    ItemsStuckInRetryQueue=(5618,"Error", $LocStrings.ItemsStuckInRetryQueue)
    CatalogRecoveryDisabled=(5619,"Error", $LocStrings.CatalogRecoveryDisabled)
    RetryQueuesStagnant=(5620,"Error", $LocStrings.RetryQueuesStagnant)
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
}

# Dictionary containing errors which are retriable
# The first element is the eventId of the message
# The second element is the max retry count before the TS will log the error in the event log
# The third element is optional and denotes the position of the 'Component' in the eventLog parameter strings
#
$TSRetrySettings = @{
    
    # TS failed are usually transient errors because of AD issues. Log the error message only of the TS fails
    # twice consequetively
    TSFailed=($LogEntries.TSFailed[0], 2)
    
    # Once a bad Ifilter is found the TS should keep it disabled for 6 runs (6 Hours because the TS runs every hour) 
    # of the troubleshooter
    #
    BadIFiltersReset=($msftesqlBadIFilterEventIdEventId, 6)
    
    # Reseed failed events are handled per database. If the troubleshooter logs two reseed 
    # failed events for the same MDB the Event will be logged as an error by the troubleshooter
    # The third element in the array specifies that the databasename should be picked up from
    # index position 0
    ReseedFailed=($LogEntries.ReseedFailure[0], 2, 0)
    
    # Reseed failed because the active catalog copy is corrupt. On the first occurrence of this problem
    # HA should cause a failover. Only alert if the catalog fails to reseed twice for the same database.
    # The third element in the array specifies that the databasename should be picked up from
    # index position 0
    ActiveCatalogCopyCorrupt=($LogEntries.ActiveCatalogCopyCorrupt[0], 2, 0)
    
    # If the number of mailboxes left to crawl remains the same for $stallDuringCrawlThreshold consequetive runs ($stallDuringCrawlThreshold hours) asssume that the service is stalled 
    # and restart it
    MailboxCrawlStalled=($stallDuringCrawlThreshold, $stallDuringCrawlThreshold, 0)

    # If the TS reseeds a catalog everytime in consequetive runs then raise an alert
    #
    CatalogReseedLoop=($LogEntries.CatalogReseedLoop[0], 3, 0)

    # If the TS restarts the service everytime in consequetive runs then raise an alert
    #
    ServiceRestartAttempt=($LogEntries.ServiceRestartAttempt[0], 3)
}

#List if IFilters that need to be disabled on a server
#
$badIfilterNames=@(
    ".docx"
    ".pptx"
    ".xlsx"
    ".xml"
)

try
{
    $affinityValue = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'MsftesqlProcessorAffinityCount' -DefaultValue $affinityValue)
    $msftefdAffinityValue = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'MsftefdProcessorAffinityCount' -DefaultValue $msftefdAffinityValue)
    $disableBadIFilters = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'DisableBadIFilters' -DefaultValue $disableBadIFilters)
    $retryItemsThreshold = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'RetryItemsThreshold' -DefaultValue $retryItemsThreshold)
    $staleThresholdInSeconds = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'StaleThresholdInSeconds' -DefaultValue $staleThresholdInSeconds)
    $stallThresholdInSeconds = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'StallThresholdInSeconds' -DefaultValue $stallThresholdInSeconds)
    $maxPercentageCatalogSize = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'MaxPercentageCatalogSize' -DefaultValue $maxPercentageCatalogSize)
    $stallDuringCrawlThreshold = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'StallDuringCrawlThreshold' -DefaultValue $stallDuringCrawlThreshold)
    $backlogThresholdInSeconds = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'BacklogThresholdInSeconds' -DefaultValue $backlogThresholdInSeconds)
    $minRetryTableIssueThreshold = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'MinRetryTableIssueThreshold' -DefaultValue $minRetryTableIssueThreshold)
    $MaxMsftefdMemoryConsumption = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'MaxMsftefdMemoryConsumption' -DefaultValue $MaxMsftefdMemoryConsumption)
    $ExtendedStallThresholdInSeconds = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'ExtendedStallThresholdInSeconds' -DefaultValue $extendedStallThresholdInSeconds)
    $msftesqlBadIFilterEventThreshold = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'MsftesqlBadIFilterEventThreshold' -DefaultValue $MsftesqlBadIFilterEventThreshold)
    $badIFilterCheckIntervalInMinutes = [int](Get-RegKeyValue -Path $troubleshooterRegKey -Name 'BadIFilterCheckIntervalInMinutes' -DefaultValue $badIFilterCheckIntervalInMinutes)
    $disableRecoveryForDatabases = (Get-RegKeyValue -Path $troubleshooterRegKey -Name 'DisableRecoveryForDatabases' -DefaultValue "").ToString()
    $disableRecoveryForDatabasesList = $disableRecoveryForDatabases.Split((@(',',';')), [StringSplitOptions]::RemoveEmptyEntries)
    $disabledIFilterSuffix = Get-RegKeyValue -Path $troubleshooterRegKey -Name 'DisabledIFilterSuffix' -DefaultValue $disabledIFilterSuffix
}
catch
{
    # Do nothing
}

# SIG # Begin signature block
# MIIaZAYJKoZIhvcNAQcCoIIaVTCCGlECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUgYbUq61pAt8rMhzx9d0L4tEI
# I0GgghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
# 9w0BAQUFADB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSMw
# IQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQTAeFw0xMjA5MDQyMTQy
# MDlaFw0xMzAzMDQyMTQyMDlaMIGDMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMR4wHAYDVQQDExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC6pElsEPsi
# nGWiFpg7y2Fi+nQprY0GGdJxWBmKXlcNaWJuNqBO/SJ54B3HGmGO+vyjESUWyMBY
# LDGKiK4yHojbfz50V/eFpDZTykHvabhpnm1W627ksiZNc9FkcbQf1mGEiAAh72hY
# g1tJj7Tf0zXWy9kwn1P8emuahCu3IWd01PZ4tmGHmJR8Ks9n6Rm+2bpj7TxOPn0C
# 6/N/r88Pt4F+9Pvo95FIu489jMgHkxzzvXXk/GMgKZ8580FUOB5UZEC0hKo3rvMA
# jOIN+qGyDyK1p6mu1he5MPACIyAQ+mtZD+Ctn55ggZMDTA2bYhmzu5a8kVqmeIZ2
# m2zNTOwStThHAgMBAAGjggENMIIBCTATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQU3lHcG/IeSgU/EhzBvMOzZSyRBZgwHwYDVR0jBBgwFoAUyxHoytK0FlgB
# yTcuMxYWuUyaCh8wVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljQ29kU2lnUENBXzA4LTMxLTIwMTAu
# Y3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNDb2RTaWdQQ0FfMDgtMzEtMjAxMC5jcnQw
# DQYJKoZIhvcNAQEFBQADggEBACqk9+7AwyZ6g2IaeJxbxf3sFcSneBPRF1MoCwwA
# Qj84D4ncZBmENX9Iuc/reomhzU+p4LvtRxD+F9qHiRDRTBWg8BH/2pbPZM+B/TOn
# w3iT5HzVbYdx1hxh4sxOZLdzP/l7JzT2Uj9HQ8AOgXBTwZYBoku7vyoDd3tu+9BG
# ihcoMaUF4xaKuPFKaRVdM/nff5Q8R0UdrsqLx/eIHur+kQyfTwcJ7SaSbrOUGQH4
# X4HnrtqJj39aXoRftb58RuVHr/5YK5F/h9xGH1GVzMNiobXHX+vJaVxxkamNViAs
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggTDMIIDq6ADAgEC
# AhMzAAAAKzkySMGyyUjzAAAAAAArMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQTAeFw0xMjA5MDQyMTEyMzRaFw0xMzEyMDQyMTEyMzRaMIGz
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRN
# T1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046QzBGNC0zMDg2LURFRjgxJTAj
# BgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCmtjAOA2WuUFqGa4WfSKEeycDuXkkHheBwlny+
# uV9iXwYm04s5uxgipS6SrdhLiDoar5uDrsheOYzCMnsWeO03ODrxYvtoggJo7Ou7
# QIqx/qEsNmJgcDlgYg77xhg4b7CS1kANgKYNeIs2a4aKJhcY/7DrTbq7KRPmXEiO
# cEY2Jv40Nas04ffa2FzqmX0xt00fV+t81pUNZgweDjIXPizVgKHO6/eYkQLcwV/9
# OID4OX9dZMo3XDtRW12FX84eHPs0vl/lKFVwVJy47HwAVUZbKJgoVkzh8boJGZaB
# SCowyPczIGznacOz1MNOzzAeN9SYUtSpI0WyrlxBSU+0YmiTAgMBAAGjggEJMIIB
# BTAdBgNVHQ4EFgQUpRgzUz+VYKFDFu+Oxq/SK7qeWNAwHwYDVR0jBBgwFoAUIzT4
# 2VJGcArtQPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1w
# UENBLmNybDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# dDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAfsywe+Uv
# vudWtc9z26pS0RY5xrTN+tf+HmW150jzm0aIBWZqJoZe/odY3MZjjjiA9AhGfCtz
# sQ6/QarLx6qUpDfwZDnhxdX5zgfOq+Ql8Gmu1Ebi/mYyPNeXxTIh+u4aJaBeDEIs
# ETM6goP97R2zvs6RpJElcbmrcrCer+TPAGKJcKm4SlCM7i8iZKWo5k1rlSwceeyn
# ozHakGCQpG7+kwINPywkDcZqJoFRg0oQu3VjRKppCMYD6+LPC+1WOuzvcqcKDPQA
# 0yK4ryJys+fEnAsooIDK4+HXOWYw50YXGOf6gvpZC3q8qA3+HP8Di2OyTRICI08t
# s4WEO+KhR+jPFTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEF
# BQAwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jv
# c29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9y
# aXR5MB4XDTEwMDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENv
# ZGUgU2lnbmluZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCy
# cllcGTBkvx2aYCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPV
# cgDbNVcKicquIEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlc
# RdyvrT3gKGiXGqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZ
# C/6SdCnidi9U3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgG
# hVxOVoIoKgUyt0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdc
# pReejcsRj1Y8wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBTLEejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYB
# BAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy8
# 2C0wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBW
# J5flJRP8KuEKU5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNy
# b3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3Js
# MFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcN
# AQEFBQADggIBAFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGj
# I8x8UJiAIV2sPS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbN
# LeNK0rxw56gNogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y
# 4k74jKHK6BOlkU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnp
# o1hW3ZsCRUQvX/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6
# H0q70eFW6NB4lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20O
# E049fClInHLR82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8
# Z4L5UrKNMxZlHg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9G
# uwdgR2VgQE6wQuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXri
# lUEnacOTj5XJjdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEv
# mtzjcT3XAH5iR9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCC
# A++gAwIBAgIKYRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPy
# LGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRN
# aWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1
# MzA5WhcNMjEwNDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZ
# USNQrc7dGE4kD+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cO
# BJjwicwfyzMkh53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn
# 1yjcRlOwhtDlKEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3
# U21StEWQn0gASkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG
# 7bfeI0a7xC1Un68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMB
# AAGjggGrMIIBpzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A
# +3b7syuwwzWzDzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1Ud
# IwSBkDCBjYAUDqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/Is
# ZAEZFgNjb20xGTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1p
# Y3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0
# BxMuZTBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUH
# AQEESDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# L2NlcnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcD
# CDANBgkqhkiG9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwT
# q86+e4+4LtQSooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+
# jwoFyI1I4vBTFd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwY
# Tp2OawpylbihOZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPf
# wgphjvDXuBfrTot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5
# ZlizLS/n+YWGzFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8cs
# u89Ds+X57H2146SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUw
# ZuhCEl4ayJ4iIdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHu
# diG/m4LBJ1S2sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9
# La9Zj7jkIeW1sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g
# 74TKIdbrHk/Jmu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUW
# a8kTo/0xggSfMIIEmwIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIG4MBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBTqruu5Kvi+ZrIPr47zIAeMWGD0OjBYBgorBgEEAYI3AgEMMUow
# SKAggB4AQwBJAFQAUwBMAGkAYgByAGEAcgB5AC4AcABzADGhJIAiaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQCG2J2r
# fxYahpN6RcwNqUXFnExB0jspof9wrvHppUUb9pFPa17bNv61z5vROrcTmMMsozRG
# IBplTSdzp2zgPDvzQ/AxHd8yluLWWE1N/DU/RQdQ3VUxopuxG44Kk1rTvcgo4m4u
# lWu3PT9cu7pcs8dDqyI1dtVEQhJD6tncfbD4M/Hq1s2wJrwiR2HY+3U7XdnXSZqr
# EDpEds/qt1NlUpZPgTlB0j34iQaV33q0dOT+U278a0J3X0uIO2Fh5WVw96j6pXcX
# 1akhG+BG8trcDofzlGvYOC7RbbWPq2W793lHYK+kBt52dwMmHF+0XRrXER7QybBs
# LyaEqafz/XmLyI8ToYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAACs5MkjBsslI8wAAAAAAKzAJBgUr
# DgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUx
# DxcNMTMwMjA2MDAxMDAxWjAjBgkqhkiG9w0BCQQxFgQUhmUQjTQwoNQewmVrqXAX
# 9HqQ87kwDQYJKoZIhvcNAQEFBQAEggEAMtsegMwtWr2Lcn+JsTKeFpQ8VlrE58f+
# sqPf7uKO/0ZVzvL9kU4XNJjE/UAQ3xntmWw31ql10tbLN6yyEamjMb18+2kzP90P
# 7xcMLL0b0w44Zzv+hnHCIe1Lm5QMmMFzJE0mJrWUEU2iL+KO0UviVUTsTfhzLX65
# hoSlopnbv5WOfejDbfNj0n2ar8bpO7qhwJqup66EP9HEEgqjmmusu2gaYQgWCznQ
# SZj4waspypdp8SyVCWKZ/5bKXWEI0pQZ6YUCqT8nBJKJkm1dN8sm1hbm2IQQUaIY
# D5qI1mmDs7A48OspUkjyCzxORLHbHzvEmcl77Ds5uesybcAY72caqg==
# SIG # End signature block
