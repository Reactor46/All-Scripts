<#
.EXTERNALHELP CheckDatabaseRedundancy-help.xml
#>

# Copyright (c) Microsoft Corporation. All rights reserved.
#
#   Checks the redundancy of databases by validating that they have at least N 
#	configured and "healthy" copies. Active and passive copies are both counted.
#
# To use this script you need to provide either $MailboxDatabaseName or $MailboxServerName.
# To generate events for Monitoring, you need to provide -MonitoringContext switch.

[CmdletBinding(DefaultParametersetName="Server")]
param(
	[Parameter(ParameterSetName="Database",Mandatory=$true,Position=0)] 
	[string] $MailboxDatabaseName,
	
	# By default, check against the local server
	[Parameter(ParameterSetName="Server",Position=0)] 
	[string] $MailboxServerName = $env:COMPUTERNAME,
	
	# Skip checking the "default" mailbox databases. eg: Mailbox Database 0017891750
	# Specify $null (or an empty string) if you don't want to skip any databases.
	[Parameter(ParameterSetName="Server")]
	[string] $SkipDatabasesRegex = "^Mailbox Database \d{10}$",
	
	[Parameter(ParameterSetName="Monitoring",Mandatory=$true)] 
	[Parameter(ParameterSetName="Database")] 
	[Parameter(ParameterSetName="Server")]
	[switch] $MonitoringContext = $false,
	
	[Parameter(ParameterSetName="Monitoring")] 
	[Parameter(ParameterSetName="Database")] 
	[Parameter(ParameterSetName="Server")]
	[UInt32] $SleepDurationBetweenIterationsSecs = 60,
	
	[Parameter(ParameterSetName="Monitoring")] 
	[Parameter(ParameterSetName="Database")] 
	[Parameter(ParameterSetName="Server")]
	[Int32] $TerminateAfterDurationSecs = 3480, # 58 minutes; -1,0 are "Infinite"
	
	[Parameter(ParameterSetName="Monitoring")] 
	[Parameter(ParameterSetName="Database")] 
	[Parameter(ParameterSetName="Server")]
	[UInt32] $SuppressGreenEventForSecs = 600, # 10 minutes
	
	# If the total duration of being "red" exceeds this amount, raise the Red event
	[Parameter(ParameterSetName="Monitoring")] 
	[Parameter(ParameterSetName="Database")] 
	[Parameter(ParameterSetName="Server")]
	[UInt32] $ReportRedEventAfterDurationSecs = 1200, # 20 minutes
	
	# Once we raise a red event, report it periodically every $ReportRedEventIntervalSecs seconds.
	[Parameter(ParameterSetName="Monitoring")] 
	[Parameter(ParameterSetName="Database")] 
	[Parameter(ParameterSetName="Server")]
	[UInt32] $ReportRedEventIntervalSecs = 900, # 15 minutes
	
	[Parameter(ParameterSetName="Monitoring")] 
	[Parameter(ParameterSetName="Database")] 
	[Parameter(ParameterSetName="Server")]
	[switch] $SkipEventLogging = $false,
	
	[UInt32] $AtLeastNCopies = 2,
	
	# If false, detailed summary status is left out of the events/objects reported
	[switch] $ShowDetailedErrors = $false,
	
	# The email FROM address to use for the summary report
	[string] $SummaryMailFrom = $null,
	
	# Send a summary report email to the following addresses
	[string[]] $SendSummaryMailTos = $null,
	
	# Useful to "dot-source" this script as a library - call the script as such:
	#	PS D:\Exchange Mailbox\v14\Scripts> . .\CheckDatabaseRedundancy.ps1 -DotSourceMode
	[Parameter(ParameterSetName="DotSourceMode",Mandatory=$true)] 
	[switch] $DotSourceMode = $false
)

Set-StrictMode -Version 2.0
Import-LocalizedData -BindingVariable CheckDatabaseRedundancy_LocalizedStrings -FileName CheckDatabaseRedundancy.strings.psd1

function LoadExchangeSnapin
{
    if (! (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:SilentlyContinue) )
    {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
    }
}

LoadExchangeSnapin

#---------------------------------------
# Aliases for commonly used enum types #
#---------------------------------------
$CopyStatusType = [Microsoft.Exchange.Management.SystemConfigurationTasks.CopyStatus]
$ReplicationTypeType = [Microsoft.Exchange.Data.Directory.SystemConfiguration.ReplicationType]
$MountDialType = [Microsoft.Exchange.Data.Directory.SystemConfiguration.AutoDatabaseMountDial]


#------------
# Constants #
#------------

# This is the maximum copy queue length considered "healthy" for a passive copy. 
# Currently, this value is higher than BestAvailability to prevent transient alerts on spikes.
$CopyQueueLengthThreshold = 40
$InspectorQueueLengthWarningThreshold = $CopyQueueLengthThreshold
$InspectorQueueLengthFailedThreshold = 1000

#-------------------
# Script variables #
#-------------------
[System.Diagnostics.Stopwatch] $script:copyStatusStopwatch = New-Object -TypeName System.Diagnostics.Stopwatch
[System.Diagnostics.Stopwatch] $script:copyStatusAllStopwatch = New-Object -TypeName System.Diagnostics.Stopwatch
[System.Diagnostics.Stopwatch] $script:clusterNodeStopwatch = New-Object -TypeName System.Diagnostics.Stopwatch
[System.Diagnostics.Stopwatch] $script:clusterNodeOverallStopwatch = New-Object -TypeName System.Diagnostics.Stopwatch
[System.Diagnostics.Stopwatch] $script:oneIterationStopwatch = New-Object -TypeName System.Diagnostics.Stopwatch

$script:databaseToStatusTable = @{} # Hashtable indexed by DatabaseName, value of Collection<DatabaseCopyStatusEntry>
$script:databasesToCheckTable = @{} # Hashtable for the databases that we want to check, indexed by DatabaseName
$script:databaseStateTable = @{} # Hashtable indexed by DatabaseName, that holds the redundancy state of each DB
$script:clusterNodeStateTable = @{} #Hashtable indexed by server name, that holds a boolean for whether the node is online
$script:outputObjects = @() # List of objects to send to the output pipeline

[Microsoft.Exchange.Data.Directory.Management.MailboxServer] $script:mailboxServer = $null
[UInt64] $script:iteration = 0
$script:clusterOutput = $null
[string]$script:dagName = $null

[System.Text.StringBuilder]$script:report = $null
$script:IsDataCenterLibraryPresent = $false
$script:ExitOut = $false


function Is-DatabaseReplicated ([Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $mdb)
{
	if ($mdb.ReplicationType -eq $ReplicationTypeType::Remote)
	{
		return $true;
	}
	return $false;
}

function Is-DagServerOnline ([string] $serverName)
{
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0000 -f $serverName,"Is-DagServerOnline")
	[bool]$isOnline = $false
	
	if ($script:clusterNodeStateTable.Contains($serverName))
	{
		$isOnline = $script:clusterNodeStateTable[$serverName]
	}
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0001 -f $isOnline,"Is-DagServerOnline")
	return $isOnline
}

# The following states are possibly healthy for a passive copy: 
# Healthy, DisconnectedAndHealthy, SeedingSource
function Is-PassiveCopyPossiblyHealthy ([Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseCopyStatusEntry] $status)
{
	$healthy = $false;
	
	switch ($status.Status)
	{
		$CopyStatusType::Healthy { $healthy = $true }
		$CopyStatusType::DisconnectedAndHealthy { $healthy = $true }
		$CopyStatusType::SeedingSource { $healthy = $true }	
		default { }
	}
	
	return $healthy
}

function Is-ActiveReplayServiceDown ([string] $databaseName, [string] $serverName)
{
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0002 -f $databaseName,$serverName,"Is-ActiveReplayServiceDown")

	[Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseCopyStatusEntry[]] $statuses = @()
	$statuses = $script:databaseToStatusTable[$databaseName]
	
	$activeStatus = $statuses | where { $_.Name -eq "$databaseName\$serverName" }
	
	if (!$activeStatus)
	{
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0003 -f $databaseName,$serverName,$databaseName,$serverName,"Is-ActiveReplayServiceDown")
		return $true
	}

	if (!$activeStatus.ActiveCopy)
	{
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0004 -f $databaseName,$serverName,$databaseName,$serverName,"Is-ActiveReplayServiceDown")
		return $true
	}
	
	if ($activeStatus.Status -eq $CopyStatusType::ServiceDown)
	{
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0005 -f $activeStatus.Name,$databaseName,$serverName,"Is-ActiveReplayServiceDown")
		return $true
	}
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0006 -f $databaseName,$serverName,"Is-ActiveReplayServiceDown")
	return $false
}

function Populate-DatabasesTable ([Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $databases)
{
	$script:databasesToCheckTable.Clear()
	
	Foreach ($database in $databases)
	{
		$script:databasesToCheckTable[$database.Name] = $database
	}
}


# Queries the DAG object for the cluster node states.  Populates a hashtable indexed by server name,
# with booleans for whether or not the node is online.
function Populate-ClusterNodeStatus
{
	$script:dagName = $null
	$script:clusterNodeStateTable.Clear()
	
	# Find the DAG name first.
	if ($script:mailboxServer)
	{
		$script:dagName = $script:mailboxServer.DatabaseAvailabilityGroup.Name
	}
	else
	{
		# running in DB mode, which means there should only be one DB in this table
		$db = $script:databasesToCheckTable.Values | select -First 1
		$script:dagName = $db.MasterServerOrAvailabilityGroup.Name
	}
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0013 -f $script:dagName,"Populate-ClusterNodeStatus")
	
	$script:clusterNodeOverallStopwatch.Reset();
	$script:clusterNodeOverallStopwatch.Start();
		
	$dag = Get-DatabaseAvailabilityGroup $script:dagName -status
	if ($dag)
	{
		if (!$dag.Servers -or `
			($dag.Servers.Count -eq 0))
		{
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0015 -f $script:dagName,"Populate-ClusterNodeStatus")	
		}
		else
		{
            foreach ($server in $dag.Servers) {
                # Server is online if it is in the operational servers list
                $IsOnline = [bool] ($dag.OperationalServers | where {$_ -eq $server})
                $script:clusterNodeStateTable.Add($server.Name, $IsOnline)
            }
        }
	}
	
	$script:clusterNodeOverallStopwatch.Stop()
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0018 -f $script:clusterNodeOverallStopwatch.Elapsed.TotalMilliseconds,"Populate-ClusterNodeStatus")
}

function Check-Databases ([Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $databases, [string] $ParameterSetName)
{
	Populate-DatabasesTable $databases

	# find the servers to check copy statuses on
	[String[]]$servers = Get-ServersForDatabases $databases
	
	if ($servers.Length -lt 2)
	{
		# Normally we should not get here, since we're only checking replicated DBs, which means
		# we should have at least 2 distinct servers. However, this can happen if copies are 
		# removed while this script is running.
		Log-Warning ($CheckDatabaseRedundancy_LocalizedStrings.res_0019 -f $servers.Length,"Get-ServersForDatabases")
	}
	
	# get the status results and index them by database name
	$script:databaseToStatusTable.Clear()
	$script:databaseToStatusTable = Get-CopyStatusFromAllServers $servers $ParameterSetName | `
									Group-Object -AsHashTable -Property DatabaseName
	
	# look up the cluster node status for the DAG
	Populate-ClusterNodeStatus
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0020 -f "Check-Databases")

	# Filter out the databases we are not going to check, and then perform the redundancy check. 
	$script:databaseToStatusTable.Keys | `
		where { $script:databasesToCheckTable.Contains( $_ ) } | `
		foreach { Check-DatabaseRedundancy $_ }
	
	# NOTE:
	# If the DB is completely removed from AD while this script is running, we may keep reporting
	# a Red alert for it. If need be, we can log a green event for the DB in this case...
}

# This object represents a DB's redundancy state. It is initialized once at script startup and 
# is subsequently maintained over multiple passes of Check-Databases.
function CreateEmptyDatabaseRedundancyEntry
{
	[CheckHADatabaseRedundancy.DatabaseRedundancyEntry]$entry = New-Object -TypeName "CheckHADatabaseRedundancy.DatabaseRedundancyEntry"
	return $entry
}

function Initialize-DatabaseRedundancyEntry ([CheckHADatabaseRedundancy.DatabaseRedundancyEntry] $dbRedundancy)
{
	$dbRedundancy.LastRedundancyCount = $dbRedundancy.CurrentRedundancyCount
	$dbRedundancy.LastState = $dbRedundancy.CurrentState
	$dbRedundancy.CurrentRedundancyCount = 0
	$dbRedundancy.CurrentState = "Unknown"
	$dbRedundancy.CurrentErrorMessages = $null
}

function Get-SummaryCopyStatusString(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)] [Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseCopyStatusEntry] $status)
{
	Begin
	{
		$statusOutputs = @()
	}
	Process
	{
		$statusOutput = $status | Select-Object  *,`
			@{Name="RealCopyQueue"; Expression={ [Math]::Max($_.LastLogGenerated - $_.LastLogCopied, 0) }}, `
			@{Name="InspectorQueue"; Expression={ [Math]::Max($_.LastLogCopied - $_.LastLogInspected, 0) }}, `
			@{Name="ReplayQueue"; Expression={ $_.ReplayQueueLength }}, `
			@{Name="CIState"; Expression={ $_.ContentIndexState }} 
			
		$statusOutputs += $statusOutput
	}
	End
	{
		[string]$statusStr = ($statusOutputs | ft -Wrap Name,Status,RealCopyQueue,InspectorQueue,ReplayQueue,CIState | Out-String)
		$statusStr = $statusStr -replace "\s+$" # trim the white space at the end
		Write-Output $statusStr
	}
}

# Logic to decide if a DB has insufficient redundancy
function Check-DatabaseRedundancy ([string] $dbName)
{
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0021 -f $dbName,"Check-DatabaseRedundancy")
	
	# Initialize the DB redundancy state if necessary
	if (!$script:databaseStateTable.Contains($dbName))
	{
		$dbState = CreateEmptyDatabaseRedundancyEntry
		$dbState.DatabaseName = $dbName
		$script:databaseStateTable.Add($dbName, $dbState)
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0022 -f $dbName,"Check-DatabaseRedundancy")
	}
	
	# Retrieve the redundancy state object, and initialize states
	$dbRedundancy = $script:databaseStateTable[$dbName]
	Initialize-DatabaseRedundancyEntry $dbRedundancy
	
	# Get the list of copy status entries from the hashtable
	[Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseCopyStatusEntry[]] $statuses = @()
	$statuses = $script:databaseToStatusTable[$dbName]
	[string[]] $tmpErrMessages = @()
	[string] $errMsg = $null
	[string] $summaryStatusStr = $null
	[CheckHADatabaseRedundancy.CopyCheckState]$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Unknown
	
	# In case there's only one configured copy, let's report that as an error
	if ($statuses.Count -lt $AtLeastNCopies)
	{
		$tmpErrMessages += "The number of configured copies for database '$dbName' ($($statuses.Count)) is less than the required redundancy count ($AtLeastNCopies)."
	}
	
	Foreach ($status in $statuses)
	{		
		# Check the health of the active or passive copy
		($errMsg,$checkState) = Get-DatabaseCopyHealth $status
		
		if ($checkState -eq [CheckHADatabaseRedundancy.CopyCheckState]::Passed)
		{
			$dbRedundancy.CurrentRedundancyCount++
		}
		elseif ($checkState -eq [CheckHADatabaseRedundancy.CopyCheckState]::Warning)
		{
			$dbRedundancy.CurrentRedundancyCount++
			$tmpErrMessages += $errMsg
		}
		else
		{
			# This copy has failed the check, so let's record the reason why
			$tmpErrMessages += $errMsg
		}
	}
	
	# If we've got some errors, remember them for emailing purposes
	if ($tmpErrMessages.Length -gt 0)
	{
		# Append the summary status
		$summaryStatusStr = ($statuses | Get-SummaryCopyStatusString)
		$tmpErrMessages += "$summaryStatusStr"
		
		# Add the overall errors to the history for this DB
		$dbRedundancy.AddErrorRecordToHistory( [DateTime]::UtcNow, $tmpErrMessages )
	}
	
	if ($ShowDetailedErrors)
	{
		# Additionally, log the copy status output into the event
		$statusStr = $summaryStatusStr
		if (!$statusStr)
		{
			$statusStr = ($statuses | Get-SummaryCopyStatusString)
			if ($statusStr)
			{
				$tmpErrMessages += "`n`n================`n Summary Status `n================`n`n$statusStr"
			}
		}
		
		$statusStr = ($statuses | fl | Out-String)
		if ($statusStr)
		{
			$tmpErrMessages += "`n`n===============`n Full Status `n===============`n`n$statusStr"
		}
	}
	
	if ($tmpErrMessages.Length -gt 0)
	{
		$dbRedundancy.CurrentErrorMessages = $tmpErrMessages
	}
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0023 -f $dbRedundancy.CurrentRedundancyCount,$dbRedundancy.LastRedundancyCount,$dbName,"Check-DatabaseRedundancy")
	
	# Decide if the state is Red or Green
	if ($dbRedundancy.CurrentRedundancyCount -lt $AtLeastNCopies)
	{
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0024 -f $dbName,$AtLeastNCopies,"Check-DatabaseRedundancy")
		$dbRedundancy.CurrentState = [CheckHADatabaseRedundancy.AlertState]::Red
	}
	else
	{
		$dbRedundancy.CurrentState = [CheckHADatabaseRedundancy.AlertState]::Green
	}
	
	# record the state transition times
	[datetime]$nowUtc = [DateTime]::UtcNow
	if ($dbRedundancy.IsTransitioningState)
	{
		$dbRedundancy.LastStateTransitionUtc = $nowUtc
		
		if ($dbRedundancy.CurrentState -eq [CheckHADatabaseRedundancy.AlertState]::Green)
		{
			$dbRedundancy.LastGreenTransitionUtc = $nowUtc
			
			if ($dbRedundancy.LastRedTransitionUtc)
			{
				[TimeSpan]$prevTimeInRed = $dbRedundancy.LastGreenTransitionUtc.Subtract( $dbRedundancy.LastRedTransitionUtc )
				$dbRedundancy.PreviousTotalRedDuration = $dbRedundancy.PreviousTotalRedDuration.Add( $prevTimeInRed )
			}
		}
		elseif ($dbRedundancy.CurrentState -eq [CheckHADatabaseRedundancy.AlertState]::Red)
		{
			$dbRedundancy.LastRedTransitionUtc = $nowUtc
		}
	}
	
	# Report a red/green event if necessary (suppression may occur if MonitoringContext is specified)
	PossiblyReport-RedGreenStatus $dbRedundancy
	
}

# Reports Red/Green status via mail/event etc, taking into account whether or not
# we are running in the MonitoringContext (which affects suppression)
function PossiblyReport-RedGreenStatus ( [CheckHADatabaseRedundancy.DatabaseRedundancyEntry] $dbRedundancy )
{
	if ($MonitoringContext)
	{
		# In the monitoring context, we should run the suppression logic
		
		if ($dbRedundancy.CurrentState -eq [CheckHADatabaseRedundancy.AlertState]::Green)
		{
			[int]$timeInGreenSecs = Get-ElapsedTimeInSeconds $dbRedundancy.LastGreenTransitionUtc
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0025 -f $dbRedundancy.DatabaseName,$timeInGreenSecs,"PossiblyReport-RedGreenStatus")
			if ($SuppressGreenEventForSecs -le 0 -or `
					(($timeInGreenSecs -gt $SuppressGreenEventForSecs) -and `
					($dbRedundancy.LastGreenReportedUtc -eq $null)))
			{
				# Only log a green event once, or if it transitions into Green again
				Report-GreenStatus $dbRedundancy
			}
		}
		elseif ($dbRedundancy.CurrentState -eq [CheckHADatabaseRedundancy.AlertState]::Red)
		{
			# We need to log an event if the total duration of being in Red (including flickering states)
			# is larger than $ReportRedEventAfterDurationSecs. 
			$totalRedSecs = $dbRedundancy.TotalRedDuration.TotalSeconds
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0026 -f $dbRedundancy.DatabaseName,$totalRedSecs,"PossiblyReport-RedGreenStatus")
			if ($ReportRedEventAfterDurationSecs -le 0 -or `
				$totalRedSecs -gt $ReportRedEventAfterDurationSecs)
			{
				# Reporting a red event for the first time
				if ($dbRedundancy.LastRedReportedUtc -eq $null)
				{
					Report-RedStatus $dbRedundancy
				}
				else
				{
					# Additionally, we need to log an event every $ReportRedEventIntervalSecs seconds
					# while the DB is in "Red".
					[int]$timeSinceLastRedEventSecs = Get-ElapsedTimeInSeconds $dbRedundancy.LastRedReportedUtc
					if ($timeSinceLastRedEventSecs -gt $ReportRedEventIntervalSecs)
					{
						Report-RedStatus $dbRedundancy	
					}
				}
			}
		}
	}
	else
	{
		# No monitoring context, so no suppression
		if ($dbRedundancy.CurrentState -eq [CheckHADatabaseRedundancy.AlertState]::Green)
		{
			Report-GreenStatus $dbRedundancy
		}
		elseif ($dbRedundancy.CurrentState -eq [CheckHADatabaseRedundancy.AlertState]::Red)
		{
			Report-RedStatus $dbRedundancy	
		}
	}	
}

# Reports Green status via mail/event as appropriate [no suppression].
function Report-GreenStatus ( [CheckHADatabaseRedundancy.DatabaseRedundancyEntry] $dbRedundancy )
{
	$dbRedundancy.LastGreenReportedUtc = [DateTime]::UtcNow
	
	[string]$dbCopyForAlerting = Get-DatabaseCopyForAlerting $dbRedundancy.DatabaseName
	[bool]$writeOutput = $true
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0027 -f $dbCopyForAlerting,"Report-GreenStatus")
	
	if ($MonitoringContext)
	{
		if (!$SkipEventLogging)
		{
			# Write green event log into Application log
			$writeOutput = $false
			# MonitoringDatabaseRedundancyCheckPassed - EventId 4114
			Write-HAAppLogInformationEvent "40041012" 1 @($dbCopyForAlerting, $dbRedundancy.CurrentRedundancyCount, $dbRedundancy.GetErrorStringForAlerting())
		}
	}
	
	if ($writeOutput)
	{
		Write-Output $dbRedundancy
	}
}

# Reports Red status via mail/event as appropriate [no suppression].
function Report-RedStatus ( [CheckHADatabaseRedundancy.DatabaseRedundancyEntry] $dbRedundancy )
{
	$dbRedundancy.LastRedReportedUtc = [DateTime]::UtcNow
	$dbRedundancy.LastGreenReportedUtc = $null
	
	[string]$dbCopyForAlerting = Get-DatabaseCopyForAlerting $dbRedundancy.DatabaseName
	[bool]$writeOutput = $true
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0028 -f $dbCopyForAlerting,"Report-RedStatus")
	
	if ($MonitoringContext)
	{
		if (!$SkipEventLogging)
		{
			# Write red event log into Application log
			$writeOutput = $false
			# MonitoringDatabaseRedundancyCheckFailed - EventId 4113
			Write-HAAppLogErrorEvent "C0041011" 1 @($dbCopyForAlerting, $dbRedundancy.CurrentRedundancyCount, $dbRedundancy.GetErrorStringForAlerting())
		}
	}
	
	if ($writeOutput)
	{
		Write-Output $dbRedundancy
	}
}

function Get-DatabaseCopyForAlerting ( [string] $dbName )
{
	# We need to report the Database name (not the DBCopy name)
	return $dbName
}

function Get-ElapsedTimeInSeconds( [DateTime] $startTimeUtc )
{
	[TimeSpan]$elapsedTime = [DateTime]::UtcNow.Subtract( $startTimeUtc )
	[int]$elapsedSeconds = [int][System.Math]::Floor($elapsedTime.TotalSeconds)
	return $elapsedSeconds
}


# Returns the string describing why the copy was not healthy. $null if the copy is healthy.
function Get-DatabaseCopyHealth ([Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseCopyStatusEntry] $copyStatus)
{
	if ($copyStatus.ActiveCopy)
	{
		return Get-ActiveDatabaseCopyHealth $copyStatus
	}
	else
	{
		return Get-PassiveDatabaseCopyHealth $copyStatus
	}
}

# Returns the string describing why the copy was not healthy. $null if the copy is healthy.
function Get-ActiveDatabaseCopyHealth ( `
	[Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseCopyStatusEntry] $copyStatus)
{	
	$dbCopy = $copyStatus.Name
	$dbName = $copyStatus.DatabaseName
	$server = $copyStatus.MailboxServer
	[string]$errMsg = $null
	[CheckHADatabaseRedundancy.CopyCheckState]$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Unknown
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0029 -f $copyStatus.Status,$copyStatus.ErrorEventId,$copyStatus.ErrorMessage,$copyStatus.SuspendComment,$dbCopy,"Get-ActiveDatabaseCopyHealth")
	
	# Log that the DB is not replicated
	if ($VerbosePreference -ne "SilentlyContinue")
	{
		if ( !(Is-DatabaseReplicated ($script:databasesToCheckTable[$dbName])) )
		{
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0030 -f $dbName,"Get-ActiveDatabaseCopyHealth")
		}
	}
	
	# First, we need the cluster node status to be Up
	if (!(Is-DagServerOnline $server))
	{
		$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Failed
		$errMsg = "Active copy '$dbCopy' is not UP according to clustering."
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0031 -f $errMsg,"Get-ActiveDatabaseCopyHealth")
		return $errMsg,$checkState
	}
	
	# If replay service is down, we'll just assume that this copy isn't healthy since 
	# all the passives will be "stalled", which means we're anyway down to at most 1 copy.
	if ($copyStatus.Status -eq $CopyStatusType::ServiceDown)
	{
		$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Failed
		$errMsg = "Active copy '$dbCopy' has replay service down. Assuming the copy is unhealthy."
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0031 -f $errMsg,"Get-ActiveDatabaseCopyHealth")
		return $errMsg,$checkState
	}
	
	if ( 	($copyStatus.Status -eq $CopyStatusType::Dismounted) `
		-or ($copyStatus.Status -eq $CopyStatusType::Dismounting) )
	{
		# There may have been a permanent failure (such as a DB corruption) that is preventing
		# the DB from mounting. If so, there will be an error message recorded.
		if ($copyStatus.ErrorMessage)
		{
			$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Failed
			$errMsg = "Active copy '$dbCopy' is dismounted with an error. Error: $($copyStatus.ErrorMessage)."
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0031 -f $errMsg,"Get-ActiveDatabaseCopyHealth")
			return $errMsg,$checkState
		}
		else
		{
			# A dismounted copy is fine
			$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Passed
			return $errMsg,$checkState
		}		
	}
	
	if ($copyStatus.Status -eq $CopyStatusType::Mounted)
	{
		$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Passed
		return $errMsg,$checkState
	}
	elseif ($copyStatus.Status -eq $CopyStatusType::Mounting)
	{
		# NOTE: This is only a warning and doesn't cause RED alert to go off!
		$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Warning
		$errMsg = "Active copy '$dbCopy' is in 'Mounting' state."
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0032 -f $errMsg,"Get-PassiveDatabaseCopyHealth")
		return $errMsg,$checkState
	}
	
	# Any other state, assume the worst
	$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Failed
	$errMsg = "Active copy '$dbCopy' has some unknown/unhealthy state. Status: $($copyStatus.Status)."
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0031 -f $errMsg,"Get-ActiveDatabaseCopyHealth")
	return $errMsg,$checkState
}

# Returns the string describing why the copy was not healthy. $null if the copy is healthy.
function Get-PassiveDatabaseCopyHealth ( `
	[Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseCopyStatusEntry] $copyStatus)
{
	$dbCopy = $copyStatus.Name
	$dbName = $copyStatus.DatabaseName
	$server = $copyStatus.MailboxServer
	$activeServer = $copyStatus.ActiveDatabaseCopy
	[string]$errMsg = $null
	[CheckHADatabaseRedundancy.CopyCheckState]$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Unknown
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0033 -f $copyStatus.Status,$copyStatus.CopyQueueLength,$copyStatus.ReplayQueueLength,$copyStatus.ErrorEventId,$copyStatus.ErrorMessage,$copyStatus.SuspendComment,$dbCopy,$dbName,"Get-PassiveDatabaseCopyHealth")
	
	# First, we need the cluster node status to be Up
	if (!(Is-DagServerOnline $server))
	{
		$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Failed
		$errMsg = "Passive copy '$dbCopy' is not UP according to clustering."
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0031 -f $errMsg,"Get-PassiveDatabaseCopyHealth")
		return $errMsg,$checkState
	}
	
	# Rule out the obviously unhealthy cases first
	if (!(Is-PassiveCopyPossiblyHealthy $copyStatus))
	{
		$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Failed
		$errMsg = "Passive copy '$dbCopy' is not in a good state. Status: $($copyStatus.Status)."
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0031 -f $errMsg,"Get-PassiveDatabaseCopyHealth")
		return $errMsg,$checkState
	}
	
	# Check the *real* copy queue length first (i.e. not including the inspector queue)
	$realCopyQ = [Math]::Max($copyStatus.LastLogGenerated - $copyStatus.LastLogCopied, 0)
	if ($realCopyQ -gt $CopyQueueLengthThreshold)
	{
		$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Failed
		$errMsg = "Passive copy '$dbCopy' has actual log copy queue higher than the threshold of '$CopyQueueLengthThreshold'. Copy queue: $realCopyQ."
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0031 -f $errMsg,"Get-PassiveDatabaseCopyHealth")
		return $errMsg,$checkState
	}
	
	# Check the inspector queue length
	$inspectorQ = [Math]::Max($copyStatus.LastLogCopied - $copyStatus.LastLogInspected, 0)
	if ($inspectorQ -gt $InspectorQueueLengthFailedThreshold)
	{
		$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Failed
		$errMsg = "Passive copy '$dbCopy' has an inspector queue higher than the failure threshold of '$InspectorQueueLengthFailedThreshold'. Inspector queue: $inspectorQ."
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0031 -f $errMsg,"Get-PassiveDatabaseCopyHealth")
		return $errMsg,$checkState
	}
	
	# Even if the copy queue is small, we can't trust it because the active replay service might
	# be down, in which case the queues will be stale. (E14# 138911)
	# So, if the active status is "ServiceDown", but the node is up, we can be fairly certain
	# that we shouldn't trust the queues.	
	if ((Is-ActiveReplayServiceDown $dbName $activeServer) -and `
		(Is-DagServerOnline $activeServer))
	{
		$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Failed
		$errMsg = "Passive copy '$dbCopy' has a small copy queue length, but it could be stale. The active replay service on server '$activeServer' appears to be down. Copy queue: $($copyStatus.CopyQueueLength)."
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0031 -f $errMsg,"Get-PassiveDatabaseCopyHealth")
		return $errMsg,$checkState
	}
	
	if ($inspectorQ -gt $InspectorQueueLengthWarningThreshold)
	{
		# NOTE: This is only a warning and doesn't cause RED alert to go off!
		$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Warning
		$errMsg = "Passive copy '$dbCopy' has an inspector queue higher than the warning threshold of '$InspectorQueueLengthWarningThreshold'. Inspector queue: $inspectorQ."
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0032 -f $errMsg,"Get-PassiveDatabaseCopyHealth")
		return $errMsg,$checkState
	}
	
	# The copy is good in the time alotted...
	$checkState = [CheckHADatabaseRedundancy.CopyCheckState]::Passed
	return $errMsg,$checkState
}

# Given a list of databases, find the set of unique servers hosting copies of all of them.
# Returns an array with all the server names.
function Get-ServersForDatabases ( `
	[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $databases)
{
	$servers = @{}
	
	Foreach ($db in $databases) 
	{
		$db.Servers | % { `
			if (!$servers.Contains($_.Name))
			{
				$servers.Add($_.Name, 1);
			}
		}
	}
	
	# convert the hashtable into an array
	[String[]]$serversList = @()
	$servers.Keys | % { $serversList += $_ }
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0034 -f $serversList.Length,"Get-ServersForDatabases")
	return $serversList;
}

# Runs get-mailboxdatabasecopystatus against copy(ies) on a server and returns an array of status results.
# Return type: Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseCopyStatusEntry[]
function Get-CopyStatusFromServer ([string] $server, [string] $ParameterSetName)
{
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0035 -f $server,"Get-CopyStatusFromServer")
	
	$script:copyStatusStopwatch.Reset();
	$script:copyStatusStopwatch.Start();
	
	[Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseCopyStatusEntry[]] $statuses = @()
	
	if ($ParameterSetName -eq "Database" )
	{
		$statuses = @( Get-MailboxDatabaseCopyStatus "$MailboxDatabaseName\$server" )
	}
	else
	{
		$statuses = Get-MailboxDatabaseCopyStatus -Server $server
	}
	
	$script:copyStatusStopwatch.Stop();
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0036 -f $script:copyStatusStopwatch.Elapsed.TotalMilliseconds,$server,"Get-CopyStatusFromServer")
	
	return $statuses
}

# Synchronously executes get-mdbcs against all the specified servers and returns an
# array of type DatabaseCopyStatusEntry, which holds all the statuses returned.
function Get-CopyStatusFromAllServers ([String[]] $servers, [string] $ParameterSetName)
{
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0037 -f "Get-CopyStatusFromAllServers")
	
	$script:copyStatusAllStopwatch.Reset();
	$script:copyStatusAllStopwatch.Start();
	
	[Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseCopyStatusEntry[]] $allStatuses = @()

	Foreach ($server in $servers)
	{
		[Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseCopyStatusEntry[]] $statuses = @()
		$statuses = Get-CopyStatusFromServer $server $ParameterSetName
		$allStatuses += $statuses
	}
	
	$script:copyStatusAllStopwatch.Stop();
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0038 -f $script:copyStatusAllStopwatch.Elapsed.TotalMilliseconds,"Get-CopyStatusFromAllServers")
	
	return $allStatuses
}

# Wouldn't it be nice to be able to run get-mdbcs in parallel across all the servers? 
# Unfortunately, when I tested this method against just 3 servers in the DAG, it took
# ~26 seconds (25854.3543ms) !!! 
# $REVIEW: Is Start-Job supposed to be that slow? I suppose it could be faster if we
# reused the same PSSession. I'll look into "Invoke-Command -AsJob" in future, but for
# now, the overhead is definitely not worth it...
function Get-CopyStatusFromAllServersAsync ([String[]] $servers)
{
	# We'll run Get-CopyStatus in parallel across all the servers so that we don't 
	# excessively slow down the status retrieval in case some servers are down.
	[System.Management.Automation.PSRemotingJob[]] $asyncJobs = @()
	
	$getStatusCmd = 
	{
		Process
		{
			$tmpServer = $_;
			
			if (! (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:SilentlyContinue) )
    		{
        		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
    		}
			
			[Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseCopyStatusEntry[]] $statuses = @()
	
			if ($PSCmdlet.ParameterSetName  -eq "Database" )
			{
				$statuses = @( Get-MailboxDatabaseCopyStatus "$MailboxDatabaseName\$tmpServer" )
			}
			else
			{
				$statuses = @( Get-MailboxDatabaseCopyStatus -Server $tmpServer )
			}
			
			return $statuses
		}
	}
	
	$sw = New-Object -TypeName System.Diagnostics.Stopwatch
	$sw.Reset();
	$sw.Start();
	
	######
	## <Timed portion>
	
	Foreach ($server in $servers)
	{
		$asyncJobs += Start-Job -ScriptBlock $getStatusCmd -InputObject $server
	}
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0039 -f $asyncJobs.Length,"Get-CopyStatusFromAllServersAsync")
	
	# wait on all of them to complete
	Wait-Job $asyncJobs
	
	$sw.Stop();
	## </Timed portion>
	######
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0040 -f $sw.Elapsed.TotalMilliseconds,"Get-CopyStatusFromAllServersAsync")
	
	Foreach ($job in $asyncJobs)
	{
		$results = Receive-Job $job
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0041 -f $results)
		$results
	}
}

#######################################################################
#    Dynamic code compiler logic
#######################################################################
function ConstructReferences([Array]$References)
{
    #
    # Build up a compiler params object...
    $refs = @()
    $refs.AddRange( @("${framework}\System.dll",
        "${framework}\system.windows.forms.dll",
        "${framework}\System.data.dll",
        "${framework}\System.Drawing.dll",
        "${framework}\System.Xml.dll"))
    if (($References -ne $null) -and ($References.Count -ge 1))
    {
        foreach ($refAssembly in $References)
        {
            [string] $refTmp = $refAssembly
            if ($refTmp.IndexOf("\") -eq -1)
            {
                $refTmp = "${framework}\$refTmp"
            }
            
            $refs.Add($refTmp);
        }           
    }
    
    return $refs
}

# Compile the types to be used for tracking the Database Redundancy state. 
# The compilation is only performed once per runspace and is entirely in memory.
function Prepare-DatabaseRedundancyEntryDefinition
{
	$code = '
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace CheckHADatabaseRedundancy
{
    public enum AlertState : int
    {
        Unknown = 0,
        Green,
        Red
    }

    // Enum describing the state of an individual database copy. At the moment, 
    // both Passed and Warning are treated as having passed the redundancy check 
    // and hence CurrentRedundancyCount is incremented.
    public enum CopyCheckState : int
    {
        Unknown = 0,
        Passed,
        Warning,
        Failed
    }

    public class DatabaseRedundancyEntry
    {
        public class ErrorRecord
        {
            public DateTime ErrorTime { get; set; }
            public string[] ErrorMessages { get; set; }
			
			public string GetErrorStringForAlerting()
            {
                if (ErrorMessages == null || ErrorMessages.Length == 0)
                {
                    return String.Empty;
                }

                return String.Join("\n", ErrorMessages);
            }
        }

        public string DatabaseName { get; set; }
        public int LastRedundancyCount { get; set; }
        public int CurrentRedundancyCount { get; set; }
        public AlertState LastState { get; set; }
        public AlertState CurrentState { get; set; }
        public DateTime? LastStateTransitionUtc { get; set; }
        public DateTime? LastGreenTransitionUtc { get; set; }
        public DateTime? LastRedTransitionUtc { get; set; }
        public DateTime? LastGreenReportedUtc { get; set; }
        public DateTime? LastRedReportedUtc { get; set; }

        // the previous total red duration (not counting the current stretch of reds)
        public TimeSpan PreviousTotalRedDuration { get; set; }

        public TimeSpan TotalRedDuration
        {
            get
            {
                if (this.CurrentState == AlertState.Red)
                {
                    // count the current duration of reds
                    return this.PreviousTotalRedDuration + (DateTime.UtcNow - this.LastRedTransitionUtc.Value);
                }
                else
                {
                    return this.PreviousTotalRedDuration;
                }
            }
        }

        public bool IsTransitioningState
        {
            get { return LastState != CurrentState; }
        }

        public bool HasErrorsInHistory
        {
            get
            {
                if (this.ErrorHistory == null || this.ErrorHistory.Count == 0)
                {
                    return false;
                }
                return true;
            }
        }
        
        public string[] CurrentErrorMessages { get; set; }
        public List<ErrorRecord> ErrorHistory { get; private set; }
			
        public string GetErrorStringForAlerting()
        {
            if (CurrentErrorMessages == null || CurrentErrorMessages.Length == 0)
            {
                return String.Empty;
            }

            return String.Join("\n", CurrentErrorMessages);
        }

        // Create a copy of the errorMessages array and then add the record to the history
        public void AddErrorRecordToHistory(DateTime errorTime, string[] errorMessages)
        {
            string[] tmpMessages = new string[errorMessages.Length];
            errorMessages.CopyTo(tmpMessages, 0);

            ErrorRecord er = new ErrorRecord();
            er.ErrorTime = errorTime;
            er.ErrorMessages = tmpMessages;

            if (this.ErrorHistory == null)
            {
                this.ErrorHistory = new List<ErrorRecord>(15);
            }
            this.ErrorHistory.Add(er);
        }
    }

    public static class EventLogger
    {
        public static void WriteLocalizedEvent(
            string logName, 		// eg: Application
            string sourceName, 		// eg: MSExchangeRepl
            long eventId, 			// Message resource ID: eg: (long)0xC0041011
            int categoryId, 		// category of the event
            EventLogEntryType entryType, // error, information, warning
            byte[] data,
            params object[] messageArgs)
        {
            EventLog eventLog = new EventLog(logName, Environment.MachineName, sourceName);

            EventInstance instance = new EventInstance(
                eventId,
                categoryId,
                entryType);

            eventLog.WriteEvent(instance, data, messageArgs);
        }
    }
}'

	$checkCompiledCmd = 
	{
		# Check if the type is loaded. If not, a RuntimeException is thrown.
		[CheckHADatabaseRedundancy.AlertState];
	}
	[bool]$isCompiled = TryExecute-ScriptBlock -runCommand:$checkCompiledCmd -silentOnErrors:$true

	if (!$isCompiled)
    {
        ##################################################################
        # So now we compile the code and use .NET object access to run it.
        ##################################################################
        
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0042 )
		Add-Type -TypeDefinition $code -Language "CSharpVersion3"
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0043 )
    }
}

function Write-HAAppLogInformationEvent(
	[Parameter(Mandatory=$true)] [string] $eventId, # eg: "C0041011"
	[Parameter(Mandatory=$true)] [int] $categoryId,
	[Object[]] $messageArgs)
{
	Write-LocalizedEventLog "Application" "MSExchangeRepl" $eventId $categoryId `
					 "Information" $null $messageArgs
}

function Write-HAAppLogWarningEvent(
	[Parameter(Mandatory=$true)] [string] $eventId, # eg: "C0041011"
	[Parameter(Mandatory=$true)] [int] $categoryId,
	[Object[]] $messageArgs)
{
	Write-LocalizedEventLog "Application" "MSExchangeRepl" $eventId $categoryId `
					 "Warning" $null $messageArgs
}

function Write-HAAppLogErrorEvent(
	[Parameter(Mandatory=$true)] [string] $eventId, # eg: "C0041011"
	[Parameter(Mandatory=$true)] [int] $categoryId,
	[Object[]] $messageArgs)
{
	Write-LocalizedEventLog "Application" "MSExchangeRepl" $eventId $categoryId `
					 "Error" $null $messageArgs
}

function Write-LocalizedEventLog( 
	[Parameter(Mandatory=$true)] [string] $logName,
    [Parameter(Mandatory=$true)] [string] $sourceName,
	[Parameter(Mandatory=$true)] [string] $eventId, # eg: "C0041011"
	[Parameter(Mandatory=$true)] [int] $categoryId,
	[Parameter(Mandatory=$true)] [System.Diagnostics.EventLogEntryType] $entryType,
	[Byte[]] $data,
	[Object[]] $messageArgs)
{
	# parse the eventId into an Int64 first.
	[Int64]$id = [Int64]::Parse($eventId, [System.Globalization.NumberStyles]::HexNumber)
	[CheckHADatabaseRedundancy.EventLogger]::WriteLocalizedEvent( `
		$logName, $sourceName, $id , $categoryId, $entryType, $data, $messageArgs)
}

# Common function to run a scriptblock, log any error that occurred, and return 
# a boolean to indicate whether it was successful or not.
# NOTE: ErrorActionPreference of "Stop" is used to catch all errors.
#
# Optional parameters:
#
# 	cleanupCommand 
#		This scriptblock will be executed with ErrorActionPreference of "Continue", 
#		if an error occurred while running $runCommand.
#
#	throwOnError
#		If true, the error from $runCommand will be rethrown. Otherwise 'false' is returned on error.
#
#	silentOnErrors
#		If true, the error from $runCommand will not be logged via Log-ErrorRecord (i.e. Write-Error)
function TryExecute-ScriptBlock ([ScriptBlock]$runCommand, [ScriptBlock]$cleanupCommand={}, [bool]$throwOnError=$false, [bool]$silentOnErrors=$false)
{
	# Run the following in a separate script block so that we can change
    # ErrorActionPerefence without affecting the rest of the script.
	&{
		$ErrorActionPreference = "Stop"
		[bool]$success = $false;
		
		try
		{
			$ignoredObjects = @(&$runCommand)
			$success = $true;
		}
		catch
		{
			# Any error will end up in this catch block
			# For some reason, PS does not write out any errors unless I use this
			# scriptblock with "Continue" ErrorActionPreference.
			&{
				$ErrorActionPreference = "Continue"
				
				if (!$silentOnErrors)
				{
					Log-ErrorRecord $_
				}
				
				# Run the cleanup scriptblock
				$ignoredObjects = @(&$cleanupCommand)
			}
			
			if ($throwOnError)
			{
				throw
			}
		}
		# Curious PS behavior: It appears that 'return' trumps 'throw', so don't return...
		if (!$throwOnError -or $success)
		{
			return $success
		}		
	}
}

# Sleep for the specified duration (in seconds)
function Sleep-ForSeconds ( [int]$sleepSecs )
{
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0044 -f $sleepSecs)
	Start-Sleep $sleepSecs
}

# Common function to retrieve the current UTC time string
function Get-CurrentTimeString
{
	return [DateTime]::UtcNow.ToString("[HH:mm:ss.fff UTC]")
}

# Common function for verbose logging
function Log-Verbose ( [string]$msg )
{
	$timeStamp = Get-CurrentTimeString
	
	Write-Verbose "$timeStamp $msg"
}

# Common function for warning logging
function Log-Warning ( [string]$msg )
{
	$timeStamp = Get-CurrentTimeString
	Write-Warning "$timeStamp $msg"
}

# Common function for error logging
function Log-Error ( [string]$msg, [switch]$Stop)
{
	$timeStamp = Get-CurrentTimeString

	if (!$Stop)
	{
		Write-Error "$timeStamp $msg"
	}
	else
	{
		Write-Error "$timeStamp $msg" -ErrorAction:Stop
	}
}

# Common function for logging an error, given an ErrorRecord
function Log-ErrorRecord( [System.Management.Automation.ErrorRecord] $errRecord, [switch]$Stop )
{
	# Trim the message so it will not display the "ErrorActionPreference is set to Stop" message
	#
	$failedMessage = $errRecord.ToString()
	if ($failedMessage.IndexOf("ErrorActionPreference") -ne -1)
    {
    	$failedMessage = $failedMessage.Substring($failedMessage.IndexOf("set to Stop: ") + 13)
	}
	$failedMessage = $failedMessage -replace "`r"
	$failedMessage = $failedMessage -replace "`n"
	$failedCommand = $errRecord.InvocationInfo.MyCommand
	Log-Error ($CheckDatabaseRedundancy_LocalizedStrings.res_0045 -f $failedCommand,$failedMessage) -Stop:$Stop
}

# Shuffles objects coming from the input pipeline.
# NOTE: This method only works when invoked via a pipeline.
function Shuffle-Objects(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)] $inputData
)
{
	Begin
	{
		$inputDataList = @()
	}
	
	Process
	{
		# build the input list first
		$inputDataList += $inputData
	}
	
	End
	{
		# now shuffle the contents of the input list
		
		$len = $inputDataList.Length
		for ([int] $i = 0; $i -lt $len; $i++)
		{
			# pick the next random number
			[int]$randomIndex = Get-Random -Minimum:$i -Maximum:$len
			# swap the values
			$temp = $inputDataList[$i]
			$inputDataList[$i] = $inputDataList[$randomIndex]
			$inputDataList[$randomIndex] = $temp
		}
		
		foreach ($element in $inputDataList)
		{
			# send each element to the output pipeline
			$element
		}
	}
}



# This will send a mail message to the specified recipients.
#
# Based on Send-Mail function from DatacenterHealthCommonLibrary.ps1 (service engineering scripts).
#
# Here we will create two SMTP clients: (1) datacenter client is configured to send mails from 
# production datacenter environment (based on Send-Mail function from DatacenterHealthCommonLibrary.ps1),
# (2) CorpNet client is configured to send mails from Topobuilder machines inside a CORPNET.
#
function Send-HANotificationMail(
    [string]$title, 
    [string]$body, 
    [string[]]$attachments,
    [string]$from,
    [string[]]$tos, 
    [string[]]$ccs, 
    [string]$pri = "Normal",
    [int]$maxRetryAttempts = 2)
{
	if ($script:IsDataCenterLibraryPresent)
	{
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0046 -f "send-mail")
		Set-StrictMode -Off
		send-mail -title:$title -body:$body -from:$from -tos:$tos -ccs:$ccs -attachments:$attachments -pri:$pri
		Set-StrictMode -Version 2.0
		return
	}
	else
	{
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0047 -f "Send-HANotificationMailCorpHub")
		[bool]$sent = Send-HANotificationMailCorpHub -title:$title -body:$body -attachments:$attachments -from:$from `
													 -tos:$tos -ccs:$ccs -pri:$pri -maxRetryAttempts:$maxRetryAttempts
		
		if (!$sent){
			Log-Error ($CheckDatabaseRedundancy_LocalizedStrings.res_0048 )
		}
	}    
}

function Get-HubServers
{
	Get-ExchangeServer | where { $_.IsHubTransportServer }
}

# Build a list of SMTP clients that can send mail to a local hub server.
# 
function Build-HubSmtpClients
{
	#FUTURE: We should return a list of hostnames and port and let the caller iterate over them if failure...
	# Also, try to choose a server in the same site as we are running
	# In production we will use the "send-mail" function provided by the svc engineering team
	
	$hubServers = (Get-HubServers | Shuffle-Objects )
	if (!$hubServers) {
		Log-Error ($CheckDatabaseRedundancy_LocalizedStrings.res_0049 )
		return
	}
	
	foreach ($hubServer in @($hubServers))
	{
		$smtpClient = New-Object System.Net.Mail.SmtpClient($hubServer.Fqdn)
		$smtpClient.UseDefaultCredentials = $true
		Write-Output $smtpClient
	}
}

function Get-SmtpClients ()
{	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0050 -f "Get-SmtpClients")
	$clients = @()
		$clients += Build-HubSmtpClients
	return $clients
}

# Build a Mail message
#
function Build-MailMsg(
    [string]$title, 
    [string]$body, 
    [string[]]$attachments,
    [string]$from,
    [string[]]$tos, 
    [string[]]$ccs, 
    [string]$pri = "Normal")
{
    $mailMessage = New-Object System.Net.Mail.MailMessage
    $mailMessage.Body = $body
    $mailMessage.Priority = $pri
    $mailMessage.Subject = $title
    $mailMessage.From = New-Object System.Net.Mail.MailAddress($from);

    # Add attachments
    if ($attachments)
    {
        foreach ($attachment in @($attachments))
        {
            if ( Test-Path $attachment )
            {
                $data = New-Object System.Net.Mail.Attachment -ArgumentList $attachment, 'Application/Octet'
                [void]$mailMessage.Attachments.Add($data);
            }
        }
    }

    foreach ($to in @($tos))
    {
        [void]$mailMessage.To.Add($to)
    }
    if ($ccs)
    {
        foreach ($cc in @($ccs))
        {
            [void]$mailMessage.CC.Add($cc)
        }
    }

	return $mailMessage
}


# Send an email message
# Return $true if an SMTP host was contacted and the mail transmitted.
# There is no guarantee that the mail will get through.
# This is cloned from Send-NotificationMail in DatacenterSvcEngCommonLibrary.ps1, but made simpler
# so I could add function and simplify it at the same time.
#
function Send-HANotificationMailCorpHub(
    [string]$title, 
    [string]$body, 
    [string[]]$attachments,
    [string]$from,
    [string[]]$tos,   
    [string[]]$ccs, 
    [string]$pri = "Normal",
    [int]$maxRetryAttempts = 2)
{
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0051 -f $from,$pri,"Send-HANotificationMailCorpHub")
	$clients = @(Get-SmtpClients)
	if (!$clients)
	{
		Log-Error ($CheckDatabaseRedundancy_LocalizedStrings.res_0052 -f "Get-SmtpClients")
		return $false
	}
	
	$mailMessage = Build-MailMsg -title $title -body $body -attachments $attachments `
								 -from $from -tos $tos -ccs $ccs -pri $pri
	if (!$mailMessage) {
		Log-Error ($CheckDatabaseRedundancy_LocalizedStrings.res_0052 -f "Build-MailMsg")
		return $false
	}
		
	try
	{
		foreach ($smtpClient in $clients)
		{
			# Change the timeout for synchronous Send() call to 30 secs
			$smtpClient.Timeout = 30000
			
			$retries = 0
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0053 -f [string]::Join(';',$tos))
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0054 -f $smtpClient.Host,$smtpClient.Port)
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0055 -f $from)
			do {
				try {
					$success = $true
					$smtpClient.Send($mailMessage)  
					Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0056 )
					return $true
				} catch {
					$success = $false
					$retries++
					if ($retries -eq $maxRetryAttempts) {
						Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0057 -f $maxRetryAttempts,$tos)
					} else {
						Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0058 -f $tos)
					}
				}
			} while ((-not $success) -and $retries -lt $maxRetryAttempts)
			
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0059 ) 
		}
	} 
	finally 
	{
		$mailMessage.Dispose()
	}
    return $success
}


function Append-RedundancyInformation(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)] [CheckHADatabaseRedundancy.DatabaseRedundancyEntry] $dbState)
{
	Process
	{
		$totalMins = $dbState.TotalRedDuration.TotalMinutes.ToString("F2")
		$msg = "
Database          : $($dbState.DatabaseName)
Redundancy Count  : $($dbState.CurrentRedundancyCount)
Total Red Minutes : $totalMins"
		$script:report.AppendLine($msg) | Out-Null
		
		foreach ($errRecord in @($dbState.ErrorHistory))
		{
			$timeStr = $errRecord.ErrorTime.ToString("HH:mm:ss.fff UTC")
			$msg = "
$timeStr    :
$($errRecord.GetErrorStringForAlerting())
"
			$script:report.AppendLine($msg) | Out-Null		
		}
	}
}

function Send-SummaryEmail
{
	if (!$SendSummaryMailTos)
	{
		return
	}
	
	[System.Text.StringBuilder]$script:report = New-Object -TypeName System.Text.StringBuilder -ArgumentList 2048
	$states = $script:databaseStateTable.Values
	$dbsWithOneCopy = $states | where { `
			  ($_.CurrentState -eq [CheckHADatabaseRedundancy.AlertState]::Red) -and `
			  ($_.TotalRedDuration.TotalSeconds -gt $ReportRedEventAfterDurationSecs) `
			} | sort -Property DatabaseName
	$dbsWithErrors = $states | where { `
			$_.HasErrorsInHistory -and `
				(	($_.CurrentState -ne [CheckHADatabaseRedundancy.AlertState]::Red) -or `
					($_.TotalRedDuration.TotalSeconds -le $ReportRedEventAfterDurationSecs) `
				)} | sort -Property DatabaseName
	
	[int]$databasesCount = ($states | Measure-Object).Count
	[int]$dbsWithErrorsCount = ($dbsWithErrors | Measure-Object).Count
	[int]$dbsOneCopyCount = ($dbsWithOneCopy | Measure-Object).Count
	
	if ( ($dbsWithErrorsCount -eq 0) -and ($dbsOneCopyCount -eq 0) )
	{
		# No need to send an email to report that everything is healthy
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0060 )
		return
	}
	
	[string]$priority = "Normal"
	$dbsWithOneCopyNames = ( $dbsWithOneCopy | select -ExpandProperty DatabaseName )
	$dbsWithErrorsNames = ( $dbsWithErrors | select -ExpandProperty DatabaseName )
	[string]$dbsWithOneCopyNamesStr = $null
	[string]$dbsWithErrorsNamesStr = $null
	if ($dbsWithOneCopyNames)
	{
		$dbsWithOneCopyNamesStr = [string]::Join(", ", $dbsWithOneCopyNames)
	}
	if ($dbsWithErrorsNames)
	{
		$dbsWithErrorsNamesStr = [string]::Join(", ", $dbsWithErrorsNames)
	}
	
	[string] $msg = "
***************************************
Database Redundancy Report
$((Get-Date).DateTime)
***************************************"

	$script:report.AppendLine($msg) | Out-Null
	
	if ($PSCmdlet.ParameterSetName  -eq "Database" )
	{
		$msg = "Database                      : $MailboxDatabaseName"
		$script:report.AppendLine($msg) | Out-Null
	}
	elseif ($PSCmdlet.ParameterSetName  -eq "Server" )
	{
		$msg = "Server                        : $MailboxServerName"
		$script:report.AppendLine($msg) | Out-Null
	}
	
	$msg = `
"DatabaseCount                 : $databasesCount
DbsCountWithLowRedundancy     : $dbsOneCopyCount
DbsCountWithErrors            : $dbsWithErrorsCount

DatabasesWithLowRedundancy    : $dbsWithOneCopyNamesStr
DatabasesWithErrors           : $dbsWithErrorsNamesStr"
	
	$script:report.AppendLine($msg) | Out-Null
	
	if ($dbsOneCopyCount -gt 0)
	{
		# mark email as urgent
		$priority = "High"
		$msg = "
=================================================================
Databases with low redundancy ( < $AtLeastNCopies copies)
================================================================="
		$script:report.AppendLine($msg) | Out-Null
		@($dbsWithOneCopy) | Append-RedundancyInformation
	}
	
	if ($dbsWithErrorsCount -gt 0)
	{
		$msg = "
=================================================================
Databases with errors ( >= $AtLeastNCopies copies)
================================================================="
		$script:report.AppendLine($msg) | Out-Null
		@($dbsWithErrors) | Append-RedundancyInformation
	}
	
	# Create the email subject
	[string]$title = "DB Redundancy: "
	if ($script:dagName)
	{
		$title += "$($script:dagName): "
	}
	if ($PSCmdlet.ParameterSetName  -eq "Database" )
	{
		$title += "$MailboxDatabaseName - "
	}
	elseif ($PSCmdlet.ParameterSetName  -eq "Server" )
	{
		$title += "$MailboxServerName - "
	}
	
	if ($dbsOneCopyCount -gt 0)
	{
		$durationMins = [TimeSpan]::FromSeconds($ReportRedEventAfterDurationSecs).TotalMinutes
		if ($dbsOneCopyCount -eq 1)
		{
			$title += "1 DB has less than $AtLeastNCopies copies for more than $durationMins mins"
		}
		else
		{
			$title += "$dbsOneCopyCount DBs have less than $AtLeastNCopies copies for more than $durationMins mins"
		}
	}
	elseif ($dbsWithErrorsCount -gt 0)
	{
		if ($dbsWithErrorsCount -eq 1)
		{
			$title += "1 DB has had errors in the past hour"
		}
		else
		{
			$title += "$dbsWithErrorsCount DBs have had errors in the past hour"
		}
	}
	else
	{
		$title += "All DBs have been sufficiently redundant for the past hour"
	}
	
		
	# send the email
	Send-HANotificationMail -title:$title -body:($script:report.ToString()) -from:$SummaryMailFrom -tos:$SendSummaryMailTos -pri:$priority
}



###################################################################
###  Entry point for the script itself
###################################################################

function RunOnce
{
	$script:outputObjects = $null
	
	# Run each iteration of the check in a separate script block so that any errors
	# can be trapped and the entire script block exits.
	$checkCmd = { $script:outputObjects = RunOnceInternal }
	[bool]$success = TryExecute-ScriptBlock -runCommand $checkCmd
	if ($success)
	{
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0061 -f $script:iteration)
	}
	else
	{
		Log-Error ($CheckDatabaseRedundancy_LocalizedStrings.res_0062 -f $script:iteration)
	}
	
	# send to the output pipeline
	$script:outputObjects
	
	if ($script:ExitOut)
	{
		Log-Verbose "Script is now exiting..."
		Exit
	}
}

function RunOnceInternal
{
	$script:oneIterationStopwatch.Reset()
	$script:oneIterationStopwatch.Start()
	
	$script:iteration++
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0063 -f $script:iteration)
	
	# The databases being monitored
	[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $mdbs = @()
	
	# Check if the node is clustered (i.e. participating member of a DAG)
	if ($MonitoringContext -and
		![Microsoft.Exchange.Cluster.Replay.DagTaskHelperPublic]::IsLocalNodeClustered())
	{
		Log-Verbose "The local node is not a participating member of a database availability group (DAG). Skipping running the monitoring checks."
		$script:ExitOut = $true
		Exit
	}
	
	# Lookup the specified database
	if ($PSCmdlet.ParameterSetName  -eq "Database" )
	{
		$mdb = Get-MailboxDatabase $MailboxDatabaseName
		if ($mdb)
		{
			if ($mdb.Recovery)
			{
				Log-Error ($CheckDatabaseRedundancy_LocalizedStrings.res_0064 -f $mdb,"non-recovery") -Stop
				return
			}
			
			# We will check databases even if they only have 1 configured copy
			$mdbs += $mdb
		}
		else
		{
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0065 -f $MailboxDatabaseName)
		}
	}
	# Lookup all databases on the specified server
	elseif ($PSCmdlet.ParameterSetName  -eq "Server" )
	{
		$server = Get-MailboxServer $MailboxServerName
		if ($server)
		{
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0066 -f $MailboxServerName)
			$script:mailboxServer = $server
			$allMdbs = @( Get-MailboxDatabase -Server $server )
			
			if ($SkipDatabasesRegex)
			{
				# filter out the DBs matching the regex specified
				Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0067 -f $SkipDatabasesRegex)
				$mdbs = @( $allMdbs | where { ($_.Name -inotmatch $SkipDatabasesRegex) -and (!$_.Recovery) } )
			}
			else
			{
				# no database name filter specified, so check against all
				$mdbs = @( $allMdbs | where { !$_.Recovery } )
			}
		}
		else
		{
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0068 -f $MailboxServerName)
		}
	}
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0069 -f $mdbs.Length)
	
	# perform the check
	if ($mdbs.Length -gt 0)
	{
		Check-Databases $mdbs $PSCmdlet.ParameterSetName
	}
	else
	{
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0070 -f "Check-Databases")
	}
	
	$script:oneIterationStopwatch.Stop()
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0071 -f $script:iteration,$script:oneIterationStopwatch.Elapsed.TotalMilliseconds)
	
}

# This function returns true if you can remove the copy without losing redundancy and false otherwise
function Check-DatabaseRedundancyForCopyRemoval(
	[string] $databaseName = $(throw ($CheckDatabaseRedundancy_LocalizedStrings.res_0072 -f "Check-DatabaseRedundancyForCopyRemoval","-databaseName")),
	[string] $serverName = $(throw ($CheckDatabaseRedundancy_LocalizedStrings.res_0072 -f "Check-DatabaseRedundancyForCopyRemoval","-serverName")))
{
	if ( -not $databaseName ) { throw ($CheckDatabaseRedundancy_LocalizedStrings.res_0073 -f "Check-DatabaseRedundancyForCopyRemoval","-databaseName") }
	if ( -not $serverName ) { throw ($CheckDatabaseRedundancy_LocalizedStrings.res_0073 -f "Check-DatabaseRedundancyForCopyRemoval","-serverName") }
	
	$MailboxDatabaseName = $databaseName
	$SkipEventLogging = $true
	$MonitoringContext = $false
	
	$databases = @( Get-MailboxDatabase $MailboxDatabaseName )
	
	Populate-DatabasesTable $databases

	# find the servers to check copy statuses on
	[String[]]$servers = Get-ServersForDatabases $databases
	
	if ($servers.Length -lt 2)
	{
		# Normally we should not get here, since we're only checking replicated DBs, which means
		# we should have at least 2 distinct servers. However, this can happen if copies are 
		# removed while this script is running.
		Log-Warning ($CheckDatabaseRedundancy_LocalizedStrings.res_0074 -f $servers.Length,"Check-DatabaseRedundancyForCopyRemoval","Get-ServersForDatabases")
	}
	
	# get the status results and index them by database name
	$script:databaseToStatusTable.Clear()
	$script:databaseToStatusTable = Get-CopyStatusFromAllServers $servers "Database" | `
									Group-Object -AsHashTable -Property DatabaseName
	
	# look up the cluster node status for the DAG
	Populate-ClusterNodeStatus	

	# Simulate copy removal
	[UInt32] $regularAtLeastNCopies = $AtLeastNCopies
	[bool] $foundCopy = $false
	for ($i = 0; $i -lt $script:databaseToStatusTable[$MailboxDatabaseName].Count; $i++)
	{
		if ( $script:databaseToStatusTable[$MailboxDatabaseName][$i].Name -eq "$MailboxDatabaseName\$serverName" )
		{
			if ( $script:databaseToStatusTable[$MailboxDatabaseName][$i].ActiveCopy )
			{
				$AtLeastNCopies = $regularAtLeastNCopies + 1
				Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0075 -f $MailboxDatabaseName,$serverName,$AtLeastNCopies,"Check-DatabaseRedundancyForCopyRemoval")
			}
			else
			{
				Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0076 -f $MailboxDatabaseName,$serverName,"Check-DatabaseRedundancyForCopyRemoval") 
				$script:databaseToStatusTable[$MailboxDatabaseName].RemoveAt($i)
			}
			$foundCopy = $true
			break
		}
	}
	
	if ( -not $foundCopy )
	{
		throw ($CheckDatabaseRedundancy_LocalizedStrings.res_0077 -f $MailboxDatabaseName,$serverName,"Check-DatabaseRedundancyForCopyRemoval")
	}
	
	$status = Check-DatabaseRedundancy $MailboxDatabaseName
	[bool] $result = $status.CurrentState -eq "Green"

	$AtLeastNCopies = $regularAtLeastNCopies
	
	return $result
}

function Main
{
	if ($PSCmdlet.ParameterSetName  -eq "DotSourceMode")
	{
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0078 -f "-DotSourceMode")
		return
	}
	
	# Ensure this table is cleared at script startup
	# Other hashtables get cleared every iteration of RunOnce.
	$script:databaseStateTable.Clear()
	
	# Validate the email parameters
	if ($SendSummaryMailTos -and !$SummaryMailFrom)
	{
		Log-Error ($CheckDatabaseRedundancy_LocalizedStrings.res_0079 -f "-SummaryMailFrom","-SendSummaryMailTos")
		# Let monitoring continue anyway...
	}
	
	# Check if the node is clustered (i.e. participating member of a DAG)
	if ($MonitoringContext -and 
		![Microsoft.Exchange.Cluster.Replay.DagTaskHelperPublic]::IsLocalNodeClustered())
	{
		Log-Verbose "The local node is not a participating member of a database availability group (DAG). Skipping running the monitoring checks."
		return
	}
	
	if (!$MonitoringContext)
	{
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0080 )
		RunOnce
		Send-SummaryEmail
		return
	}	

	# We are in the MonitoringContext

	[bool] $keepRunning = $true
	[System.Diagnostics.Stopwatch] $overallScriptStopwatch = New-Object -TypeName System.Diagnostics.Stopwatch
	$overallScriptStopwatch.Reset()
	$overallScriptStopwatch.Start()
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0081 )
	
	while ($keepRunning)
	{
		RunOnce		
		
		# decide if we should run the next iteration
		if (($TerminateAfterDurationSecs -eq -1) -or `
			($TerminateAfterDurationSecs -eq 0))
		{
			# infinite duration specified
			$keepRunning = $true
		}
		else
		{
			[double]$lastIterationMsecs = $script:oneIterationStopwatch.Elapsed.TotalMilliseconds
			[double]$timeLeftMsecs = [double]($TerminateAfterDurationSecs * 1000) - $overallScriptStopwatch.Elapsed.TotalMilliseconds
			# Is there enough time left for a (sleep + RunOnce) ?
			if ( ([double]($SleepDurationBetweenIterationsSecs * 1000) + $lastIterationMsecs) -lt $timeLeftMsecs )
			{
				$keepRunning = $true
			}
			else
			{
				$keepRunning = $false
				break
			}
		}
				
		Sleep-ForSeconds $SleepDurationBetweenIterationsSecs
	}
	
	Send-SummaryEmail
}

$Command = $MyInvocation.MyCommand
Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0082 -f $Command.Path)
# The command below is useful to see what parameters are defined in this script cmdlet.
# $Command | fl Path, CommandType, Parameters, ParameterSets

# Get the code compilation out of the way
Prepare-DatabaseRedundancyEntryDefinition

LoadExchangeSnapin

# In datacenter configurations we can use libraries provided by service engineering
$InstallPath = (Get-ItemProperty -path 'HKLM:SOFTWARE\Microsoft\ExchangeServer\v14\Setup').MsiInstallPath.Trim().TrimEnd("\")
$DatacenterLibraryPath = "$InstallPath\DataCenter"
$SvcLibaryFileName = "DatacenterHealthCommonLibrary.ps1"
$ServiceCommonLib = "$DatacenterLibraryPath\$SvcLibaryFileName"
$script:IsDataCenterLibraryPresent = Test-Path $ServiceCommonLib
if ($script:IsDataCenterLibraryPresent)
{
	if ($SendSummaryMailTos)
	{
		# Get a send-mail function
		Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0083 -f $ServiceCommonLib)
		# The common lib doesn't use clean practices so we have to avoid strict mode
		Set-StrictMode -Off
		. $ServiceCommonLib
		Set-StrictMode -Version 2.0
	}
}
else
{
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0084 -f $DatacenterLibraryPath,$SvcLibaryFileName)
}


Main



# SIG # Begin signature block
# MIIazwYJKoZIhvcNAQcCoIIawDCCGrwCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUGJP7T4uxtrfQW6lK9Ku8csfo
# 83WgghWCMIIEwzCCA6ugAwIBAgITMwAAAG9lLVhtBxFGKAAAAAAAbzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTUwMzIwMTczMjAy
# WhcNMTYwNjIwMTczMjAyWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkMwRjQtMzA4Ni1ERUY4MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAz+ZtzcEqza6o
# XtiVTy0DQ0dzO7hC0tBXmt32UzZ31YhFJGrIq9Bm6YvFqg+e8oNGtirJ2DbG9KD/
# EW9m8F4UGbKxZ/jxXpSGqo4lr/g1E/2CL8c4XlPAdhzF03k7sGPrT5OaBfCiF3Hc
# xgyW0wAFLkxtWLN/tCwkcHuWaSxsingJbUmZjjo+ZpWPT394G2B7V8lR9EttUcM0
# t/g6CtYR38M6pR6gONzrrar4Q8SDmo2XNAM0BBrvrVQ2pNQaLP3DbvB45ynxuUTA
# cbQvxBCLDPc2Ynn9B1d96gV8TJ9OMD8nUDhmBrtdqD7FkNvfPHZWrZUgNFNy7WlZ
# bvBUH0DVOQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFPKmSSl4fFdwUmLP7ay3eyA0
# R9z9MB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAI2zTLbY7A2Hhhle5ADnl7jVz0wKPL33VdP08KCvVXKcI5e5
# girHFgrFJxNZ0NowK4hCulID5l7JJWgnJ41kp235t5pqqz6sQtAeJCbMVK/2kIFr
# Hq1Dnxt7EFdqMjYxokRoAZhaKxK0iTH2TAyuFTy3JCRdu/98U0yExA3NRnd+Kcqf
# skZigrQ0x/USaVytec0x7ulHjvj8U/PkApBRa876neOFv1mAWRDVZ6NMpvLkoLTY
# wTqhakimiM5w9qmc3vNTkz1wcQD/vut8/P8IYw9LUVmrFRmQdB7/u72qNZs9nvMQ
# FNV69h/W4nXzknQNrRbZEs+hm63SEuoAOyMVDM8wggTsMIID1KADAgECAhMzAAAA
# ymzVMhI1xOFVAAEAAADKMA0GCSqGSIb3DQEBBQUAMHkxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xIzAhBgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBMB4XDTE0MDQyMjE3MzkwMFoXDTE1MDcyMjE3MzkwMFowgYMxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIx
# HjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJZxXe0GRvqEy51bt0bHsOG0ETkDrbEVc2Cc66e2bho8
# P/9l4zTxpqUhXlaZbFjkkqEKXMLT3FIvDGWaIGFAUzGcbI8hfbr5/hNQUmCVOlu5
# WKV0YUGplOCtJk5MoZdwSSdefGfKTx5xhEa8HUu24g/FxifJB+Z6CqUXABlMcEU4
# LYG0UKrFZ9H6ebzFzKFym/QlNJj4VN8SOTgSL6RrpZp+x2LR3M/tPTT4ud81MLrs
# eTKp4amsVU1Mf0xWwxMLdvEH+cxHrPuI1VKlHij6PS3Pz4SYhnFlEc+FyQlEhuFv
# 57H8rEBEpamLIz+CSZ3VlllQE1kYc/9DDK0r1H8wQGcCAwEAAaOCAWAwggFcMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQfXuJdUI1Whr5KPM8E6KeHtcu/
# gzBRBgNVHREESjBIpEYwRDENMAsGA1UECxMETU9QUjEzMDEGA1UEBRMqMzE1OTUr
# YjQyMThmMTMtNmZjYS00OTBmLTljNDctM2ZjNTU3ZGZjNDQwMB8GA1UdIwQYMBaA
# FMsR6MrStBZYAck3LjMWFrlMmgofMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9j
# cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY0NvZFNpZ1BDQV8w
# OC0zMS0yMDEwLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljQ29kU2lnUENBXzA4LTMx
# LTIwMTAuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQB3XOvXkT3NvXuD2YWpsEOdc3wX
# yQ/tNtvHtSwbXvtUBTqDcUCBCaK3cSZe1n22bDvJql9dAxgqHSd+B+nFZR+1zw23
# VMcoOFqI53vBGbZWMrrizMuT269uD11E9dSw7xvVTsGvDu8gm/Lh/idd6MX/YfYZ
# 0igKIp3fzXCCnhhy2CPMeixD7v/qwODmHaqelzMAUm8HuNOIbN6kBjWnwlOGZRF3
# CY81WbnYhqgA/vgxfSz0jAWdwMHVd3Js6U1ZJoPxwrKIV5M1AHxQK7xZ/P4cKTiC
# 095Sl0UpGE6WW526Xxuj8SdQ6geV6G00DThX3DcoNZU6OJzU7WqFXQ4iEV57MIIF
# vDCCA6SgAwIBAgIKYTMmGgAAAAAAMTANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZIm
# iZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQD
# EyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMTAwODMx
# MjIxOTMyWhcNMjAwODMxMjIyOTMyWjB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALJyWVwZMGS/HZpgICBC
# mXZTbD4b1m/My/Hqa/6XFhDg3zp0gxq3L6Ay7P/ewkJOI9VyANs1VwqJyq4gSfTw
# aKxNS42lvXlLcZtHB9r9Jd+ddYjPqnNEf9eB2/O98jakyVxF3K+tPeAoaJcap6Vy
# c1bxF5Tk/TWUcqDWdl8ed0WDhTgW0HNbBbpnUo2lsmkv2hkL/pJ0KeJ2L1TdFDBZ
# +NKNYv3LyV9GMVC5JxPkQDDPcikQKCLHN049oDI9kM2hOAaFXE5WgigqBTK3S9dP
# Y+fSLWLxRT3nrAgA9kahntFbjCZT6HqqSvJGzzc8OJ60d1ylF56NyxGPVjzBrAlf
# A9MCAwEAAaOCAV4wggFaMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFMsR6MrS
# tBZYAck3LjMWFrlMmgofMAsGA1UdDwQEAwIBhjASBgkrBgEEAYI3FQEEBQIDAQAB
# MCMGCSsGAQQBgjcVAgQWBBT90TFO0yaKleGYYDuoMW+mPLzYLTAZBgkrBgEEAYI3
# FAIEDB4KAFMAdQBiAEMAQTAfBgNVHSMEGDAWgBQOrIJgQFYnl+UlE/wq4QpTlVnk
# pDBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtp
# L2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEE
# SDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2Nl
# cnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDANBgkqhkiG9w0BAQUFAAOCAgEAWTk+
# fyZGr+tvQLEytWrrDi9uqEn361917Uw7LddDrQv+y+ktMaMjzHxQmIAhXaw9L0y6
# oqhWnONwu7i0+Hm1SXL3PupBf8rhDBdpy6WcIC36C1DEVs0t40rSvHDnqA2iA6VW
# 4LiKS1fylUKc8fPv7uOGHzQ8uFaa8FMjhSqkghyT4pQHHfLiTviMocroE6WRTsgb
# 0o9ylSpxbZsa+BzwU9ZnzCL/XB3Nooy9J7J5Y1ZEolHN+emjWFbdmwJFRC9f9Nqu
# 1IIybvyklRPk62nnqaIsvsgrEA5ljpnb9aL6EiYJZTiU8XofSrvR4Vbo0HiWGFzJ
# NRZf3ZMdSY4tvq00RBzuEBUaAF3dNVshzpjHCe6FDoxPbQ4TTj18KUicctHzbMrB
# 7HCjV5JXfZSNoBtIA1r3z6NnCnSlNu0tLxfI5nI3EvRvsTxngvlSso0zFmUeDord
# EN5k9G/ORtTTF+l5xAS00/ss3x+KnqwK+xMnQK3k+eGpf0a7B2BHZWBATrBC7E7t
# s3Z52Ao0CW0cgDEf4g5U3eWh++VHEK1kmP9QFi58vwUheuKVQSdpw5OPlcmN2Jsh
# rg1cnPCiroZogwxqLbt2awAdlq3yFnv2FoMkuYjPaqhHMS+a3ONxPdcAfmJH0c6I
# ybgY+g5yjcGjPa8CQGr/aZuW4hCoELQ3UAjWwz0wggYHMIID76ADAgECAgphFmg0
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
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBLcwggSz
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAAAymzVMhI1xOFV
# AAEAAADKMAkGBSsOAwIaBQCggdAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFEdK
# lqNVqywSb1UZEuGnO9qS1SQ+MHAGCisGAQQBgjcCAQwxYjBgoDiANgBDAGgAZQBj
# AGsARABhAHQAYQBiAGEAcwBlAFIAZQBkAHUAbgBkAGEAbgBjAHkALgBwAHMAMaEk
# gCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEB
# AQUABIIBAA5lM/q92MmTy/7d5SuCDY1Rv/ojhsqFqksMkMKTaF4HpaRhel9IxUhi
# 4YYJOL6R2p+dDY2ARjSNinq28uQKKYSDgPI1vCb3/LN9rvEM9HSBgvf1fLAA6tYW
# JqqzQRhtPzACZ2mlnEBQ2veQeYWPw07T0M267FzrK98RtdIp0+TnHisdO7ECeple
# R9JGsPPZd9H5i1JUd4pagJfJ2BqvBwOsA3x/SVWx8FRm3EKtgkIrefRi8KHXCPWG
# 29iYcwYctpD+NCYAllQwIupSSW1miD+spEnOz8+qfZm2i6PI4NnfMxfg1jatzT7P
# Yhc4cNJn1F2DBZeYnJ2QVVdhfGw5Fi2hggIoMIICJAYJKoZIhvcNAQkGMYICFTCC
# AhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
# BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEh
# MB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAb2UtWG0HEUYo
# AAAAAABvMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwG
# CSqGSIb3DQEJBTEPFw0xNTA0MTAwMjU3MzFaMCMGCSqGSIb3DQEJBDEWBBQnASFA
# P7smhM2Kv/T/TlTEzsKWqTANBgkqhkiG9w0BAQUFAASCAQCQwCr3VAyb9tkyc5c5
# 0W2bZmAt6QemwvHHYx3JUCXmeAV3OHVAFF4JaonNU1MR6ddpRLhD19IR1BjrD3LM
# d/Ikawd5G1TSyJfsH0oimaKjxYfTmXP/9umkPlM6mKjriHx4HIXZp9tv1vvTLSd0
# 77kYP6KJ2WeuHwwDfwAfTnEeWb/oE+BT0/HvO/3kEGRmpg0Em8HAH2qevNhhEqVy
# fpOPOM5kXdLzmm3vKRS68z8OoKEzhQbpKHf/p17T/KQmZfAG/honoJf67iGpyUWi
# aEVBsll+P+gvgmXxDuXsOjkIKmpdgDC8C4+/KY0GXDfpYqxloOzglf9LXAfcF+31
# fod6
# SIG # End signature block
