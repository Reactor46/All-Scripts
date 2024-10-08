<#
.EXTERNALHELP RedistributeActiveDatabases-help.xml
#>

# Copyright (c) Microsoft Corporation. All rights reserved.
#
#   Attempts to redistribute active databases in the specified DAG. 
# 	If required, it will also try to balance active DBs across sites.
#
# This is a perfectly balanced distribution of 16 Databases with 4 copies each across 4 servers:
#
#	ServerName  ActiveDbs    PassiveDbs     MountedDbs    DismountedDbs PreferenceCountList
#	----------  ---------    ----------     ----------    ------------- -------------------
#	EXCH-E-552          4            12              4                0 {4, 4, 4, 4}
#	EXCH-D-668          4            12              4                0 {4, 4, 4, 4}
#	EXCH-D-796          4            12              4                0 {4, 4, 4, 4}
#	EXCH-D-058          4            12              4                0 {4, 4, 4, 4}
#
#
# This is an example of an uneven distribution (even the Activation Preferences are not balanced):
#
#	ServerName   ActiveDbs   PassiveDbs     MountedDbs    DismountedDbs PreferenceCountList
#	----------   ---------   ----------     ----------    ------------- -------------------
#	EXCH-E-552           5           11              5                0 {4, 4, 3, 5}
#	EXCH-D-668           1           15              1                0 {1, 8, 6, 1}
#	EXCH-D-796          12            4             12                0 {13, 2, 1, 0}
#	EXCH-D-058           1           15              1                0 {1, 1, 5, 9}
#

[CmdletBinding(DefaultParametersetName="Common")]
param(
	
	# DagName can be $null or omitted, in which case we will try to lookup the local DAG
	[Parameter(Position=0)] 
	[System.Management.Automation.ValidateNotNullOrEmptyAttribute()]
	[string] $DagName,
	
	# Tries to move DBs to their most preferred copy w/o regard for AD site balance
	[Parameter(ParameterSetName="BalanceDbsByActivationPreference",Mandatory=$true)] 
	[switch] $BalanceDbsByActivationPreference = $false,
	
	# Tries to balance DBs evenly across the servers w/o looking at Activation Preference or site balance.
	# When this parameter is used, the script automatically ignores servers that are blocked from automatic
	# activation (i.e. DatabaseAutoActivationPolicy = Blocked), and also ignores servers that are *DOWN*
	# according to windows clustering.
	[Parameter(ParameterSetName="BalanceDbsIgnoringActivationPreference",Mandatory=$true)] 
	[switch] $BalanceDbsIgnoringActivationPreference = $false,
	
	#------------------------------------------------------------------------------------------------
	#	Parameters for site balancing
	#------------------------------------------------------------------------------------------------
	# Tries to move DBs to their most preferred copy, while also trying to balance
	# active DBs by site.
	[Parameter(ParameterSetName="BalanceDbsBySiteAndActivationPreference",Mandatory=$true)] 
	[switch] $BalanceDbsBySiteAndActivationPreference = $false,
	
	# Determines the definition of "balanced" active DBs across Sites. Best explained via an example:
	#
	# 	-100 DBs in total with 3 AD Sites 	==> Ideal distribution: 34, 33, 33			[...for instance]
	#	-Deviation 10% = ±3.33 per site 	==> ±4 per site [ceiling is taken] 	
	# 		==> Allowed range per site: 29 <= n <= 38	[floor and ceiling respectively]
	#
	# Think of this as the added load (in terms of active databases) that a site can tolerate.
	#
	[Parameter(ParameterSetName="BalanceDbsBySiteAndActivationPreference")] 
	[double] $AllowedDeviationFromMeanPercentage = 20.0,
	#------------------------------------------------------------------------------------------------
	
	# Tries to balance Activation Preferences across database copies in the specified DAG.
	[Parameter(ParameterSetName="BalanceActivationPreferences",Mandatory=$true)]
	[switch] $BalanceActivationPreferences = $false,
	
	[Parameter(ParameterSetName="BalanceDbsByActivationPreference")]
	[Parameter(ParameterSetName="BalanceDbsBySiteAndActivationPreference")]
	[Parameter(ParameterSetName="BalanceDbsIgnoringActivationPreference")] 
	[Parameter(ParameterSetName="ShuffleActiveDatabases")]
	[Parameter(ParameterSetName="BalanceActivationPreferences")]
	[switch] $ShowFinalDatabaseDistribution = $false,
	
	[Parameter(ParameterSetName="ShowDatabaseCurrentActives",Mandatory=$true)] 
	[switch] $ShowDatabaseCurrentActives = $false,
	
	[Parameter(ParameterSetName="ShowDatabaseDistributionByServer",Mandatory=$true)] 
	[switch] $ShowDatabaseDistributionByServer = $false,
	
	[Parameter(ParameterSetName="ShowDatabaseDistributionBySite",Mandatory=$true)] 
	[switch] $ShowDatabaseDistributionBySite = $false,

	# Useful to "dot-source" this script as a library - call the script as such:
	#	PS D:\Exchange Mailbox\v15\Scripts> . .\RedistributeActiveDatabases.ps1 -DotSourceMode
	[Parameter(ParameterSetName="DotSourceMode",Mandatory=$true)] 
	[switch] $DotSourceMode = $false,
	
	# Useful for test purposes - Shuffles active DBs around to generate a random active DB distribution
	[Parameter(ParameterSetName="ShuffleActiveDatabases",Mandatory=$true)] 
	[switch] $ShuffleActiveDatabases = $false,
	
	# Specify if the balancing should only happen on a server that is the Primary Active Manager.
	[switch] $RunOnlyOnPAM = $false,
	
	# Specify if DB balancing should be logged as an event
	[switch] $LogEvents = $true,
	
	# Should non-replicated DBs be included in the distribution/balancing?
	[switch] $IncludeNonReplicatedDatabases = $false,

	# Should move-suppression checks be skipped.
	[switch] $SkipMoveSuppressionChecks = $false,

	# If used, no changes are made (i.e. no DBs get moved around)
	[switch] $WhatIf = $false,
	
	# If used, causes Moves to request confirmation
	[switch] $Confirm = $true,
	
	# Maximum allowed delta to validate distribution
	[Parameter(ParameterSetName="BalanceDbsByActivationPreference")]
	[Parameter(ParameterSetName="BalanceDbsBySiteAndActivationPreference")]
	[Parameter(ParameterSetName="ShuffleActiveDatabases")]
	[Parameter(ParameterSetName="BalanceDbsIgnoringActivationPreference")] 
	[int] $MaxAllowedDbsDeltaForBalanceValidation = [int]::MaxValue,
	
	# Option to override Content-Indexing health checks when moving databases around
	[Parameter(ParameterSetName="BalanceDbsByActivationPreference")]
	[Parameter(ParameterSetName="BalanceDbsBySiteAndActivationPreference")]
	[Parameter(ParameterSetName="ShuffleActiveDatabases")]
	[Parameter(ParameterSetName="BalanceDbsIgnoringActivationPreference")]
	[switch] $SkipClientExperienceChecks = $false,
	
	# If used, runs redistribute even if automatic redistribute is disabled
	[switch] $SkipAutoRedistributeCheck = $false,
	
	# Optional string field to collect data on who calls the script
	[string] $RedistributeRequestor = "Operator"
)


Set-StrictMode -Version 2.0
Import-LocalizedData -BindingVariable RedistributeActiveDatabases_LocalizedStrings -FileName RedistributeActiveDatabases.strings.psd1
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

$ReplicationTypeType = [Microsoft.Exchange.Data.Directory.SystemConfiguration.ReplicationType]
$MoveStatusType = [Microsoft.Exchange.Management.SystemConfigurationTasks.MoveStatus]
$MountStatusType = [Microsoft.Exchange.Management.SystemConfigurationTasks.MountStatus]
$AutoActivationType = [Microsoft.Exchange.Data.Directory.SystemConfiguration.DatabaseCopyAutoActivationPolicyType]
$CopyStatusType = [Microsoft.Exchange.Management.SystemConfigurationTasks.CopyStatus]

#-------------------
# Script variables #
#-------------------

# This is the DAG we are trying to balance DBs in.
[Microsoft.Exchange.Data.Directory.SystemConfiguration.DatabaseAvailabilityGroup] $script:dag = $null
# The list of DBs we're trying to balance.
[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $script:databases = @()

$script:serverDbDistribution = @{} 	# Hashtable indexed by ServerName, value of HADatabaseLoadBalancing.ServerDbDistributionEntry
$script:currentServerDbDistribution = @{} # Hashtable indexed by ServerName, value of HADatabaseLoadBalancing.ServerDbDistributionEntry [this is used to keep track of the current distribution as DBs are being moved around]

$script:serverToActiveDatabaseTable = @{} # Hashtable of active databases on a server: indexed by Servername, value of Collection<MailboxDatabase>
$script:serverToDatabaseCopyTable = @{} # Hashtable of database copies on a server (active or passive): indexed by Servername, value of Collection<MailboxDatabase>
$script:serverNameToMailboxServerTable = @{} # Hashtable indexed by ServerName, value of MailboxServer (table of mailboxserver objects)
$script:serverNameToSiteTable = @{} # Hashtable indexed by ServerName, value of AD Site name

$script:siteDbDistribution = @{} 	# Hashtable indexed by Site name, value of HADatabaseLoadBalancing.SiteDbDistributionEntry
$script:currentSiteDbDistribution = @{} # Hashtable indexed by Site name, value of HADatabaseLoadBalancing.SiteDbDistributionEntry [this is used to keep track of the current distribution as DBs are being moved around]

$script:databaseToMoveStatusTable = @{} # Hashtable indexed by Database name, value of HADatabaseLoadBalancing.DatabaseMoveStatus
$script:databaseHashTable = @{} # Hashtable indexed by Database name, value of MailboxDatabase (table of MailboxDatabase objects)

$script:databaseCopiesTried = @{} # indexed by databaseName\serverName, value of $true (just a hashset)
$script:siteToActiveDatabasesTable = @{} # Hashtable indexed by SiteName, value of Collection<MailboxDatabase> (table of active DBs in each site)
$script:siteToServerNameTable = @{} # Hashtable indexed by SiteName, value of Collection<server name> (table of servers in a site)

$script:clusterNodeStateTable = @{} #Hashtable indexed by server name, with booleans for whether or not the server is online
[System.Diagnostics.Stopwatch] $script:clusterNodeStopwatch = New-Object -TypeName System.Diagnostics.Stopwatch
[System.Diagnostics.Stopwatch] $script:clusterNodeOverallStopwatch = New-Object -TypeName System.Diagnostics.Stopwatch
$script:clusterOutput = $null

[DateTime]$script:startTime = Get-Date # Time the script started running (used only for reporting)
[DateTime]$script:endTime = Get-Date # Time the script stopped running (used only for reporting)

[bool]$script:printSiteInformation = $false # Controls whether or not we display Site information

# This represents the report to be logged in the application event
[System.Text.StringBuilder] $script:eventReport = $null


# Build additional hashtables used in site balancing
function Populate-AdditionalSiteTables
{
	# build hashtable of Active DBs in a site
	$script:siteToActiveDatabasesTable.Clear()
	$script:siteToServerNameTable.Clear()
	
	# First validate that we actually have all the active databases indexed in $script:serverToActiveDatabaseTable
	[int]$countOfActiveDbs = ($script:serverToActiveDatabaseTable.Values | % { $_ } | Measure-Object).Count
	[int]$numDbs = $script:databases.Count
	if ($countOfActiveDbs -ne $numDbs)
	{
		Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0000 -f $countOfActiveDbs,$numDbs,"Populate-AdditionalSiteTables") -Stop
		return
	}
	
	if ( $script:serverToActiveDatabaseTable -and `
		($script:serverToActiveDatabaseTable.Count -gt 0))
	{
		Foreach ($kvp in $script:serverToActiveDatabaseTable.GetEnumerator())
		{
			[string]$tmpServerName = $kvp.Key
			[string]$siteName = $script:serverNameToSiteTable[$tmpServerName]
			
			if ( !$script:siteToActiveDatabasesTable.Contains($siteName) )
			{
				[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $tmpList = `
					[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] @()
				$script:siteToActiveDatabasesTable[$siteName] = $tmpList
			}
			$script:siteToActiveDatabasesTable[$siteName] += $kvp.Value
		}
	}
	
	if ( $script:serverNameToSiteTable -and `
		($script:serverNameToSiteTable.Count -gt 0))
	{
		Foreach ($kvp in $script:serverNameToSiteTable.GetEnumerator())
		{
			[string]$tmpServerName = $kvp.Key
			[string]$siteName = $script:serverNameToSiteTable[$tmpServerName]
			
			if ( !$script:siteToServerNameTable.Contains($siteName) )
			{
				[String[]] $tmpList = [String[]] @()
				$script:siteToServerNameTable[$siteName] = $tmpList
			}
			$script:siteToServerNameTable[$siteName] += $tmpServerName
		}
	}
}

function Is-DatabaseReplicated ([Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $mdb)
{
	if ($mdb.ReplicationType -eq $ReplicationTypeType::Remote)
	{
		return $true;
	}
	return $false;
}

filter Select-ReplicatedDatabases ( [bool]$selectNonReplicatedDatabases )
{
	if ( (Is-DatabaseReplicated $_) -or $selectNonReplicatedDatabases )
	{
		$_
	}
}

# Given a list of Databases, find the maximum possible number of copies of any DB there exists.
# This is also referred to as the highest possible Activation Preference.
# 	eg: mdb1 has 2 copies; mdb2 has 4 copies ====> result: 4
function Get-MaximumActivationPreference(
	[Parameter(Mandatory=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $dbs
)
{
	[int]$maxAP = ( $dbs | foreach { $_.Servers.Count } | Measure-Object -Maximum ).Maximum	
	
	Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0001 -f $maxAP,"Get-MaximumActivationPreference")
	return $maxAP
}

# Returns a list of MailboxDatabases in the given DAG (can even return non-replicated MDBs)
function Lookup-DatabasesFromDag (
	[Parameter(Mandatory=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.DatabaseAvailabilityGroup] $dag)
{
	Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0002 -f $dag,"Lookup-DatabasesFromDag")

	[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $mdbs = $null
	
	# Use -Status to get the MountedOnServer from ActiveManager, and also to get store mount state
	$sb = 
	{ 
		# database list will be sent to the pipeline
		$dag.Servers | foreach { Get-MailboxDatabase -Server $_ } | select -Unique | Select-ReplicatedDatabases $IncludeNonReplicatedDatabases | Get-MailboxDatabase -Status
	}

	$mdbs = Run-TimedScriptBlock $null $sb ($RedistributeActiveDatabases_LocalizedStrings.res_0093 -f "Lookup-DatabasesFromDag")
	
	return $mdbs
}

# Get the short name of the active server for the given DB
function Get-ActiveServerForDatabase (
	[Parameter(Mandatory=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $mdb)
{
	[string]$activeServer = $null
	
	if ($mdb.MountedOnServer)
	{
		# MountedOnServer is an FQDN, so convert to short name
		$activeServer = $mdb.MountedOnServer -replace "\..*$"
	}
	if ( !$activeServer )
	{
		$activeServer = $mdb.ServerName
	}
	
	return $activeServer
}

# Build the server and site hash tables
function Populate-ServerHashTables
{
	$sb = { Populate-ServerHashTablesInternal }
	Run-TimedScriptBlock $null $sb ($RedistributeActiveDatabases_LocalizedStrings.res_0094 -f "Populate-ServerHashTables")
}

# Populate the database hash table
function Populate-DatabasesTable (
	[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $databases)
{
	$script:databaseHashTable.Clear()
	
	if ($databases)
	{
		Foreach ($database in $databases)
		{
			$script:databaseHashTable[$database.Name] = $database
		}
	}
}

# Build the server and site hash tables
function Populate-ServerHashTablesInternal
{
	$script:serverNameToMailboxServerTable.Clear()
	$script:serverNameToSiteTable.Clear()
	
	($script:dag).Servers | foreach `
	{
		$serverName = $_.Name
		
		# Get the mailbox server
		$mbxServer = Get-MailboxServer $serverName
		if ($mbxServer)
		{
			$script:serverNameToMailboxServerTable[$serverName] = $mbxServer
		}
		else
		{
			Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0003 -f $serverName,"Populate-ServerHashTables")
		}
		
		# Get the exchange server to lookup the AD site
		$exServer = Get-ExchangeServer $serverName
		if ($exServer)
		{
			$script:serverNameToSiteTable[$serverName] = $exServer.Site.Name
		}
		else
		{
			Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0004 -f $serverName,"Populate-ServerHashTables")
		}
	}
}

# Build the hashtable describing the list of DBs on each server.
function Populate-ServerToDatabaseTable
{
	$script:serverToActiveDatabaseTable.Clear()
	$script:serverToDatabaseCopyTable.Clear()
	
	if ($script:databases)
	{
		# Initialize each to an empty list
		foreach ($server in $script:dag.Servers)
		{
			[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $tmpList = `
				[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] @()
			$script:serverToActiveDatabaseTable[$server.Name] = $tmpList
			
			[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $tmpList2 = `
				[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] @()
			$script:serverToDatabaseCopyTable[$server.Name] = $tmpList2
		}
		
		foreach ($mdb in $script:databases)
		{
			[string]$activeServer = Get-ActiveServerForDatabase $mdb
			
			# First, add this DB as an active to the appropriate server
			if ( !$script:serverToActiveDatabaseTable.Contains($activeServer) )
			{
				[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $tmpList = `
					[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] @()
				$script:serverToActiveDatabaseTable[$activeServer] = $tmpList
			}
			$script:serverToActiveDatabaseTable[$activeServer] += $mdb
			
			# Next, add this DB to all servers where it has copies
			$mdb.Servers | foreach `
			{
				$tmpServer = $_.Name
				
				if ( !$script:serverToDatabaseCopyTable.Contains($tmpServer) )
				{
					[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $tmpList = `
						[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] @()
					$script:serverToDatabaseCopyTable[$tmpServer] = $tmpList
				}
				$script:serverToDatabaseCopyTable[$tmpServer] += $mdb
			}
		}
	}
}

# Build the hashtable describing what DBs are active/passive/mounted/etc on each server.
function Populate-DatabaseDistributionMap
{
	$sb = { Populate-DatabaseDistributionMapInternal }
	Run-TimedScriptBlock $null $sb ($RedistributeActiveDatabases_LocalizedStrings.res_0095 -f "Populate-DatabaseDistributionMap")
}

function Populate-DatabaseDistributionMapInternal
{
	$script:serverDbDistribution.Clear()
	$script:siteDbDistribution.Clear()
	
	# First, lookup all DBs in this DAG
	$script:databases = Lookup-DatabasesFromDag $script:dag	

	# Populate the database hash table
	Populate-DatabasesTable $script:databases

	# Next, populate the tables describing the list of DBs on each server.
	Populate-ServerToDatabaseTable

	# build additional site-based hashtables
	Populate-AdditionalSiteTables

	# Determine the current maximum number of DB copies for any DB
	[int]$maxAP = Get-MaximumActivationPreference $script:databases
	
	# First, figure out the distribution of DBs by server
	($script:dag).Servers | foreach `
	{
		$serverName = $_.Name
		
		# initialize a new entry
		[HADatabaseLoadBalancing.ServerDbDistributionEntry]$entry = CreateEmptyServerDbRedundancyEntry $maxAP
		$entry.ServerName = $serverName
		$entry.DagName = $script:dag.Name
		
		# get the list of databases for the given server (this includes both Actives and Passives)
		$dbs = $script:serverToDatabaseCopyTable[$serverName]
		
		foreach ($mdb in $dbs)
		{
			if ( Is-DatabaseActiveOnServer $mdb $serverName )
			{
				Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0005 -f $mdb.Mounted,$mdb,$serverName,"Populate-DatabaseDistributionMap")
				
				$entry.ActiveDbs++
				if ($mdb.Mounted)
				{
					$entry.MountedDbs++
				}
				else
				{
					$entry.DismountedDbs++
				}
			}
			else
			{
				Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0006 -f $mdb,$serverName,"Populate-DatabaseDistributionMap")
				$entry.PassiveDbs++
			}
			
			# Record the Activation Preference of this DB on this server
			$activationPref = Get-ActivationPreferenceOfDatabaseCopy $mdb $serverName
			$entry.AddDatabaseCopyOfPreference($activationPref)
		}
		
		$script:serverDbDistribution[$serverName] = $entry
	}
	
	# Now, figure out the distribution of DBs by AD site	
	($script:dag).Servers | foreach `
	{
		$serverName = $_.Name
		$siteName = $script:serverNameToSiteTable[$serverName]
		
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0007 -f $serverName,$siteName,"Populate-DatabaseDistributionMap")
		
		if ( !$script:siteDbDistribution.Contains($siteName) )
		{
			# initialize a new entry
			[HADatabaseLoadBalancing.SiteDbDistributionEntry]$siteEntry = CreateEmptySiteDbRedundancyEntry $maxAP
			$siteEntry.SiteName = $siteName
			$siteEntry.DagName = $script:dag.Name
			$script:siteDbDistribution[$siteName] = $siteEntry
		}
		
		$siteEntry = $script:siteDbDistribution[$siteName]
		$serverEntry = $script:serverDbDistribution[$serverName]
		$siteEntry.AddServerDistribution($serverEntry)
	}
	
}


# Queries the DAG object for the cluster node states.  Populates a hashtable indexed by server name,
# with booleans for whether or not the node is online.
function Populate-ClusterNodeStatus
{
	[string]$dagName = $null
	
	# Find the DAG name first.
	$dagName = $script:dag.Name
	
	Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0013 -f $dagName,"Populate-ClusterNodeStatus")
	
	$script:clusterNodeOverallStopwatch.Reset();
	$script:clusterNodeOverallStopwatch.Start();
		
	$dag = $script:dag
	if ($dag)
	{
		if (!$dag.Servers -or `
			($dag.Servers.Count -eq 0))
		{
			Log-Verbose ($CheckDatabaseRedundancy_LocalizedStrings.res_0015 -f $dagName,"Populate-ClusterNodeStatus")	
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


# Create a copy of the $script:serverDbDistribution table.
function Copy-ServerDbDistributionTable
{
	# Copy the DB server distribution to track during moves
	$script:currentServerDbDistribution.Clear()
	$script:currentSiteDbDistribution.Clear()
	
	Foreach ($kvp in $script:serverDbDistribution.GetEnumerator())
	{
		$serverName = $kvp.Key
		$entry = $kvp.Value
		# Create a copy of the distribution entry object
		$script:currentServerDbDistribution[$serverName] = $entry.Clone()
	}
	
	Foreach ($kvp in $script:siteDbDistribution.GetEnumerator())
	{
		$siteName = $kvp.Key
		$entry = $kvp.Value
		# Create a copy of the distribution entry object
		$script:currentSiteDbDistribution[$siteName] = $entry.Clone()
	}
}

# Returns the ActivationPreference number for the given DB copy
function Get-ActivationPreferenceOfDatabaseCopy(
	[Parameter(Mandatory=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $mdb,
	[Parameter(Mandatory=$true)] [string] $serverName)
{
	# Record the Activation Preference of this DB on this server (should only be one match)
	$kvp = ( $mdb.ActivationPreference | where { $_.Key -ilike $serverName } | select -First 1 )
	[int]$activationPref = $kvp.Value
	return $activationPref
}

# Returns the server name for the DB copy at the specified activation preference
function Get-ServerAtActivationPreference(
	[Parameter(Mandatory=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $mdb,
	[Parameter(Mandatory=$true)] [int] $activationPreference)
{
	# Servers is always sorted by ActivationPreference
	return $mdb.Servers[$activationPreference - 1].Name
}

# Determine if the specified DB is active on the given server
function Is-DatabaseActiveOnServer(
	[Parameter(Mandatory=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $mdb,
	[Parameter(Mandatory=$true)] [string] $serverName)
{
	[string]$activeServer = Get-ActiveServerForDatabase $mdb
	if ($activeServer -ieq $serverName)
	{
		return $true
	}
	
	return $false
}

# Determines if the local server is the PAM
function Is-LocalServerPAM
{
	[bool]$isLocalPam = $false;
	Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0022 -f "Is-LocalServerPAM"  )
	
	# Get the PAM first
	$tmpDag = Get-DatabaseAvailabilityGroup $script:dag.Name -Status
	if ($tmpDag.PrimaryActiveManager -ilike $env:COMPUTERNAME)
	{
		$isLocalPam = $true	
	}
	
	Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0023 -f $isLocalPam,"Is-LocalServerPAM")
	return $isLocalPam
}

# This object represents a server's Database distribution. 
function CreateEmptyServerDbRedundancyEntry(
	[Parameter(Mandatory=$true)] [int] $maximumActivationPreference
)
{
	[HADatabaseLoadBalancing.ServerDbDistributionEntry]$entry = New-Object -TypeName "HADatabaseLoadBalancing.ServerDbDistributionEntry" -ArgumentList $maximumActivationPreference
	return $entry
}

# This object represents an AD Site's Database distribution. 
function CreateEmptySiteDbRedundancyEntry(
	[Parameter(Mandatory=$true)] [int] $maximumActivationPreference
)
{
	[HADatabaseLoadBalancing.SiteDbDistributionEntry]$entry = New-Object -TypeName "HADatabaseLoadBalancing.SiteDbDistributionEntry" -ArgumentList $maximumActivationPreference
	return $entry
}

# This object represents a DB move status (it might have not moved, moved, or failed to move)
function CreateEmptyDbMoveStatusObject
{
	[HADatabaseLoadBalancing.DatabaseMoveStatus]$moveStatus = New-Object -TypeName "HADatabaseLoadBalancing.DatabaseMoveStatus"
	return $moveStatus
}

function Show-DatabaseCurrentActives(
	[Parameter(Mandatory=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $dbs)
{
	$sb = { $dbs | Get-DatabaseCurrentActive }
	Run-TimedScriptBlock $null $sb "Show-DatabaseCurrentActives:"
}

function Get-DatabaseCurrentActive (
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $mdb)
{
	Begin
	{
		# Print a useful report header
		Print-ReportHeader
	}
	Process
	{
		[HADatabaseLoadBalancing.DatabaseMoveStatus]$moveStatus = CreateEmptyDbMoveStatusObject
		$moveStatus.DbName = $mdb.Name
		$script:databaseToMoveStatusTable[$mdb.Name] = $moveStatus
		
		[string]$activeServer = Get-ActiveServerForDatabase $mdb
		[int]$activePref = Get-ActivationPreferenceOfDatabaseCopy $mdb $activeServer
		$moveStatus.ActiveServerAtStart = $activeServer
		$moveStatus.ActiveOnPreferenceAtStart = $activePref
		$moveStatus.MoveStatus = [HADatabaseLoadBalancing.MoveStatus]::NoMoveAttempted
		
		Write-Output $moveStatus
	}
	End
	{
		Print-ReportFooter
	}
}

# Prints a header to the console
function Print-ReportHeader(
	[int] $allowedMaxDelta = 0
)
{
	[int]$copiesCount = ($script:serverDbDistribution.Values | Measure-Object -Sum -Property TotalDbs).Sum
	
	[string] $msg = "
***************************************
Balance DAG DBs
$((Get-Date).DateTime)
***************************************
Dag                                :   $($script:dag.Name)
ServerCount                        :   $($script:serverNameToMailboxServerTable.Count)
DatabaseCount                      :   $($script:databases.Count)
CopiesCount                        :   $copiesCount"

$script:eventReport.AppendLine($msg) | Out-Null
Log-Info $msg

	if ($PSCmdlet.ParameterSetName  -eq "BalanceDbsBySiteAndActivationPreference")
	{
		$msg = "AllowedDeviationFromMeanPercentage :   $AllowedDeviationFromMeanPercentage %
AllowedMaxDelta                    :   $allowedMaxDelta
"
		$script:eventReport.AppendLine($msg) | Out-Null
		Log-Info $msg
	}
	
	Print-ServerSiteDistribution -starting:$true

}

function Print-ServerSiteDistribution(
	[bool] $starting = $true
)
{	
	[string]$msg = $null
	[string]$startStopStr = "Starting"
	if (!$starting)
	{
		$startStopStr = "Ending"
	}
	
	if ($script:siteDbDistribution.Count -gt 1)
	{
		[string]$distribStr = $script:siteDbDistribution.Values | Sort-Object -Property ActiveDbs | Format-Table -AutoSize -Wrap | Out-String
		$distribStr = $distribStr -replace "\s+$" # trim the white space at the end
		
		$msg = "
--------------------------
$startStopStr Site Distribution
--------------------------
$distribStr
"
		$script:eventReport.AppendLine($msg) | Out-Null		
		Log-Info $msg
	}
	
	[string]$serverDistribStr = $script:serverDbDistribution.Values  | Sort-Object -Property ServerName | Format-Table -AutoSize -Wrap | Out-String
	$serverDistribStr = $serverDistribStr -replace "\s+$" # trim the white space at the end
	
	$msg = "
----------------------------
$startStopStr Server Distribution
----------------------------
$serverDistribStr
"
	
	$script:eventReport.AppendLine($msg) | Out-Null
	Log-Info $msg

}

# Tries to balance active DBs by AD site and activation preference
function Balance-DbsBySiteAndActivationPreference
{
	[int]$totalDatabasesCount = $script:databases.Count
	[int]$numSites = $script:siteDbDistribution.Count
	
	if ($numSites -lt 1)
	{
		Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0024 -f "Balance-DbsBySiteAndActivationPreference"  ) -Stop
		return
	}	
	
	if ($numSites -lt 2)
	{
		[string]$siteName = $script:siteDbDistribution.Keys | select -First 1
		# only one site (eg: Datacenter R4)
		Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0025 -f $siteName)
		Balance-DbsByPreferenceOptimized
		return
	}
	
	# We have at least 2 sites... party on!
	$script:printSiteInformation = $true
	
	# Calculate the allowed min-max delta (aka deviation)
	[int]$deltaMin = 2 	# Constant: this is the lowest possible value the allowed min-max delta can be!
	[double]$avgDbsPerSite = $totalDatabasesCount / $numSites
	[int]$allowedMaxDelta = [Math]::Ceiling( (2 * $avgDbsPerSite * ($AllowedDeviationFromMeanPercentage / 100)) )
	$allowedMaxDelta = [Math]::Max( $allowedMaxDelta, $deltaMin)
	
	# Sort the site distribution in ascending order
	$sortedSiteDistrib = $script:siteDbDistribution.Values | Sort-Object -Property ActiveDbs
	[int]$maxActives = ($sortedSiteDistrib | Select-Object -Last 1).ActiveDbs
	[int]$minActives = ($sortedSiteDistrib | Select-Object -First 1).ActiveDbs
	
	# Copy the DB server distribution to track during moves -
	# (this is the in-memory server-wise active copy distribution)
	Copy-ServerDbDistributionTable	
		
	# Print a useful report header
	Print-ReportHeader -allowedMaxDelta $allowedMaxDelta

	Log-Info "
-----------------------
Starting Database Moves
-----------------------
"

	[int]$currentMaxDelta = $maxActives - $minActives
	if ( $currentMaxDelta -le $allowedMaxDelta )
	{
		# The sites started off balanced within the allowed deviation
		Log-Info ($RedistributeActiveDatabases_LocalizedStrings.res_0026 -f $currentMaxDelta)
	}
	else
	{
		Log-Info ($RedistributeActiveDatabases_LocalizedStrings.res_0027 -f $currentMaxDelta)
	}
	
	# The real balancing logic is here
	Balance-DbsByOptimizedInternal	-sortServersBySite:$true -AllowedMaxSiteDelta:$allowedMaxDelta
	
	($curMinActives, $curMaxActives) = Get-CurrentMinMaxSiteActives
	$currentMaxDelta = $curMaxActives - $curMinActives
	if ($currentMaxDelta -gt $allowedMaxDelta)
	{
		Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0028 -f $currentMaxDelta,$allowedMaxDelta)
		# Now, try moving some DBs to less preferred copies
		Balance-DbsByOptimizedInternal	-sortServersBySite:$true -AllowedMaxSiteDelta:$allowedMaxDelta -MoveDbsToLessPreferred
	}
	
	($curMinActives, $curMaxActives) = Get-CurrentMinMaxSiteActives
	$currentMaxDelta = $curMaxActives - $curMinActives
	if ($currentMaxDelta -le $allowedMaxDelta)
	{
		Log-Success ($RedistributeActiveDatabases_LocalizedStrings.res_0102 -f $currentMaxDelta,$allowedMaxDelta)
	}
	else
	{
		Log-Info -Color:"Red" "Sites are still unbalanced! CurrentMaxDelta=$currentMaxDelta, AllowedMaxDelta=$allowedMaxDelta"
	}

	
	if ($ShowFinalDatabaseDistribution)
	{
		# Re-build the hashtable describing what DBs are active/passive/mounted/etc on each server.
		Populate-DatabaseDistributionMap
		Print-ServerSiteDistribution -starting:$false
	}
	
	Print-ReportFooter
	
	# write all the move status objects to the output pipeline
	$script:databaseToMoveStatusTable.Values | Sort-Object -Property DbName | Foreach { Write-Output $_ }
}

function Get-CurrentMinMaxSiteActives
{
	$curSortedSiteDistrib = $script:currentSiteDbDistribution.Values | Sort-Object -Property ActiveDbs
	[int]$curMaxActives = ($curSortedSiteDistrib | Select-Object -Last 1).ActiveDbs
	[int]$curMinActives = ($curSortedSiteDistrib | Select-Object -First 1).ActiveDbs
	
	return $curMinActives,$curMaxActives
}

function Balance-ActivationPreferences
{
	# Copy the DB server distribution to track during moves -
	# (this is the in-memory server-wise active copy distribution)
	Copy-ServerDbDistributionTable
	
	# Print a useful report header
	Print-ReportHeader
	
	Log-Info "
----------------------------------------
Starting Activation Preference Balancing
----------------------------------------
"
	# The real balancing logic is here
	Balance-ActivationPreferencesByGrouping
	
	if ($ShowFinalDatabaseDistribution)
	{
		# Re-build the hashtable describing what DBs are active/passive/mounted/etc on each server.
		Populate-DatabaseDistributionMap
		
		if ($WhatIf)
		{
			# In -WhatIf mode, we will show the hypothetical preference distribution
			($script:dag).Servers | foreach `
			{
				$serverName = $_.Name
				[HADatabaseLoadBalancing.ServerDbDistributionEntry]$entry = $script:serverDbDistribution[$serverName]
				
				# Overwrite the ActivationPreferences with the hypothetical values
				$entry.ClearPreferenceCountList()
				
				# get the list of databases for the given server (this includes both Actives and Passives)
				$dbs = $script:serverToDatabaseCopyTable[$serverName]
				
				foreach ($mdb in $dbs)
				{
					$ap = [HADatabaseLoadBalancing.DatabaseCopyPreferences]::GetActivationPreferenceForCopy($mdb.Name, $serverName)
					$entry.AddDatabaseCopyOfPreference($ap)
				}
			}
		}	

		Print-ServerSiteDistribution -starting:$false
	}
}

# Tries to balance DBs across servers w/o regard for ActivationPreference or site balance.
function Balance-DbsIgnoringActivationPreference
{
	# Copy the DB server distribution to track during moves -
	# (this is the in-memory server-wise active copy distribution)
	Copy-ServerDbDistributionTable	
	
	# Print a useful report header
	Print-ReportHeader
	
	Log-Info "
-----------------------
Starting Database Moves
-----------------------
"

	# The real balancing logic is here
	Balance-DbsByOptimizedInternal -IgnoreActivationPreference:$true -MoveDbsToLessPreferred:$true
	
	if ($ShowFinalDatabaseDistribution)
	{
		# Re-build the hashtable describing what DBs are active/passive/mounted/etc on each server.
		Populate-DatabaseDistributionMap
		Print-ServerSiteDistribution -starting:$false
	}
	
	Print-ReportFooter
	
	# write all the move status objects to the output pipeline
	$script:databaseToMoveStatusTable.Values | Sort-Object -Property DbName | Foreach { Write-Output $_ }
}

# Tries to move DBs to their most preferred copy w/o regard for AD site balance.
# This version tries to minimize the Active DB copy imbalance while moving the DBs around.
# That is, if you have a starting configuration of active copies: 4,4,4,4 (16 DBs on 4 servers, 
# with each DB having 4 copies, and none of them being on their most preffered copy), we will
# try to avoid going through 8,0,4,4 to get to the final balanced state.
function Balance-DbsByPreferenceOptimized
{
	# Copy the DB server distribution to track during moves -
	# (this is the in-memory server-wise active copy distribution)
	Copy-ServerDbDistributionTable	
	
	# Print a useful report header
	Print-ReportHeader
	
	Log-Info "
-----------------------
Starting Database Moves
-----------------------
"

	# The real balancing logic is here
	Balance-DbsByOptimizedInternal
	
	if ($ShowFinalDatabaseDistribution)
	{
		# Re-build the hashtable describing what DBs are active/passive/mounted/etc on each server.
		Populate-DatabaseDistributionMap
		Print-ServerSiteDistribution -starting:$false
	}
	
	Print-ReportFooter
	
	# write all the move status objects to the output pipeline
	$script:databaseToMoveStatusTable.Values | Sort-Object -Property DbName | Foreach { Write-Output $_ }
}


function Balance-DbsByOptimizedInternal(
	[bool]$sortServersBySite = $false,
	[int] $AllowedMaxSiteDelta = [int]::MaxValue,
	[switch] $MoveDbsToLessPreferred = $false,
	[switch] $IgnoreActivationPreference = $false)
{
	[bool]$keepRunning = $true
	[int]$iterations = 0
	$script:databaseCopiesTried.Clear() # indexed by databaseName\serverName, value of $true (just a hashset)
	
	while ($keepRunning)
	{
		$iterations++
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0029 -f $iterations,"Balance-DbsByPreferenceOptimized")
		
		[bool]$breakOutOfOuterLoop = $false;
		[int]$maxActives = 0
		[int]$minActives = 0
		[String[]] $sortedServers = $null
		$descSortedDistrib = $null # server distribution sorted in descending order
		
		if ($sortServersBySite)
		{
			# Sort the server distribution in descending order (grouped by AD sites)
			$descSortedSiteDistrib = $script:currentSiteDbDistribution.Values | Sort-Object -Descending -Property ActiveDbs
			$descSortedDistrib = $descSortedSiteDistrib `
				| where { $script:siteToServerNameTable[$_.SiteName].Count -gt 0 } `
				| foreach { `
					$script:siteToServerNameTable[$_.SiteName] `
					| foreach { $script:currentServerDbDistribution[$_] } `
					| Sort-Object -Descending -Property ActiveDbs `
				}
			
			$maxActives = ($descSortedDistrib | Select-Object -First 1).ActiveDbs
			$minActives = ($descSortedDistrib | Select-Object -Last 1).ActiveDbs
			
			# Sorted server list in ascending order (grouped by sites)
			$sortedServers = $script:currentSiteDbDistribution.Values | Sort-Object -Property ActiveDbs `
				| where { $script:siteToServerNameTable[$_.SiteName].Count -gt 0 } `
				| foreach { `
					$script:siteToServerNameTable[$_.SiteName] `
					| foreach { $script:currentServerDbDistribution[$_] } `
					| Sort-Object -Property ActiveDbs `
					| foreach { $_.ServerName } `
				} 
		}
		elseif ($IgnoreActivationPreference)
		{			
			# Sort the server distribution in descending order (don't factor in the AD site).
			# But, don't include the servers that are ActivationBlocked!
			$descSortedDistrib = $script:currentServerDbDistribution.Values `
									| Sort-Object -Descending -Property ActiveDbs `
									| Filter-ServersDistributionActivationBlocked
			$maxActives = ($descSortedDistrib | Select-Object -First 1).ActiveDbs
			$minActives = ($descSortedDistrib | Select-Object -Last 1).ActiveDbs
			
			# Sorted in ascending order
			$sortedServers = $descSortedDistrib `
								| Sort-Object -Property ActiveDbs `
								| Foreach { $_.ServerName } 
		}
		else # Common case when using -BalanceDbsByActivationPreference switch
		{
			# Sort the server distribution in descending order (don't factor in the AD site)
			$descSortedDistrib = $script:currentServerDbDistribution.Values | Sort-Object -Descending -Property ActiveDbs
			$maxActives = ($descSortedDistrib | Select-Object -First 1).ActiveDbs
			$minActives = ($descSortedDistrib | Select-Object -Last 1).ActiveDbs
			
			# Sorted in ascending order
			$sortedServers = $script:currentServerDbDistribution.Values | Sort-Object -Property ActiveDbs | Foreach { $_.ServerName }
		}
		
		$descServerNameList = @($descSortedDistrib | foreach { $_.ServerName })
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0030 -f $descServerNameList,"Balance-DbsByPreferenceOptimized")
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0031 -f $maxActives,$minActives,"Balance-DbsByPreferenceOptimized")
		
		# Defines the 'break' condition for the balancing loop.
		if ($IgnoreActivationPreference)
		{
			# The actives are balanced if the delta is 0 or 1
			if ( ($maxActives - $minActives) -lt 2 )
			{
				Log-Success ($RedistributeActiveDatabases_LocalizedStrings.res_0103 )
				# Stop balancing DBs 
				$keepRunning = $false
				break;
			}
		}
		elseif ($MoveDbsToLessPreferred)
		{
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0032 -f "Balance-DbsByPreferenceOptimized"  )
			
			($curMinActives, $curMaxActives) = Get-CurrentMinMaxSiteActives
			$currentMaxDelta = $curMaxActives - $curMinActives
			if ($currentMaxDelta -le $allowedMaxDelta)
			{
				# Stop balancing DBs once we are within the allowed site delta
				$keepRunning = $false
				break;
			}
		}
		
		Foreach ($serverWithMostDbs in $descServerNameList)
		{
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0033 -f $serverWithMostDbs,"Balance-DbsByPreferenceOptimized")
						
			# Find any active DBs that are yet to be moved 
			# Note: we shuffle the databases since our algorithm is somewhat sensitive to the order
			# in which we move DBs.
			$possDbsToMove = @($script:serverToActiveDatabaseTable[$serverWithMostDbs] | `
				Get-DatabasesToPossiblyMoveOff -serverName:$serverWithMostDbs `
					-MoveDbsToLessPreferred:$MoveDbsToLessPreferred | Shuffle-Objects)
			
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0034 -f $serverWithMostDbs,$possDbsToMove,"Balance-DbsByPreferenceOptimized")			
			if (!$possDbsToMove)
			{
				# There are no DBs to move off of this server, so lets move to the next server
				Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0035 -f $serverWithMostDbs,"Balance-DbsByPreferenceOptimized") 
				continue;
			}
			
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0036 -f $sortedServers,"Balance-DbsByPreferenceOptimized")
			Foreach ($possTargetServer in $sortedServers)
			{
				Foreach ($possDbToMove in $possDbsToMove)
				{
					[string]$dbCopyName = $possDbToMove.Name + "\" + $possTargetServer
					
					# does this DB have a copy on this server of a lower ActivationPreference? 
					$serversList = ($possDbToMove.Servers | foreach { $_.Name })
					if ($serversList -notcontains $possTargetServer)
					{
						Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0037 -f $possDbToMove,$possTargetServer,"Balance-DbsByPreferenceOptimized")
						continue;
					}
					
					# Have we already attempted a move for this DB copy?
					if ($script:databaseCopiesTried[$dbCopyName])
					{
						Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0038 -f $dbCopyName,"Balance-DbsByPreferenceOptimized")
						continue;
					}
					
					[int]$possAP = Get-ActivationPreferenceOfDatabaseCopy $possDbToMove $possTargetServer
					[int]$sourceAP = Get-ActivationPreferenceOfDatabaseCopy $possDbToMove $serverWithMostDbs
					
					# Is it OK to move the DB to a less preferred copy?
					if ($MoveDbsToLessPreferred)
					{
						# Skip the currently active copy
						if ($possAP -eq $sourceAP)
						{
							Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0039 -f $dbCopyName,$possAP,"Balance-DbsByPreferenceOptimized")
							continue;
						}
					}
					else
					{
						# We only want to move to a more preferred copy
						if ($possAP -ge $sourceAP)
						{
							Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0040 -f $dbCopyName,$possAP,$serverWithMostDbs,$sourceAP,"Balance-DbsByPreferenceOptimized")
							continue;
						}
					}
					
					# This check is not valid if we dont' care about the ActivationPreferences.
					if (!$IgnoreActivationPreference)
					{
						[bool]$skipThisCopy = $false
						# determine if this is the most preferred copy that has not yet been tried. Also, exclude
						# the currently active copy.
						foreach ($kvp in $possDbToMove.ActivationPreference)
						{
							# this list of APs is always sorted in increasing numeric value (i.e. first entry is always the most preferred)
							if ( ($kvp.Value -lt $possAP) -and ($kvp.Value -ne $sourceAP) )
							{
								[string]$tmpDbCopyName = $possDbToMove.Name + "\" + $kvp.Key.Name
								if (!$script:databaseCopiesTried[$tmpDbCopyName])
								{
									$skipThisCopy = $true
									break;
								}
							}
						}
						if ($skipThisCopy)
						{
							continue;
						}
					}
					
					# Record a move attempt for this DB to this target copy
					$script:databaseCopiesTried[$dbCopyName] = $true
					
					# Attempt the actual move here
					[bool]$moveSuccessful = $false
					if ($sortServersBySite)
					{
						$moveSuccessful = Move-DatabaseToOrderedList $possDbToMove @($possTargetServer) -MoveOnlyIfCurrentActiveTheSame -MoveWhatIf:$WhatIf -CheckSiteBalance -AllowedMaxSiteDelta:$AllowedMaxSiteDelta
					}
					else
					{
						$moveSuccessful = Move-DatabaseToOrderedList $possDbToMove @($possTargetServer) -MoveOnlyIfCurrentActiveTheSame -MoveWhatIf:$WhatIf
					}
					
					if ($moveSuccessful)
					{
						# start over since the active copy distribution has changed
						$breakOutOfOuterLoop = $true;
						break;
					}
				}
				
				if ($breakOutOfOuterLoop)
				{
					break;
				}
			}
			
			if ($breakOutOfOuterLoop)
			{
				break;
			}
		}
		
		# If we got through all the servers and didn't move any DBs, it means we're done.
		if (!$breakOutOfOuterLoop)
		{
			$keepRunning = $false;
			break;
		}
	}	
	
	# Some DBs might not have been attempted for moves since they were already on their most preferred copies...
	# For these, let's just initialize "no-op" move status entries for reporting purposes.
	foreach ($mdb in $script:databases)
	{
		if (!$script:databaseToMoveStatusTable.Contains($mdb.Name))
		{
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0041 -f $mdb,"Balance-DbsByPreferenceOptimized")
			[string]$currActive = Get-ActiveServerForDatabase $mdb
			[int]$ap = Get-ActivationPreferenceOfDatabaseCopy $mdb $currActive
			
			# When IgnoreActivationPreference is used, we can skip ActivationBlocked servers altogether, so this is a valid case.
			if ($ap -ne 1 -and !$IgnoreActivationPreference)
			{
				Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0042 -f $mdb)
			}
			
			# Create the DatabaseMoveStatus object to report the status of the move
			[HADatabaseLoadBalancing.DatabaseMoveStatus]$moveStatus = CreateEmptyDbMoveStatusObject
			$moveStatus.DbName = $mdb.Name
			$script:databaseToMoveStatusTable[$mdb.Name] = $moveStatus		
			$moveStatus.MoveStatus = [HADatabaseLoadBalancing.MoveStatus]::NoMoveAttempted
			$moveStatus.ActiveServerAtStart = $currActive
			$moveStatus.ActiveOnPreferenceAtStart = $ap
			$moveStatus.ActiveServerAtEnd = $currActive
			$moveStatus.ActiveOnPreferenceAtEnd = $ap
		}
	}
}

# Filters an incoming MailboxDatabase pipeline for only those DBs that we have yet to attempt a move for, 
# off of the server specified.
# NOTE: $_ is a MailboxDatabase object
filter Get-DatabasesToPossiblyMoveOff(
	[string]$serverName,
	[switch] $MoveDbsToLessPreferred = $false)
{
	# check if the DB has already been moved here.
	if ( $script:databaseToMoveStatusTable.Contains($_.Name) -and `
		($script:databaseToMoveStatusTable[$_.Name].MoveStatus -eq [HADatabaseLoadBalancing.MoveStatus]::MoveSucceeded))
	{
		# skip this DB
	}
	elseif ($_.ReplicationType -ne $ReplicationTypeType::Remote)
	{
		# skip this DB
	}
	else
	{		
		[int]$ap = Get-ActivationPreferenceOfDatabaseCopy $_ $serverName;
		
		if ( !$MoveDbsToLessPreferred -and ($ap -eq 1) )
		{
			# skip this DB
		}
		else
		{
			# figure out what are the possible target copies for this DB
			[String[]]$dbTargetServers = [String[]] @()
			
			if ($MoveDbsToLessPreferred)
			{
				# Any of the copies except for the current active, are valid targets
				$dbTargetServers = @( $_.ActivationPreference | where { $_.Value -ne $ap } | foreach { $_.Key.Name } )
			}
			else
			{
				# Only more preferred copies are valid targets
				$dbTargetServers = @( $_.ActivationPreference | Get-ServersOfLowerActivationPreference -activationPreference $ap )
			}
			
			[bool]$allPossibleTargetsTried = $true;
			Foreach ($dbTargetServer in $dbTargetServers)
			{
				[string]$possCopyName = $_.Name + "\" + $dbTargetServer;
				if (!$script:databaseCopiesTried[$possCopyName])
				{
					$allPossibleTargetsTried = $false;
					break;
				}
			}
			if ($allPossibleTargetsTried)
			{
				# skip this DB
			}
			else
			{
				# Let's check to see if the active copy of this DB is currently stuck in mounting/dismounting
				# We will skip these DBs since the move cmdlet will anyway be stuck behind the DB action already in progress.
				$activeStatus = Get-MailboxDatabaseCopyStatus $_ -Active | select -first 1
				if ($activeStatus -and (($activeStatus.Status -eq $CopyStatusType::Mounting) -or ($activeStatus.Status -eq $CopyStatusType::Dismounting)))
				{
					# skip this DB and record an "attempt" made for this DB so that we don't do this get-mdbcs check again
					
					Foreach ($dbTargetServer in $_.Servers)
					{
						# Record a move attempt for this DB to this target copy
						[string]$possCopyName = $_.Name + "\" + $dbTargetServer.Name;
						$script:databaseCopiesTried[$possCopyName] = $true
					}
					
					$msg = $RedistributeActiveDatabases_LocalizedStrings.res_0108 -f $_.Name,$activeStatus.ActiveDatabaseCopy,$activeStatus.Status.ToString()
					Log-Warning $msg
					# skip this DB
				}
				else
				{				
					# Since not all copies of this DB have been tried yet, let's return it.
					$_
				}
			}
		}
	}
}


# Tries to move DBs to their most preferred copy w/o regard for AD site balance
function Balance-DbsByPreference
{
	$sb = 
	{
		# Sort the server distribution in descending order
		$descSortedDistrib = $script:serverDbDistribution.Values | Sort-Object -Descending -Property ActiveDbs
		[int]$maxActives = ($descSortedDistrib | Select-Object -First 1).ActiveDbs
		[int]$minActives = ($descSortedDistrib | Select-Object -Last 1).ActiveDbs
		$databasesList = @()
		
		# One server is more overloaded
		if ( ($maxActives - $minActives) -gt 1 )
		{
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0043 -f $maxActives,$minActives,"Balance-DbsByPreference")
			
			# We should move off of servers with most active databases first.
			# Additionally, we're just going through the active databases.
			$databasesList = $descSortedDistrib `
				| where { $_.ActiveDbs -gt 0 } `
				| foreach { $script:serverToActiveDatabaseTable[$_.ServerName] | Shuffle-Objects }
		}
		else # servers are mostly equally balanced, so just move DBs randomly 
		{
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0044 -f $maxActives,$minActives,"Balance-DbsByPreference")
			$databasesList = $script:databases | Shuffle-Objects
		}
		
		# balance the DBs in the listed order
		$databasesList | Balance-DatabaseByPreference | Sort-Object -Property DbName
	}
	
	Run-TimedScriptBlock $null $sb ($RedistributeActiveDatabases_LocalizedStrings.res_0096 -f "Balance-DbsByPreference")
}

function Balance-DatabaseByPreference (
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $mdb)
{
	Begin
	{
		# Copy the DB server distribution to track during moves
		Copy-ServerDbDistributionTable	
		
		# Print a useful report header
		Print-ReportHeader
		
		Log-Info "
-----------------------
Starting Database Moves
-----------------------
"
	}
	Process
	{
		[string]$activeServer = Get-ActiveServerForDatabase $mdb
		[int]$activePref = Get-ActivationPreferenceOfDatabaseCopy $mdb $activeServer
		
		[String[]]$targetServers = $mdb.ActivationPreference | `
								   Get-ServersOfLowerActivationPreference -activationPreference $activePref
								   
		[bool]$moveSuccessful = Move-DatabaseToOrderedList $mdb $targetServers -MoveOnlyIfCurrentActiveTheSame -MoveWhatIf:$WhatIf
		
		# output the move status to the pipeline
		Write-Output $script:databaseToMoveStatusTable[$mdb.Name]
	}
	End
	{
		if ($ShowFinalDatabaseDistribution)
		{
			# Re-build the hashtable describing what DBs are active/passive/mounted/etc on each server.
			Populate-DatabaseDistributionMap
			Print-ServerSiteDistribution -starting:$false
		}
		
		Print-ReportFooter
	}
}

function Balance-ActivationPreferencesByGrouping
{
	# clear the preference assignments
	[HADatabaseLoadBalancing.DatabaseCopyPreferences]::Clear()
	
	# need to group databases together by the servers they have copies on
	
	# build a hashtable, indexed by copies list (i.e. server list) and value of MailboxDatabase collection
	$script:groupedDbs = @{}
	$script:rotatedGroupedDbs = @{}
	foreach ($db in ($script:databaseHashTable.Values | Sort-Object -Property Name))
	{
		[String[]]$dbServers = @($db.Servers | foreach { $_.Name } | Sort-Object)
		[string]$dbServersStr = [String]::Join("*", $dbServers)
		
		if (!$script:groupedDbs.Contains($dbServersStr))
		{
			[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $tmpList = `
					[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] @()
			$script:groupedDbs[$dbServersStr] = $tmpList;
		}
		$script:groupedDbs[$dbServersStr] += $db
	}
	
	# E15# 29672: Now rotate each group. If we don't do this, we may have a less-balanced preference distribution. Consider:
	# 
	# The logic is to group databases by servers they have copies on. So, for instance, with striping, we can have groups like: 
	# "server1,server2,server3" and "server1,server2,server4". The grouping logic will assign more databases on server 1 & server 2 
	# with the same preferences, which can make it less balanced overall. Instead, we will do something simple like rotate the 
	# servers from one group to the next. Meaning: in this example, the two groups could be something like: 
	# "server1,server2,server3" and "server4,server1,server2".
	[int]$rotateBy = 0
	foreach ($grouping in ($script:groupedDbs.Keys | Sort-Object))
	{
		$rotatedGroupName = Rotate-Group $grouping $rotateBy
		$rotateBy++
		$script:rotatedGroupedDbs[$rotatedGroupName] = $script:groupedDbs[$grouping]
	}
	
	# Now assign preferences group by group
	foreach ($grouping in ($script:rotatedGroupedDbs.Keys | Sort-Object))
	{
		[String[]]$serverNames = $grouping.Split("*")
		[int]$numCopies = $serverNames.Length
		$databases = @($script:rotatedGroupedDbs[$grouping] | Sort-Object -Property Name)
		
		for ($i = 0; $i -lt $databases.Length; $i++)
		{
			# start assigning activation preferences from an offset of the database we're on
			[int]$ap = ($i % $numCopies) + 1
			$db = $databases[$i]
			
			foreach ($server in $serverNames)
			{
				Assign-ActivationPreferenceForDatabaseCopy $server $db.Name $ap
				$ap = ($ap % $numCopies) + 1
			}
		}
	}
	
	# Check that every DatabaseCopy has been assigned an AP
	$script:serverToDatabaseCopyTable.Keys | foreach { `
		$serverName = $_
		$dbCopiesNotAssigned = $script:serverToDatabaseCopyTable[$serverName] | `
			where { ![HADatabaseLoadBalancing.DatabaseCopyPreferences]::IsAnyActivationPreferenceAssignedForCopy($_.Name, $serverName) }
		if ($dbCopiesNotAssigned) 
		{
			$dbCopiesNotAssigned | foreach { `
				$dbName = $_.Name
				$msg = $RedistributeActiveDatabases_LocalizedStrings.res_0107 -f $dbName,$serverName
				Log-Warning $msg
			}
		}
	}
	
	# Set the actual desired activation preferences in the ActiveDirectory
	foreach ($pair in [HADatabaseLoadBalancing.DatabaseCopyPreferences]::GetAllSortedDatabaseCopyPreferenceAssignments())
	{
		($dbName, $serverName) = $pair.Key.Split("\");
		$db = Get-MailboxDatabase $dbName
		[int]$prevAP = Get-ActivationPreferenceOfDatabaseCopy $db $serverName
		if ($prevAP -ne $pair.Value)
		{
			# Do the actual assignment of the activation preference
			Set-MailboxDatabaseCopy $dbName\$serverName -ActivationPreference:$pair.Value -WhatIf:$WhatIf -Confirm:$Confirm
		}
	}
}

function Assign-ActivationPreferenceForDatabaseCopy([string]$serverName, [string]$dbName, [int]$activationPref)
{
	Log-Info "Database copy '$dbName\$serverName': Activation Preference = $activationPref" Cyan
	
	[HADatabaseLoadBalancing.DatabaseCopyPreferences]::AssignPreference($dbName, $serverName, $activationPref)
}

# Returns a rotated group
function Rotate-Group([string]$grouping, [int]$count)
{
	[String[]]$rotatedArr = @()
	[string[]]$arr = $grouping.Split("*")
	[int]$total = $arr.Length
	
	for ($i = 0; $i -lt $total; $i++)
	{
		$rotatedArr += $arr[($i + $count) % $total]
	}
	
	[string]$rotatedGroup = [String]::Join("*", $rotatedArr)
	return $rotatedGroup
}

function Print-ReportFooter
{
	$groups = $script:databaseToMoveStatusTable.Values | Group-Object -AsHashTable -Property MoveStatus
	
	[int]$successfulMoves = 0
	if ($groups[[HADatabaseLoadBalancing.MoveStatus]::MoveSucceeded])
	{
		$successfulMoves = $groups[[HADatabaseLoadBalancing.MoveStatus]::MoveSucceeded].Count
	}
	
	[int]$failedMoves = 0
	if ($groups[[HADatabaseLoadBalancing.MoveStatus]::MoveFailed])
	{
		$failedMoves = $groups[[HADatabaseLoadBalancing.MoveStatus]::MoveFailed].Count
	}
	
	[int]$notMoved = 0
	if ($groups[[HADatabaseLoadBalancing.MoveStatus]::NoMoveAttempted])
	{
		$notMoved = $groups[[HADatabaseLoadBalancing.MoveStatus]::NoMoveAttempted].Count
	}
	
	[int]$movedToLessPreferred = 0
	if ($successfulMoves -gt 0)
	{
		$movedToLessPreferred = ($script:databaseToMoveStatusTable.Values `
			| where { $_.ActiveOnPreferenceAtStart -lt $_.ActiveOnPreferenceAtEnd } `
			| Measure-Object).Count
	}
	
	$script:endTime = Get-Date
	[TimeSpan]$duration = $script:endTime - $script:startTime
	
	[string]$msg = "
----------------
Summary of Moves
----------------
Successfully moved      : $successfulMoves
Moved to less preferred : $movedToLessPreferred
Failed to move          : $failedMoves
Not moved               : $notMoved

Start time              : $($script:startTime.DateTime)
End time                : $($script:endTime.DateTime)
Duration                : $($duration.ToString())
"

	$script:eventReport.AppendLine($msg) | Out-Null
	Log-Info $msg
}

function Get-ServersOfLowerActivationPreference (
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)] $activationPrefKvp,
	[Parameter(Mandatory=$true)] [int] $activationPreference)
{
	Process
	{
		if ($activationPrefKvp.Value -lt $activationPreference)
		{
			# output to pipeline
			Write-Output $activationPrefKvp.Key.Name
		}
	}
}

# FOR TEST: Randomly moves DBs around to create an uneven distribution
function Shuffle-ActiveDatabases (
	[Parameter(Mandatory=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] $dbs)
{
	$sb = {	$dbs | Shuffle-ActiveDatabase }
	Run-TimedScriptBlock $null $sb ( $RedistributeActiveDatabases_LocalizedStrings.res_0097 -f "Shuffle-ActiveDatabases")
}

# FOR TEST: Randomly move the given DB to one of its copies (we may even *NOT* move the DB)
function Shuffle-ActiveDatabase (
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $mdb)
{
	Begin
	{
		# Copy the DB server distribution to track during moves
		Copy-ServerDbDistributionTable	
		
		# Print a useful report header
		Print-ReportHeader
		
		Log-Info "
-----------------------
Starting Database Moves
-----------------------
"
	}
	Process
	{
		# $mdb.Servers is ordered by AP, so lets randomize the server list
		$randomServers = $mdb.Servers | foreach { $_.Name } | Shuffle-Objects
		
		# The returned boolean will be placed on the output pipeline
		[bool]$moveSuccessful = Move-DatabaseToOrderedList $mdb $randomServers -MoveWhatIf:$WhatIf
		
		# output the move status to the pipeline
		Write-Output $script:databaseToMoveStatusTable[$mdb.Name]
	}
	End
	{
		if ($ShowFinalDatabaseDistribution)
		{
			# Re-build the hashtable describing what DBs are active/passive/mounted/etc on each server.
			Populate-DatabaseDistributionMap
			Print-ServerSiteDistribution -starting:$false
		}
		
		Print-ReportFooter
	}
}

function Print-CurrentServerDbDistribution
{
	[string]$serverDistribStr = $script:currentServerDbDistribution.Values `
								| Sort-Object -Property ServerName `
								| Format-Table ServerName,ActiveDbs,PassiveDbs -AutoSize -Wrap `
								| Out-String
	$serverDistribStr = $serverDistribStr -replace "\s+$" # trim the white space at the end
	Log-Info ($RedistributeActiveDatabases_LocalizedStrings.res_0045 -f $serverDistribStr)
}

function Print-CurrentSiteDbDistribution
{
	[string]$siteDistribStr = $script:currentSiteDbDistribution.Values `
							  | Sort-Object -Property SiteName `
							  | Format-Table SiteName,ActiveDbs,PassiveDbs -AutoSize -Wrap `
							  | Out-String
	
	$siteDistribStr = $siteDistribStr -replace "\s+$" # trim the white space at the end
	Log-Info ($RedistributeActiveDatabases_LocalizedStrings.res_0045 -f $siteDistribStr)						  
}

# Attempts to move the given DB to a list of servers, in the order specified.
# If the current active server is in the list and we try to "move" to it, this method is considered successful.
# If the DB cannot be moved to any of the servers in the list, $false is returned.
# If the move is successful to one of the servers in the list, $true is returned.
function Move-DatabaseToOrderedList(
	[Parameter(Mandatory=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $mdb,
	[string[]] $serverList = $null,
	[switch] $MoveOnlyIfCurrentActiveTheSame = $false,
	[switch] $CheckSiteBalance = $false,
	[int] $AllowedMaxSiteDelta = [int]::MaxValue,
	[bool] $MoveWhatIf = $false)
{
	# Check if we are only supposed to run on the PAM. This covers the case where the PAM has failed
	# over while this script was running.
	if (!(CheckPAM))
	{
		# exit from the script
		Exit
	}
	
	# The returned boolean will be placed on the output pipeline
	$sb = { Move-DatabaseToOrderedListInternal $mdb $serverList `
		-MoveOnlyIfCurrentActiveTheSame:$MoveOnlyIfCurrentActiveTheSame `
		-CheckSiteBalance:$CheckSiteBalance `
		-AllowedMaxSiteDelta:$AllowedMaxSiteDelta `
		-MoveWhatIf:$MoveWhatIf }	
	[bool]$moveSuccessful = Run-TimedScriptBlock $null $sb "Move-DatabaseToOrderedList: Possibly moving DB '$mdb'"
	
	if (!$moveSuccessful)
	{
		# It is important to not error on move errors, because we should not use move failures as the criteria
		# for script failure or not. Instead we validate distribution the end of the script execution.
		Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0046 -f $mdb)
	}
	else
	{
		$moveStatus = $script:databaseToMoveStatusTable[$mdb.Name]
		if ($moveStatus.MoveStatus -eq [HADatabaseLoadBalancing.MoveStatus]::MoveSucceeded)
		{
			# This means the DB was actually moved successfully -- adjust the counts
			$activeEntry = $script:currentServerDbDistribution[$moveStatus.ActiveServerAtEnd]
			$activeEntry.ActiveDbs++
			$activeEntry.PassiveDbs--
			
			# Also adjust the site counts
			[string]$endSiteName = $script:serverNameToSiteTable[$moveStatus.ActiveServerAtEnd]
			$activeSiteEntry = $script:currentSiteDbDistribution[$endSiteName]
			$activeSiteEntry.ActiveDbs++
			$activeSiteEntry.PassiveDbs--
			
			$prevActiveEntry = $script:currentServerDbDistribution[$moveStatus.ActiveServerAtStart]
			$prevActiveEntry.ActiveDbs--
			$prevActiveEntry.PassiveDbs++
			
			[string]$startSiteName = $script:serverNameToSiteTable[$moveStatus.ActiveServerAtStart]
			$prevActiveSiteEntry = $script:currentSiteDbDistribution[$startSiteName]
			$prevActiveSiteEntry.ActiveDbs--
			$prevActiveSiteEntry.PassiveDbs++
			
			# Also remove the DB from the active table for this server
			# NOTE: This block is so ugly probably because of a PS bug. In Win2k8 R2 we may be able to simplify this...
			$originalActiveList = @($script:serverToActiveDatabaseTable[$moveStatus.ActiveServerAtStart])
			$script:serverToActiveDatabaseTable[$moveStatus.ActiveServerAtStart] = `
				@( $originalActiveList | where { $_.Name -ne $mdb.Name } )
			$newActiveList = @($script:serverToActiveDatabaseTable[$moveStatus.ActiveServerAtEnd])
			if (!$newActiveList)
			{
				# empty list
				$newActiveList = [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] @()
			}
			$newActiveList += $mdb
			$script:serverToActiveDatabaseTable[$moveStatus.ActiveServerAtEnd] = $newActiveList
			
			
			# Also remove this DB from the active table for this site
			# NOTE: This block is so ugly probably because of a PS bug. In Win2k8 R2 we may be able to simplify this...
			$originalActiveSiteList = @($script:siteToActiveDatabasesTable[$startSiteName])
			$script:siteToActiveDatabasesTable[$startSiteName] = `
				@( $originalActiveSiteList | where { $_.Name -ne $mdb.Name } )
			$newActiveSiteList = @($script:siteToActiveDatabasesTable[$endSiteName])
			if (!$newActiveSiteList)
			{
				# empty list
				$newActiveSiteList = [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase[]] @()
			}
			$newActiveSiteList += $mdb
			$script:siteToActiveDatabasesTable[$endSiteName] = $newActiveSiteList
			
			Log-Success ($RedistributeActiveDatabases_LocalizedStrings.res_0104 -f $moveStatus.ActiveServerAtStart,$moveStatus.ActiveServerAtEnd,$mdb)
			if ($startSiteName -ne $endSiteName)
			{
				Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0047 -f $mdb,$startSiteName,$endSiteName)
			}
			
			Print-CurrentServerDbDistribution			
			
			if ($script:printSiteInformation)
			{
				Print-CurrentSiteDbDistribution
			}
		}
	}
	
	return $moveSuccessful
}
	
# Attempts to move the given DB to a list of servers, in the order specified.
# If the current active server is in the list and we try to "move" to it, this method is considered successful.
# If the DB cannot be moved to any of the servers in the list, $false is returned.
# If the move is successful to one of the servers in the list, $true is returned.
function Move-DatabaseToOrderedListInternal(
	[Parameter(Mandatory=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $mdb,
	[string[]] $serverList = $null,
	[switch] $MoveOnlyIfCurrentActiveTheSame = $false,
	[switch] $CheckSiteBalance = $false,
	[int] $AllowedMaxSiteDelta = [int]::MaxValue,
	[bool] $MoveWhatIf = $false)
{
	[string]$serverListString = [string]::Empty
	if ($serverList -ne $null)
	{
		$serverListString = [string]::Join(",", $serverList)
	}
	Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0048 -f $mdb,$serverListString,"Move-DatabaseToOrderedList")
	
	[bool]$moveSuccessful = $false
	[bool]$firstResultHasBeenSet = $false
	[Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseMoveResult] $firstMoveResult = $null
	[Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseMoveResult[]] $moveResults = @()
	[Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseMoveResult] $lastMoveResult = $null
	[string]$originalActive = Get-ActiveServerForDatabase $mdb
	
	# Create the DatabaseMoveStatus object to report the status of the move
	[HADatabaseLoadBalancing.DatabaseMoveStatus]$moveStatus = CreateEmptyDbMoveStatusObject
	$moveStatus.DbName = $mdb.Name
	$script:databaseToMoveStatusTable[$mdb.Name] = $moveStatus
	# Initialize the move status to NoMoveAttempted
	$moveStatus.MoveStatus = [HADatabaseLoadBalancing.MoveStatus]::NoMoveAttempted
	$moveStatus.ActiveServerAtStart = $originalActive
	$moveStatus.ActiveOnPreferenceAtStart = Get-ActivationPreferenceOfDatabaseCopy $mdb $originalActive
	
	if ($serverList)
	{
		if ($MoveOnlyIfCurrentActiveTheSame)
		{
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0049 -f $mdb,$originalActive,"Move-DatabaseToOrderedList")
			$currentMdb = Get-MailboxDatabase $mdb -Status
			[string]$currentActive = Get-ActiveServerForDatabase $currentMdb
			
			if ($originalActive -ne $currentActive)
			{			
				# The DB has moved or failed over while this script was running! Let's fail this move.
				# Update the DB move status
				$moveStatus.ActiveServerAtEnd = $currentActive
				$moveStatus.ActiveOnPreferenceAtEnd = Get-ActivationPreferenceOfDatabaseCopy $currentMdb $currentActive
				
				Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0050 -f $mdb,$originalActive,$currentActive,"Move-DatabaseToOrderedList")
				Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0050 -f $mdb,$originalActive,$currentActive,"Move-DatabaseToOrderedList")
				
				Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0051 -f $moveSuccessful,"Move-DatabaseToOrderedList")
				return $moveSuccessful
			}
		}
	}
	else
	{
		# This is a success case since we don't actually have to move the DB
		$moveSuccessful = $true
		$moveStatus.MoveStatus = [HADatabaseLoadBalancing.MoveStatus]::NoMoveAttempted
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0052 -f $mdb,"Move-DatabaseToOrderedList")
		
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0051 -f $moveSuccessful,"Move-DatabaseToOrderedList")
		return $moveSuccessful
	}
	
	try
	{
		if ($serverList)
		{
			foreach ($server in $serverList)
			{
				if (Is-DatabaseActiveOnServer $mdb $server)
				{
					# This is a success case since we don't actually have to move the DB
					$moveSuccessful = $true
					$moveStatus.MoveStatus = [HADatabaseLoadBalancing.MoveStatus]::NoMoveAttempted
					Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0053 -f $mdb,$server,"Move-DatabaseToOrderedList")
					break;
				}
				
				[int]$sourceAP = Get-ActivationPreferenceOfDatabaseCopy $mdb $originalActive
				[int]$targetAP = Get-ActivationPreferenceOfDatabaseCopy $mdb $server
				Log-Info ($RedistributeActiveDatabases_LocalizedStrings.res_0054 -f $mdb,$originalActive,$sourceAP,$server,$targetAP)
				
				if ($CheckSiteBalance)
				{
					[string]$targetSiteName = $script:serverNameToSiteTable[$server]
					[string]$sourceSiteName = $script:serverNameToSiteTable[$originalActive]
					
					if ($targetSiteName -ne $sourceSiteName)
					{					
						$sortedSiteDistrib = $script:currentSiteDbDistribution.Values | Sort-Object -Property ActiveDbs
						$curMax = $sortedSiteDistrib | Select-Object -Last 1
						$curMin = $sortedSiteDistrib | Select-Object -First 1
						[int]$possMaxActives = $curMax.ActiveDbs
						[int]$possMinActives = $curMin.ActiveDbs
						
						if ($sourceSiteName -eq $curMax.SiteName)
						{
							$possMaxActives--		
						}
						elseif ($sourceSiteName -eq $curMin.SiteName)
						{
							$possMinActives--
						}
						
						if ($targetSiteName -eq $curMax.SiteName)
						{
							$possMaxActives++
						}
						elseif ($targetSiteName -eq $curMin.SiteName)
						{
							$possMinActives++
						}
						
						[int]$currentDelta = $curMax.ActiveDbs - $curMin.ActiveDbs
						[int]$possibleDelta = $possMaxActives - $possMinActives
						
						# If we are already starting out unbalanced, we should allow the move only if
						# the imbalance is not increased!
						if ($currentDelta -gt $AllowedMaxSiteDelta)
						{
							if ($possibleDelta -gt $currentDelta)
							{
								Log-Info ($RedistributeActiveDatabases_LocalizedStrings.res_0055 -f $curMax.ActiveDbs,$curMin.ActiveDbs,$mdb,$server,$AllowedMaxSiteDelta)
								continue;
							}
						}
						# We are starting out balanced. Make sure we don't exceed the AllowedMaxSiteDelta
						else
						{
							if ($possibleDelta -gt $AllowedMaxSiteDelta)
							{
								Log-Info ($RedistributeActiveDatabases_LocalizedStrings.res_0056 -f $curMax.ActiveDbs,$curMin.ActiveDbs,$mdb,$server,$AllowedMaxSiteDelta)		
								continue;
							}
						}
					}					
				}
				
				($serverPassed, $message) = Test-ServerForMove $originalActive $server
				if (!$serverPassed)
				{
					# Log-Verbose instead of Log-Error so that we can try another copy to move to.
					Log-Info -color:"Red" ($RedistributeActiveDatabases_LocalizedStrings.res_0101 -f $mdb,$server,$message)
					continue;
				}
				
				Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0057 -f $mdb,$server,"Move-DatabaseToOrderedList")
				
				if ($MoveWhatIf)
				{
					Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0058 -f $mdb,$server,"Move-DatabaseToOrderedList")
					$moveSuccessful = $true
					$moveStatus.MoveStatus = [HADatabaseLoadBalancing.MoveStatus]::MoveSucceeded
					# populate these fields here as a special-case for -WhatIf
					$moveStatus.ActiveServerAtEnd = $server
					$moveStatus.ActiveOnPreferenceAtEnd = Get-ActivationPreferenceOfDatabaseCopy $mdb $server
					break;
				}
				
				# Now that we have a move candidate, let's reset the move status to Failed
				$moveStatus.MoveStatus = [HADatabaseLoadBalancing.MoveStatus]::MoveFailed
				
				[string]$moveComment = $RedistributeActiveDatabases_LocalizedStrings.res_0100 -f "RedistributeActiveDatabases.ps1", $RedistributeRequestor
				# We use -ErrorAction:SilentlyContinue instead of -ErrorAction:Continue, since we are later doing Validate-Balance check,
				# so lets make sure the workflow operation is marked as failure or success
				# only based on the result of this check and not based on any previous move previously thrown.
				# Unfortunately this means move error will not be logged, but $lastMoveResult will still contain the information about failed move and will 
				# be added to the workflow operation logs.
				$lastMoveResult = Move-ActiveMailboxDatabase -Identity $mdb -ActivateOnServer $server `
					-SkipClientExperienceChecks:$SkipClientExperienceChecks -SkipMoveSuppressionChecks:$SkipMoveSuppressionChecks `
					-MoveComment:$moveComment -Confirm:$Confirm -ErrorAction:SilentlyContinue
				$moveResults += $lastMoveResult
				
				if (!$firstResultHasBeenSet)
				{
					$firstResultHasBeenSet = $true
					$firstMoveResult = $lastMoveResult
				}
				
				if ( $lastMoveResult -and `
					($lastMoveResult.Status -eq $MoveStatusType::Succeeded) -and `
					($lastMoveResult.ActiveServerAtEnd -eq $server) -and `
					($lastMoveResult.MountStatusAtMoveStart -eq $lastMoveResult.MountStatusAtMoveEnd))
				{
					$moveSuccessful = $true
					$moveStatus.MoveStatus = [HADatabaseLoadBalancing.MoveStatus]::MoveSucceeded
					Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0059 -f $mdb,$server,"Move-DatabaseToOrderedList")
					break;
				}	
				else
				{
					[string]$resultStr = $lastMoveResult | Format-Table | Out-String
					# Log-Verbose instead of Log-Error so that we can try another copy to move to. 
					Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0060 -f $mdb,$server,$false,$resultStr,"Move-DatabaseToOrderedList")
				}
			}
		}
		else
		{
			# This is a success case since we don't actually have to move the DB
			$moveSuccessful = $true
			$moveStatus.MoveStatus = [HADatabaseLoadBalancing.MoveStatus]::NoMoveAttempted
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0052 -f $mdb,"Move-DatabaseToOrderedList")
		}
	}
	finally
	{
		if (!$moveSuccessful)
		{
			&{
				# Cleanup code is run with "Continue" ErrorActionPreference
				$ErrorActionPreference = "Continue"

				# Initialize the originally known Active/Mount states to whatever the MailboxDatabase object was
				# However, since it could be stale by the time we got to moving this DB, we should rely more on the
				# $firstMoveResult returned by the first move attempt.
				[bool]$originallyMounted = $mdb.Mounted
				if ($firstMoveResult)
				{
					if ($firstMoveResult.ActiveServerAtStart)
					{
						$originalActive = $firstMoveResult.ActiveServerAtStart
					}
					if ($firstMoveResult.MountStatusAtMoveStart -ne $MountStatusType::Unknown)
					{
						$originallyMounted = (	($firstMoveResult.MountStatusAtMoveStart -eq $MountStatusType::Mounted) -or
												($firstMoveResult.MountStatusAtMoveStart -eq $MountStatusType::Mounting))
					}
				}
				
				# update more fields (its possible that $originalActive has changed since the start of this method)
				$moveStatus.ActiveServerAtStart = $originalActive
				$moveStatus.ActiveOnPreferenceAtStart = Get-ActivationPreferenceOfDatabaseCopy $mdb $originalActive
				
				Rollback-DatabaseOnFailedMove -desiredActiveServer $originalActive `
						-shouldBeMounted $originallyMounted -mdb $mdb -lastMoveResult $lastMoveResult
			}
		}
		elseif (!$MoveWhatIf)
		{
			if ($lastMoveResult)
			{
				$moveStatus.ActiveServerAtEnd = $lastMoveResult.ActiveServerAtEnd
				$moveStatus.ActiveOnPreferenceAtEnd = Get-ActivationPreferenceOfDatabaseCopy $mdb $lastMoveResult.ActiveServerAtEnd
			}
		}		
		
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0051 -f $moveSuccessful,"Move-DatabaseToOrderedList")
	}
	
	return $moveSuccessful
}

# This function tries to move the DB back to the specified server.
# It is only meant to be called as part of "rollback" of another move that failed.
# Any other usage of this method should be carefully evaluated.
# Typically, this should be called in a 'catch' or 'finally' block as part of error handling.
function Rollback-DatabaseOnFailedMove(
	[Parameter(Mandatory=$true)] [string]$desiredActiveServer, 
	[Parameter(Mandatory=$true)] [bool]$shouldBeMounted, 
	[Parameter(Mandatory=$true)] [Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $mdb,
	[Microsoft.Exchange.Management.SystemConfigurationTasks.DatabaseMoveResult] $lastMoveResult = $null)
{
	[string] $currentActive = $null
	[int] $activationPrefAtEnd = 0
	
	# Use the last move result if available
	if ($lastMoveResult)
	{
		if ($lastMoveResult.ActiveServerAtEnd)
		{
			$currentActive = $lastMoveResult.ActiveServerAtEnd
			$activationPrefAtEnd = Get-ActivationPreferenceOfDatabaseCopy $mdb $currentActive
		}
	}
	
	# If we still don't know what the active was from the Move result, lookup the MDB object again
	if (!$currentActive)
	{
		[Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase]$currentDb = Get-MailboxDatabase $mdb -Status
		$currentActive = Get-ActiveServerForDatabase $currentDb	
		$activationPrefAtEnd = Get-ActivationPreferenceOfDatabaseCopy $currentDb $currentActive
	}
	
	# Update the DB move status
	[HADatabaseLoadBalancing.DatabaseMoveStatus]$moveStatus = $script:databaseToMoveStatusTable[$mdb.Name]
	$moveStatus.ActiveServerAtEnd = $currentActive
	$moveStatus.ActiveOnPreferenceAtEnd = $activationPrefAtEnd
	
	
	# Case 1: The active server hasn't changed so we simply need to mount the DB if its not already mounted
	if ($desiredActiveServer -eq $currentActive)
	{
		if ($shouldBeMounted)
		{
			if($lastMoveResult -eq $null -or $lastMoveResult.MountStatusAtMoveEnd -ne [Microsoft.Exchange.Management.SystemConfigurationTasks.MountStatus].Mounted)
			{
				Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0061 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
				$mountCmd = { Mount-Database -Identity $mdb }
				if (TryExecute-ScriptBlock $mountCmd)
				{
					Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0062 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
					return
				}
				Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0063 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
			}
			return
		}
		else
		{
			# The DB needs to be left dismounted, but we won't explicitly dismount it in the script. 
			# The question "what could go wrong?" comes to mind...
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0064 -f $mdb,$currentActive,$shouldBeMounted,$false,"Rollback-DatabaseOnFailedMove")
			return		
		}
	}
	# Case 2: The active server has changed so we need to move the DB back
	else
	{
		Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0065 -f $mdb,"Move-DatabaseToOrderedList")
		# We'll try the most basic move first
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0066 -f $mdb,$currentActive,$mdb,$desiredActiveServer,$false,"Rollback-DatabaseOnFailedMove")
		$moveCmd = { Move-ActiveMailboxDatabase -Identity $mdb -ActivateOnServer $desiredActiveServer -Confirm:$false }
		if (TryExecute-ScriptBlock $moveCmd)
		{
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0067 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
			return
		}
		Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0068 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
		
		# Now, try the move with -SkipClientExperienceChecks in case ContentIndexing was preventing the move
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0069 -f $mdb,$currentActive,$mdb,$desiredActiveServer,$false,"Rollback-DatabaseOnFailedMove")
		$moveCmd = { Move-ActiveMailboxDatabase -Identity $mdb -ActivateOnServer $desiredActiveServer -SkipClientExperienceChecks -SkipMoveSuppressionChecks -Confirm:$false }
		if (TryExecute-ScriptBlock $moveCmd)
		{
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0070 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
			return
		}
		Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0071 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
	}
}

# Filter out servers that are activation blocked, activation disabled or down. Input is server distribution.
filter Filter-ServersDistributionActivationBlocked
{
	$tmpServer = $script:serverNameToMailboxServerTable[$_.ServerName]
	if ($tmpServer.DatabaseCopyAutoActivationPolicy -eq $AutoActivationType::Blocked)
	{
		Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0072 -f $_.ServerName)
	}
	elseif ($tmpServer.DatabaseCopyActivationDisabledAndMoveNow -eq $true)
	{
		# we use the same message as activionBlocked for now, as this is R6DC only. Utah is still an open book
		Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0072 -f $_.ServerName)
	}
	elseif (!(Is-DagServerOnline $_.ServerName))
	{
		Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0073 -f $_.ServerName)
	}
	else
	{
		$_
	}
}


# Returns a bool and an error message if validation fails.
# Returns $true if the given server is OK to move to for the given database; $false otherwise.
function Test-ServerForMove(
	[Parameter(Mandatory=$true)] [string] $sourceServerName,
	[Parameter(Mandatory=$true)] [string] $targetServerName)
{
	[string]$errMsg = ""
	$sourceServer = $script:serverNameToMailboxServerTable[$sourceServerName]
	$targetServer = $script:serverNameToMailboxServerTable[$targetServerName]
	$targetActivationPolicy = $targetServer.DatabaseCopyAutoActivationPolicy
	
	# Check DatabaseCopyAutoActivationAllowed policy setting on the target server
	if ($targetActivationPolicy -eq $AutoActivationType::Blocked)
	{
		$errMsg = $RedistributeActiveDatabases_LocalizedStrings.res_0098 -f $targetServerName,'Blocked'
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0074 -f $errMsg,"Test-ServerForMove")
		return $false,$errMsg
	}
	elseif ($targetServer.DatabaseCopyActivationDisabledAndMoveNow -eq $true)
	{
		# we use the same message as activionBlocked for now, as this is R6DC only. Utah is still an open book
		$errMsg = $RedistributeActiveDatabases_LocalizedStrings.res_0098 -f $targetServerName,'DatabaseCopyActivationDisabledAndMoveNow'
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0074 -f $errMsg,"Test-ServerForMove")
		return $false,$errMsg
	}
	elseif ($targetActivationPolicy -eq $AutoActivationType::Unrestricted)
	{
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0075 -f $targetServerName,"Test-ServerForMove")
	}
	elseif ($targetActivationPolicy -eq $AutoActivationType::IntrasiteOnly)
	{
		$targetSiteName = $script:serverNameToSiteTable[$targetServerName] 
		$sourceSiteName = $script:serverNameToSiteTable[$sourceServerName] 
		if ($targetSiteName -ne $sourceSiteName)
		{
			$errMsg = $RedistributeActiveDatabases_LocalizedStrings.res_0099 -f $targetServerName,'IntraSiteOnly',$sourceServerName
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0074 -f $errMsg,"Test-ServerForMove")
			return $false,$errMsg
		}
	}
	
	# Check if the target server is online according to clustering
	if (!(Is-DagServerOnline $targetServerName))
	{
		$errMsg = $RedistributeActiveDatabases_LocalizedStrings.res_0106 -f $targetServerName
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0074 -f $errMsg,"Test-ServerForMove")
		return $false,$errMsg
	}
	
	# Check if we are exceeding MaxActiveDatabases on target server
	if ($targetServer.MaximumActiveDatabases)
	{
		[int]$currentActives = $script:currentServerDbDistribution[$targetServerName].ActiveDbs
		[int]$maxActives = $targetServer.MaximumActiveDatabases
		if ($currentActives -ge $maxActives)
		{
			$errMsg = $RedistributeActiveDatabases_LocalizedStrings.res_0105 -f $targetServerName,$maxActives
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0074 -f $errMsg,"Test-ServerForMove")
			return $false,$errMsg
		}
	}
	
	# so far, the server is healthy
	return $true,$errMsg
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


# Common function to run a scriptblock, and log how long it took.
function Run-TimedScriptBlock (
	[System.Diagnostics.Stopwatch] $stopWatch = $null,
	[Parameter(Mandatory=$true)] [ScriptBlock] $scriptBlock,
	[Parameter(Mandatory=$true)] [string] $operationDescription ## eg: "Database lookup"
)
{
	if (!$stopWatch)
	{
		$stopWatch = New-Object -TypeName System.Diagnostics.Stopwatch
	}
	$stopWatch.Reset()
	$stopWatch.Start()
	
	# Execute the script block
	&$scriptBlock
		
	$stopWatch.Stop()
	Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0076 -f $stopWatch.Elapsed.TotalMilliseconds,$operationDescription,"Run-TimedScriptBlock")
}

# Compile the types to be used in this script
# The compilation is only performed once per runspace and is entirely in memory.
function Prepare-LoadBalancingDefinitions
{
	$code = '
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace HADatabaseLoadBalancing
{
    public abstract class DbDistributionEntryBase
    {
        public int TotalDbs { get { return ActiveDbs + PassiveDbs; } }
        public int ActiveDbs { get; set; }
        public int PassiveDbs { get; set; }
        public int[] PreferenceCountList { get; protected set; } 	// Array of integers indexed by (AP - 1), and whose value 
        // is the number of DBs at that AP.
        public int MountedDbs { get; set; }
        public int DismountedDbs { get; set; }
        public string DagName { get; set; }

        // Constructor
        protected DbDistributionEntryBase(int maxActivationPreference)
        {
            this.PreferenceCountList = new int[maxActivationPreference];
        }
    }

    public class ServerDbDistributionEntry : DbDistributionEntryBase, ICloneable
    {
        public string ServerName { get; set; }

        // Constructor
        public ServerDbDistributionEntry(int maxActivationPreference)
            : base(maxActivationPreference)
        {
        }

        public void AddDatabaseCopyOfPreference(int activationPreference)
        {
            this.PreferenceCountList[activationPreference - 1]++;
        }
		
		public void RemoveDatabaseCopyOfPreference(int activationPreference)
		{
			this.PreferenceCountList[activationPreference - 1]--;
		}
		
		public void ClearPreferenceCountList()
		{
			this.PreferenceCountList = new int[this.PreferenceCountList.Length];
		}

        #region ICloneable Members

        public object Clone()
        {
            // A shallow copy is good enough here since the Preference array is non-changing.
            ServerDbDistributionEntry clone = (ServerDbDistributionEntry)this.MemberwiseClone();
            return clone;
        }

        #endregion
    }

    public class SiteDbDistributionEntry : DbDistributionEntryBase, ICloneable
    {
        public string SiteName { get; set; }

        // Constructor
        public SiteDbDistributionEntry(int maxActivationPreference)
            : base(maxActivationPreference)
        {
        }

        public void AddServerDistribution(ServerDbDistributionEntry entry)
        {
            this.ActiveDbs += entry.ActiveDbs;
            this.PassiveDbs += entry.PassiveDbs;
            this.MountedDbs += entry.MountedDbs;
            this.DismountedDbs += entry.DismountedDbs;

            for (int i = 0; i < entry.PreferenceCountList.Length; i++)
            {
                this.PreferenceCountList[i] += entry.PreferenceCountList[i];
            }
        }

        #region ICloneable Members

        public object Clone()
        {
            // A shallow copy is good enough here since the Preference array is non-changing.
            SiteDbDistributionEntry clone = (SiteDbDistributionEntry)this.MemberwiseClone();
            return clone;
        }

        #endregion
    }

    public enum MoveStatus : int
    {
        NoMoveAttempted = 0,
        MoveSucceeded,
        MoveFailed
    }

    public class DatabaseMoveStatus
    {
        public string DbName { get; set; }
        public int ActiveOnPreferenceAtStart { get; set; }
        public string ActiveServerAtStart { get; set; }

        public int? ActiveOnPreferenceAtEnd { get; set; }
        public string ActiveServerAtEnd { get; set; }

        public bool IsOnMostPreferredCopy
        {
            get
            {
                if (!String.IsNullOrEmpty(this.ActiveServerAtEnd))
                {
                    return (this.ActiveOnPreferenceAtEnd == 1);
                }
                if (!String.IsNullOrEmpty(this.ActiveServerAtStart))
                {
                    return (this.ActiveOnPreferenceAtStart == 1);
                }
                // default
                return false;
            }
        }

        public MoveStatus MoveStatus { get; set; }

    }

	public static class DatabaseCopyPreferences
    {
        // Indexed by database name, and further indexed by server name. Value of: Activation Preference
        private static Dictionary<string, Dictionary<string, int>> m_dbCopyPreferences = 
			new Dictionary<string,Dictionary<string,int>>(200, StringComparer.OrdinalIgnoreCase);

        public static void AssignPreference(string dbName, string serverName, int activationPref)
        {
            if (!m_dbCopyPreferences.ContainsKey(dbName))
            {
                Dictionary<string, int> serverAPTable = new Dictionary<string, int>(5, StringComparer.OrdinalIgnoreCase);
                m_dbCopyPreferences[dbName] = serverAPTable;
            }

            m_dbCopyPreferences[dbName][serverName] = activationPref;
        }

        public static bool IsActivationPreferenceAssignedForDatabase(string dbName, int activationPref)
        {
            if (!m_dbCopyPreferences.ContainsKey(dbName))
            {
                return false;
            }

            Dictionary<string, int> serverAPTable = m_dbCopyPreferences[dbName];
            return serverAPTable.Any(kvp => kvp.Value == activationPref);
        }
		
		public static bool IsAnyActivationPreferenceAssignedForCopy(string dbName, string serverName)
        {
            if (!m_dbCopyPreferences.ContainsKey(dbName))
            {
                return false;
            }

            return m_dbCopyPreferences[dbName].ContainsKey(serverName);
        }
        
        public static int GetActivationPreferenceForCopy(string dbName, string serverName)
        {
            int ap = -1;
            if (IsAnyActivationPreferenceAssignedForCopy(dbName, serverName))
            {
                ap = m_dbCopyPreferences[dbName][serverName];
            }
            return ap;
        }
		
		public static int GetFirstMissingActivationPreferenceForDatabase(string dbName, int maxActivationPref)
        {
            Dictionary<string, int> serverAPTable = m_dbCopyPreferences[dbName];
            var prefs = 
                from kvp in serverAPTable
                orderby kvp.Value ascending
                select kvp.Value;

            int[] prefArray = prefs.ToArray();
            for (int i = 0; i < prefArray.Length; i++)
            {
                if (i != prefArray[i] - 1)
                {
                    // this is missing from the activation preferences
                    return i + 1;
                }
            }

            // the missing activation pref is the last one
            if (prefArray.Length < maxActivationPref)
            {
                return maxActivationPref;
            }

            // no missing AP found
            return 0;
        }
		
		// Returns a sorted (ascending) KeyValuePair with key of "database\Server", and value of ActivationPreference
        public static IEnumerable<KeyValuePair<string, int>> GetSortedDatabaseCopyPreferenceAssignments(string dbName)
        {
            Dictionary<string, int> serverAPTable = m_dbCopyPreferences[dbName];
            if (serverAPTable != null)
            {
                var returnList = 
                    from kvp in serverAPTable
                    orderby kvp.Value ascending
                    select new KeyValuePair<string, int>(dbName + "\\" + kvp.Key, kvp.Value);

                foreach (KeyValuePair<string, int> pair in returnList)
                {
                    yield return pair;
                }
            }
        }
		
		// Returns a sorted (ascending) KeyValuePair with key of "database\Server", and value of ActivationPreference
        public static IEnumerable<KeyValuePair<string, int>> GetAllSortedDatabaseCopyPreferenceAssignments()
        {
            var databases = from dbName in m_dbCopyPreferences.Keys
                            orderby dbName ascending
                            select dbName;

            foreach (string dbName in databases)
            {
                foreach (KeyValuePair<string, int> pair in GetSortedDatabaseCopyPreferenceAssignments(dbName))
                {
                    yield return pair;
                }
            }
        }
		
		public static void Clear()
        {
            m_dbCopyPreferences.Clear();
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
		[HADatabaseLoadBalancing.ServerDbDistributionEntry];
	}
	[bool]$isCompiled = TryExecute-ScriptBlock -runCommand:$checkCompiledCmd -silentOnErrors:$true

	if (!$isCompiled)
    {
        ##################################################################
        # So now we compile the code and use .NET object access to run it.
        ##################################################################
        
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0077 )
		Add-Type -TypeDefinition $code
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0078 )
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
	[HADatabaseLoadBalancing.EventLogger]::WriteLocalizedEvent( `
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
					Log-ErrorRecord $_ -Stop:$false
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
	Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0079 -f $sleepSecs)
	Start-Sleep $sleepSecs
}

# Common function to retrieve the current UTC time string
function Get-CurrentTimeString
{
	return [DateTime]::UtcNow.ToString("[HH:mm:ss.fff UTC]")
}

# Common function for printing a message
function Log-Info ( [string]$msg, [System.ConsoleColor] $color = "Gray")
{	
	Write-Host $msg -ForegroundColor:$color
}

# Common function for printing a success message
function Log-Success ( [string]$msg )
{
	Log-Verbose $msg
	Write-Host $msg -ForegroundColor:Green
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
	if ($Stop)
	{
		Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0080 -f $failedCommand,$failedMessage) -Stop:$Stop
	}
	else
	{
		Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0080 -f $failedCommand,$failedMessage)
	}
}


###################################################################
###  Entry point for the script itself
###################################################################

# Check if we should run on the local server
function CheckPAM
{
	if ($RunOnlyOnPAM -and !(Is-LocalServerPAM))
	{
		Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0081 -f $env:COMPUTERNAME)
		return $false
	}
	return $true
}

# Lookup the DAG either based on the specified name, or based on the local mailbox server.
function Resolve-DagParameter
{
	if (!$DagName)
	{
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0082 -f "Resolve-DagParameter"  )
		$server = Get-MailboxServer $env:COMPUTERNAME
		if ($server)
		{
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0083 -f $env:COMPUTERNAME,"Resolve-DagParameter")
			
			if ($server.DatabaseAvailabilityGroup)
			{
				$DagName = $server.DatabaseAvailabilityGroup.Name
				Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0084 -f $DagName,"Resolve-DagParameter")
			}
			else
			{
				Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0085 -f $env:COMPUTERNAME) -Stop
			}
		}
		else
		{
			Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0086 -f $env:COMPUTERNAME) -Stop
		}		
	}
	
	if (!$DagName)
	{
		Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0087 ) -Stop
	}
	else
	{
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0088 -f $DagName,"Resolve-DagParameter")
		$script:dag = Get-DatabaseAvailabilityGroup $DagName -Status
	}
		
	if (!$script:dag)
	{
		Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0089 -f $DagName) -Stop
	}
	
	if (!$SkipAutoRedistributeCheck -and !$script:dag.AutoDagAutoRedistributeEnabled)
	{
		Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0110 -f $DagName) -Stop
	}
}

# Prints error messages (if any) as warning
function Print-Errors
{
	if ($Error.Count -gt 0)
	{
		[string] $error_string = ""
		foreach ( $err in $Error )
		{
			$error_string += $err.ToString()
		}
		$Error.Clear()
		Log-Warning $error_string
	}
}

# Validates current database balance
function Validate-Balance
{
	if ( $MaxAllowedDbsDeltaForBalanceValidation -lt [int]::MaxValue)
	{		
		# Rebuild the hashtables describing what DBs are active/passive/mounted/etc on each server.
		Populate-DatabaseDistributionMap
	
		# Sort the server distribution in descending order (don't factor in the AD site).
		# But, don't include the servers that are ActivationBlocked!
		$descSortedDistrib = $script:currentServerDbDistribution.Values `
			| Sort-Object -Descending -Property ActiveDbs `
			| Filter-ServersDistributionActivationBlocked
		[int] $maxActives = ($descSortedDistrib | Select-Object -First 1).ActiveDbs
		[int] $minActives = ($descSortedDistrib | Select-Object -Last 1).ActiveDbs
		
		if ( $maxActives - $minActives -gt $MaxAllowedDbsDeltaForBalanceValidation)
		{
			Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0109 -f $minActives,$maxActives,$MaxAllowedDbsDeltaForBalanceValidation,"Validate-Balance") -Stop
		}
	}	
}

# Script execution starts here
function Main
{
	$script:startTime = Get-Date
	$script:databaseToMoveStatusTable.Clear()
	[System.Text.StringBuilder] $script:eventReport = New-Object -TypeName System.Text.StringBuilder -ArgumentList 2048
	
	if ($PSCmdlet.ParameterSetName -eq "Common")
	{
		Log-Warning ($RedistributeActiveDatabases_LocalizedStrings.res_0090 )
		return
	}
	
	if ($PSCmdlet.ParameterSetName  -eq "DotSourceMode")
	{
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0091 )
		return
	}

	# First, lookup the DAG
	Resolve-DagParameter
	
	# Error out early if we should only run on a PAM
	if (!(CheckPAM))
	{
		return
	}
	
	# Build the server hashtables
	Populate-ServerHashTables
	
	# Build the hashtables describing what DBs are active/passive/mounted/etc on each server.
	Populate-DatabaseDistributionMap
	
	# look up the cluster node status for the DAG
	Populate-ClusterNodeStatus
	
	if ($PSCmdlet.ParameterSetName -eq "ShowDatabaseCurrentActives")
	{
		# outputs DatabaseMoveStatus (defined in this script) objects to the pipeline
		Show-DatabaseCurrentActives $script:databases
		return
	}
	
	if ($PSCmdlet.ParameterSetName  -eq "ShowDatabaseDistributionByServer")
	{
		# output to the pipeline
		$script:serverDbDistribution.Values
		return
	}
	
	if ($PSCmdlet.ParameterSetName  -eq "ShowDatabaseDistributionBySite")
	{
		# output to the pipeline
		$script:siteDbDistribution.Values
		return
	}
	
	if ($PSCmdlet.ParameterSetName  -eq "ShuffleActiveDatabases")
	{
		# outputs DatabaseMoveStatus (defined in this script) objects to the pipeline
		Shuffle-ActiveDatabases $script:databases
		return
	}
	
	if ($PSCmdlet.ParameterSetName  -eq "BalanceDbsByActivationPreference")
	{
		try
		{
			$Error.Clear()
			# outputs DatabaseMoveStatus (defined in this script) objects to the pipeline
			$output = Balance-DbsByPreferenceOptimized
			if ($LogEvents)
			{
				# DatabaseRedistributionReport - Event 4115, Category (3) = Move
				Write-HAAppLogInformationEvent "40041013" 3 ($script:eventReport.ToString())
			}
		
			Write-Output $output
		}
		catch
		{
			Print-Errors
		}
		Validate-Balance
		return
	}
	
	if ($PSCmdlet.ParameterSetName  -eq "BalanceDbsBySiteAndActivationPreference")
	{
		try
		{
			$Error.Clear()
			# outputs DatabaseMoveStatus (defined in this script) objects to the pipeline
			$output = Balance-DbsBySiteAndActivationPreference
			if ($LogEvents)
			{
				# DatabaseRedistributionReport - Event 4115, Category (3) = Move
				Write-HAAppLogInformationEvent "40041013" 3 ($script:eventReport.ToString())
			}
		
			Write-Output $output
		}
		catch
		{
			Print-Errors
		}
		Validate-Balance
		return
	}
	
	if ($PSCmdlet.ParameterSetName  -eq "BalanceDbsIgnoringActivationPreference")
	{
		# outputs DatabaseMoveStatus (defined in this script) objects to the pipeline
		try
		{
			$Error.Clear()
			$output = Balance-DbsIgnoringActivationPreference
			if ($LogEvents)
			{
				# DatabaseRedistributionReport - Event 4115, Category (3) = Move
				Write-HAAppLogInformationEvent "40041013" 3 ($script:eventReport.ToString())
			}
		
			Write-Output $output
		}
		catch
		{
			Print-Errors
		}
		Validate-Balance
		return
	}
	
	if ($PSCmdlet.ParameterSetName  -eq "BalanceActivationPreferences")
	{
		Balance-ActivationPreferences
		return
	}
	
}


$Command = $MyInvocation.MyCommand
Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0092 -f $Command.Path)
# The command below is useful to see what parameters are defined in this script cmdlet.
# $Command | fl Path, CommandType, Parameters, ParameterSets

# Get the code compilation out of the way
Prepare-LoadBalancingDefinitions

LoadExchangeSnapin

Main

# SIG # Begin signature block
# MIIdvgYJKoZIhvcNAQcCoIIdrzCCHasCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUWmYYzuXs/CT8BwF6JbxUCA+8
# snugghhkMIIEwzCCA6ugAwIBAgITMwAAAJqamxbCg9rVwgAAAAAAmjANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBMQwggTAAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB2DAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUgmqPnH+DfhINtHrwOH1y9gHD5WUweAYKKwYB
# BAGCNwIBDDFqMGigQIA+AFIAZQBkAGkAcwB0AHIAaQBiAHUAdABlAEEAYwB0AGkA
# dgBlAEQAYQB0AGEAYgBhAHMAZQBzAC4AcABzADGhJIAiaHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQBoAwNBYf6SIufw
# T1J2zwY4ciTufH48GO/j4m6RfHdxETPYZ5RCpz1/XeHDqtfmZmlvmijatrhZVRp8
# KnnBvLFzRK9KJhJe+Ph9sGPljarzhs2ALQ7JognfhlYha41LVgA89Eq8ut3MdES6
# gjME0rEOAqq6+jOyj4ETFzk/oY/NrosCWpPBR0Hd8UcwJ2fW2K61xDO9JcZGSIcE
# uUWXWaXFKdbNT//KDPBtnu45W3xLsIcDXoIlB0mcAFLS2K2Gf7cWpmN7DjUgWCVN
# WXsvs6XlhMLZ8Ji4+MV8LM3d2bk7LdPHDeuJgH7HUJI7U6p0jy0UBkZxH0CjCxEu
# l9OmmZZooYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29m
# dCBUaW1lLVN0YW1wIFBDQQITMwAAAJqamxbCg9rVwgAAAAAAmjAJBgUrDgMCGgUA
# oF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTYw
# OTAzMTg0NDA4WjAjBgkqhkiG9w0BCQQxFgQUXkgmdtbeb8Pr0yjpsP7RXM3lYrMw
# DQYJKoZIhvcNAQEFBQAEggEASW+hGQE9wOE19kpx4HbKRE1xYzzZ3sfB64Nl3Xrq
# 4JtGendWwgyMfe+atTE9dRwzAYgt4sRKyJEGWDtZNk+LkW4+3BBbOqcF2ajDyEz5
# LUj8JZIVshEuQ6fMlxEaW9CgGp4FJ6CNmS6qxJd9ZC4r1/gessvWWTdS4yqkyLcF
# i9Yy6LEV9UD0R90z5JZ7lyoh2sqk6wLTqD3SQKxDl90ZlT32rJlNp6c0gBAWAZ0x
# +zfjPczkikEPZP1H4S2T+PrsmxMvsFdA3vXb0BAC2leHpEUbdNNRlUBiZO2ROaG+
# 6l6ZGEsld/inW3CixsqIAiZ3SWBEw3gF7E/lzYVVAvcZjw==
# SIG # End signature block
