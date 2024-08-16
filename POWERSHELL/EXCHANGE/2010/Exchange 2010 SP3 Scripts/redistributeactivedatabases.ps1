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
	#	PS D:\Exchange Mailbox\v14\Scripts> . .\RedistributeActiveDatabases.ps1 -DotSourceMode
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
	
	# If used, no changes are made (i.e. no DBs get moved around)
	[switch] $WhatIf = $false,
	
	# If used, causes Moves to request confirmation
	[switch] $Confirm = $true
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

Log-Info $msg

	if ($PSCmdlet.ParameterSetName  -eq "BalanceDbsBySiteAndActivationPreference")
	{
		$msg = "AllowedDeviationFromMeanPercentage :   $AllowedDeviationFromMeanPercentage %
AllowedMaxDelta                    :   $allowedMaxDelta
"
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
				# Since not all copies of this DB have been tried yet, let's return it.
				$_
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
	
	# Now assign preferences group by group
	foreach ($grouping in ($script:groupedDbs.Keys | Sort-Object))
	{
		[String[]]$serverNames = $grouping.Split("*")
		[int]$numCopies = $serverNames.Length
		$databases = @($script:groupedDbs[$grouping] | Sort-Object -Property Name)
		
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
		Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0046 -f $mdb)
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
				Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0050 -f $mdb,$originalActive,$currentActive,"Move-DatabaseToOrderedList")
				
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
				
				[string]$moveComment = $RedistributeActiveDatabases_LocalizedStrings.res_0100 -f "RedistributeActiveDatabases.ps1"
				$lastMoveResult = Move-ActiveMailboxDatabase -Identity $mdb -ActivateOnServer $server -MoveComment:$moveComment -Confirm:$Confirm -ErrorAction:Continue
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
		return $moveSuccessful
	}
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
	
	
	# Case 1: The active server hasn't changed so we simply need to mount the DB 
	if ($desiredActiveServer -eq $currentActive)
	{
		if ($shouldBeMounted)
		{
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0061 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
			$mountCmd = { Mount-Database -Identity $mdb }
			if (TryExecute-ScriptBlock $mountCmd)
			{
				Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0062 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
				return
			}
			Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0063 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
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
		Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0068 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
		
		# Now, try the move with -SkipClientExperienceChecks in case ContentIndexing was preventing the move
		Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0069 -f $mdb,$currentActive,$mdb,$desiredActiveServer,$false,"Rollback-DatabaseOnFailedMove")
		$moveCmd = { Move-ActiveMailboxDatabase -Identity $mdb -ActivateOnServer $desiredActiveServer -SkipClientExperienceChecks -Confirm:$false }
		if (TryExecute-ScriptBlock $moveCmd)
		{
			Log-Verbose ($RedistributeActiveDatabases_LocalizedStrings.res_0070 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
			return
		}
		Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0071 -f $mdb,$currentActive,"Rollback-DatabaseOnFailedMove")
	}
}

# Filter out servers that are activation blocked or down. Input is server distribution.
filter Filter-ServersDistributionActivationBlocked
{
	$tmpServer = $script:serverNameToMailboxServerTable[$_.ServerName]
	if ($tmpServer.DatabaseCopyAutoActivationPolicy -eq $AutoActivationType::Blocked)
	{
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
		Add-Type -TypeDefinition $code -Language "CSharpVersion3"
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
		finally
		{
			# Curious PS behavior: It appears that 'return' trumps 'throw', so don't return...
			if (!$throwOnError -or $success)
			{
				return $success
			}
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
	Log-Error ($RedistributeActiveDatabases_LocalizedStrings.res_0080 -f $failedCommand,$failedMessage) -Stop:$Stop
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
		# outputs DatabaseMoveStatus (defined in this script) objects to the pipeline
		$output = Balance-DbsByPreferenceOptimized
		if ($LogEvents)
		{
			# DatabaseRedistributionReport - Event 4115, Category (3) = Move
			Write-HAAppLogInformationEvent "40041013" 3 ($script:eventReport.ToString())
		}
		
		Write-Output $output
		return
	}
	
	if ($PSCmdlet.ParameterSetName  -eq "BalanceDbsBySiteAndActivationPreference")
	{
		# outputs DatabaseMoveStatus (defined in this script) objects to the pipeline
		$output = Balance-DbsBySiteAndActivationPreference
		if ($LogEvents)
		{
			# DatabaseRedistributionReport - Event 4115, Category (3) = Move
			Write-HAAppLogInformationEvent "40041013" 3 ($script:eventReport.ToString())
		}
		
		Write-Output $output
		return
	}
	
	if ($PSCmdlet.ParameterSetName  -eq "BalanceDbsIgnoringActivationPreference")
	{
		# outputs DatabaseMoveStatus (defined in this script) objects to the pipeline
		$output = Balance-DbsIgnoringActivationPreference
		if ($LogEvents)
		{
			# DatabaseRedistributionReport - Event 4115, Category (3) = Move
			Write-HAAppLogInformationEvent "40041013" 3 ($script:eventReport.ToString())
		}
		
		Write-Output $output
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
# MIIacgYJKoZIhvcNAQcCoIIaYzCCGl8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUGFz2LmSLwtn0wSnmjiDQ0Kou
# W8igghUmMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggS6MIIDoqADAgEC
# AgphAo5CAAAAAAAfMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQTAeFw0xMjAxMDkyMjI1NThaFw0xMzA0MDkyMjI1NThaMIGzMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMScwJQYD
# VQQLEx5uQ2lwaGVyIERTRSBFU046RjUyOC0zNzc3LThBNzYxJTAjBgNVBAMTHE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQCW7I5HTVTCXJWA104LPb+XQ8NL42BnES8BTQzY0UYvEEDeC6RQ
# UhKIC0N6LT/uSG5mx5HmA8pu7HmpaiObzWKezWqkP+ejQ/9iR6G0ukT630DBhVR+
# 6KCnLEMjm1IfMjX0/ppWn41jd3swngozhXIbykrIzCXN210RLsewjPGPQ0hHBbV6
# IAvl8+/BuvSz2M04j/shqj0KbYUX0MrnhgPAM4O1JcTMWpzEw9piJU1TJRRhj/sb
# 4Oz3R8aAReY1UyM2d8qw3ZgrOcB1NQ/dgUwhPXYwxbKwZXMpSCfYwtKwhEe7eLrV
# dAPe10sZ91PeeNqG92GIJjO0R8agVIgVKyx1AgMBAAGjggEJMIIBBTAdBgNVHQ4E
# FgQUL+hGyGjTbk+yINDeiU7xR+5IwfIwHwYDVR0jBBgwFoAUIzT42VJGcArtQPt2
# +7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3NvZnQu
# Y29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNybDBY
# BggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNVHSUE
# DDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAc/99Lp3NjYgrfH3jXhVx
# 6Whi8Ai2Q1bEXEotaNj5SBGR8xGchewS1FSgdak4oVl/de7G9TTYVKTi0Mx8l6uT
# dTCXBx0EUyw2f3/xQB4Mm4DiEgogOjHAB3Vn4Po0nOyI+1cc5VhiIJBFL11FqciO
# s3xybRAnxUvYb6KoErNtNSNn+izbJS25XbEeBedDKD6cBXZ38SXeBUcZbd5JhaHa
# SksIRiE1qHU2TLezCKrftyvZvipq/d81F8w/DMfdBs9OlCRjIAsuJK5fQ0QSelzd
# N9ukRbOROhJXfeNHxmbTz5xGVvRMB7HgDKrV9tU8ouC11PgcfgRVEGsY9JHNUaeV
# ZTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEFBQAwXzETMBEG
# CgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsG
# A1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5MB4XDTEw
# MDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENvZGUgU2lnbmlu
# ZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCycllcGTBkvx2a
# YCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPVcgDbNVcKicqu
# IEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlcRdyvrT3gKGiX
# GqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZC/6SdCnidi9U
# 3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgGhVxOVoIoKgUy
# t0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdcpReejcsRj1Y8
# wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTL
# EejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYBBAGCNxUBBAUC
# AwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy82C0wGQYJKwYB
# BAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBWJ5flJRP8KuEK
# U5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNyb3NvZnQuY29t
# L3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3JsMFQGCCsGAQUF
# BwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
# aS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcNAQEFBQADggIB
# AFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGjI8x8UJiAIV2s
# PS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbNLeNK0rxw56gN
# ogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y4k74jKHK6BOl
# kU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnpo1hW3ZsCRUQv
# X/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6H0q70eFW6NB4
# lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20OE049fClInHLR
# 82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8Z4L5UrKNMxZl
# Hg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9GuwdgR2VgQE6w
# QuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXrilUEnacOTj5XJ
# jdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEvmtzjcT3XAH5i
# R9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCCA++gAwIBAgIK
# YRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPyLGQBGRYDY29t
# MRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRNaWNyb3NvZnQg
# Um9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1MzA5WhcNMjEw
# NDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZUSNQrc7dGE4k
# D+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cOBJjwicwfyzMk
# h53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn1yjcRlOwhtDl
# KEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3U21StEWQn0gA
# SkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG7bfeI0a7xC1U
# n68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMBAAGjggGrMIIB
# pzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A+3b7syuwwzWz
# DzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1UdIwSBkDCBjYAU
# DqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/IsZAEZFgNjb20x
# GTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBS
# b290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0BxMuZTBQBgNV
# HR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9w
# cm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEESDBGMEQG
# CCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01p
# Y3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG
# 9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwTq86+e4+4LtQS
# ooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+jwoFyI1I4vBT
# Fd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwYTp2Oawpylbih
# OZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPfwgphjvDXuBfr
# Tot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5ZlizLS/n+YWG
# zFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8csu89Ds+X57H21
# 46SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUwZuhCEl4ayJ4i
# IdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHudiG/m4LBJ1S2
# sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9La9Zj7jkIeW1
# sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g74TKIdbrHk/J
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0xggS2
# MIIEsgIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQITMwAAAJ0ejSeu
# uPPYOAABAAAAnTAJBgUrDgMCGgUAoIHYMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBSeb1KrqfXKrC99Siq34OUU/B5G6jB4BgorBgEEAYI3AgEMMWowaKBAgD4AUgBl
# AGQAaQBzAHQAcgBpAGIAdQB0AGUAQQBjAHQAaQB2AGUARABhAHQAYQBiAGEAcwBl
# AHMALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2Ug
# MA0GCSqGSIb3DQEBAQUABIIBAGg1eXyvILBXhUuusum9hag9s4MeZNahpMESrY4u
# qxkcSEUV3MlV7YYmYgPm4e1qiHC6/u1kKkoCRLiYZ+zcaQDuJvJFZzZJZGiElefK
# F/i8hbgUV8VeZJf9PIG7ECimjTUqohjRtvL0ferROA1OCeqLZP9D/mpCpkpaQt8G
# Qe7PUWj/4dM0VDeWtP0fH6G5JJ6jKRtBqRqgwkE3MnV2p45bT193BxDb2YOocfVX
# YV0MSTa+AEnilPxxPJrqhLQ5NFA9Q+sCsYMeoEyKD4WHoqQIdwgJehZrxKREDfjU
# TMUFDQHLYYmKXBcmxYBuHF0BIsNTzoDdMkWugyvctWJb8XahggIfMIICGwYJKoZI
# hvcNAQkGMYICDDCCAggCAQEwgYUwdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAgph
# Ao5CAAAAAAAfMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0xMzAxMDgwODQ2NTlaMCMGCSqGSIb3DQEJBDEWBBTX
# nlEYYhVrpgtIgqKsPT3hBwiqfjANBgkqhkiG9w0BAQUFAASCAQB9qREme61rW1Sz
# 7VJFqRxsSBSA7+EKjKemmAX5OQM3vQULLE9mzXr7bXwxN2dPPFllGMz5NdILcX9F
# w3w9gRi1fEkkELDF/MBZEb8f+W684SuORZRmRI52xdPwIm3AaFmO4U/oWhwJ4qIL
# l/926HyFZKXW/W0368qo6YXe9gydr2LGQSwz1uykodzkrTJ/r6EPsGgQczJK8D/l
# GMHwvNy+8+bfiDGm5QqMLBqD0RnKoBTVe5Ygxwz+KKR040+vuGzDYQSoWDvkP6xC
# tFg4i6CgBFycYtbZPVjas42AdXmtSWNcc39kEgTgzGX0wJEMS0S0Zp8pk/kTMkOD
# U+maxKRB
# SIG # End signature block
