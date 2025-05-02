# Copyright (c) Microsoft Corporation. All rights reserved.
#
# DagCommonLibrary.ps1
# A collection of DAG-related functions for use by other scripts.


Import-LocalizedData -BindingVariable DagCommonLibrary_LocalizedStrings -FileName DagCommonLibrary.strings.psd1

# Checks for RSAT-Clustering
function Test-RsatClusteringInstalled
{
	if ( test-path $( [System.IO.Path]::Combine( $env:windir, "system32\cluster.exe" ) ) )
	{
		log-verbose ($DagCommonLibrary_LocalizedStrings.res_0000 -f "RSAT-Clustering");
	}
	else
	{
		log-error ($DagCommonLibrary_LocalizedStrings.res_0001 -f "RSAT-Clustering");
	}
}

# Returns the objects that are hosted on the specified mailbox server.
# The checks that are done include:
# -No replicated mailboxes are sources.
# -Any non-replicated mailbox contains no mailboxes (regular or arbitration).
# -Not the Primary Active Manager (PAM).
# -No active mailbox databases ('sources').
#
# If the server is NOT being monitored, then the following is also checked:
# -Any target replicated mailbox databases are suspended.
#
# Only the objects are returned, not the reason that they're returned.
#
# NOTE: There is no 'PAM' object, so if the PAM is owned by the
# server, a string 'Primary Active Manager=<servername>' is returned instead.
function GetCriticalMailboxResources(
	[string] $serverName,
	[bool] $serverIsMonitored = $false,
	[UInt32] $AtLeastNCriticalCopies
)
{
	Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0002 -f $serverName,$serverIsMonitored)
	
	# Look for the mailbox server so that we can find out if it belongs to a DAG 
	$mbxServer = get-mailboxserver $serverName;

	$clusterService = get-service clussvc -ErrorAction:SilentlyContinue
	[bool] $clusterServiceDown = $false
	if ( ( $clusterService -eq $null ) -or ( $clusterService.Status -ne 'Running' ) )
	{
		$clusterServiceDown = $true
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0003 -f $serverName,"Get-CriticalMailboxResources")
	}
	
	if ( $mbxServer.DatabaseAvailabilityGroup )
	{
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0004 -f $mbxServer.DatabaseAvailabilityGroup,$serverName)

	    $dag = Get-DatabaseAvailabilityGroup $mbxServer.DatabaseAvailabilityGroup -status;
	    if (!$dag)
	    {
			Log-Error ($DagCommonLibrary_LocalizedStrings.res_0005 -f $serverName) -Stop
	    }

		$nodeCount = Get-NodeCountOfDag $dag;
		if ( $nodeCount -lt 2 )
		{
			Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0006 -f $nodeCount);
		}
		else
		{
			$pamServerName = $dag.PrimaryActiveManager
			Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0012 -f $pamServerName)
			if ( $serverName -ieq $pamServerName )
			{
				Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0013 -f $serverName);
				write-output "(Primary Active Manager=$serverName)"
			}
		}
	}
	else
	{
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0014 -f $serverName)
	}
	
	# Check for replicated databases.
	Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0015 -f $serverName,"non-replicated")
	$databases = get-mailboxdatabase -server $serverName
	$numDatabases = ($databases | Measure-Object).Count
	Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0016 -f $numDatabases,$serverName)

	$setupRegistryPath = Get-ItemProperty -path 'HKLM:SOFTWARE\Microsoft\ExchangeServer\v14\Setup'
	$exchangeInstallPath = $setupRegistryPath.MsiInstallPath
	$redundacyCheck = [System.IO.Path]::Combine($exchangeInstallPath, "Scripts\CheckDatabaseRedundancy.ps1")
	if ( -not ( Test-Path $redundacyCheck ) )
	{
		Log-Error ($DagCommonLibrary_LocalizedStrings.res_0017 -f $redundacyCheck) -Stop
	}
	if ($AtleastNCriticalCopies)
	{
		. $redundacyCheck -DotSourceMode -AtLeastNCopies $AtleastNCriticalCopies
	}
	else
	{
		. $redundacyCheck -DotSourceMode
	}
	
	if ( $numDatabases -gt 0 )
	{
		foreach ( $database in $databases )
		{
			if ( $database.ReplicationType -eq 'Remote' )
			{
				# Make sure replicated databases aren't hosted on the server.
				# And if they aren't hosted, they need to be suspended.
				$status = get-mailboxdatabasecopystatus $database\$serverName

				Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0018 -f $status.ActiveCopy,$status.Status,$status.ActivationSuspended,$database,"Get-CriticalMailboxResources");

				if ( $status.ActiveCopy )
				{
					Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0019 -f $status.ActiveCopy,$status.Status,$status.ActivationSuspended,$database,"Get-CriticalMailboxResources");
					write-output "(Database='$($database.Name)', Reason='Copy is active')";
				}
				elseif ( $status.Status -eq 'ServiceDown' )
				{
					# Check if the copy is not active
					$statuses = @( get-mailboxdatabasecopystatus $database )
					[bool] $foundAnotherActiveCopy = $false
					foreach ( $copyStatus in $statuses )
					{
						if ( $copyStatus.ActiveCopy -and ( $copyStatus.MailboxServer.ToUpper() -ne $serverName.ToUpper() ) )
						{
							$foundAnotherActiveCopy = $true
							Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0020 -f $copyStatus.MailboxServer,$database,"Get-CriticalMailboxResources");							
							break;
						}
					}
					if ( -not $foundAnotherActiveCopy )
					{
						Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0019 -f $status.ActiveCopy,$status.Status,$status.ActivationSuspended,$database,"Get-CriticalMailboxResources");
						write-output "(Database='$($database.Name)', Reason='Copy is potentially active while replay service is down')";
					}
					elseif ( $mbxServer.DatabaseCopyAutoActivationPolicy -ne 'Blocked' )
					{
						Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0021 -f $status.ActiveCopy,$status.Status,$status.ActivationSuspended,$database,$serverName,"Get-CriticalMailboxResources");
						write-output "(Database='$($database.Name)', Reason='Service is not running on $serverName and database copy auto activation policy is not Blocked. Database copy auto activation policy is $($mbxServer.DatabaseCopyAutoActivationPolicy)')";
					}					
				}
				elseif ( ( $status.Status -inotlike '*suspended*' ) -and ( ! $status.ActivationSuspended ) )
				{
					# If the server is monitored, then we don't want suspended copies hanging around.
					# If the server is no longer being monitored, the copies
					# should be suspended (as this is preparation for some sort
					# of maintenance).
					if ( ! $serverIsMonitored )
					{
						Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0022 -f $status.ActiveCopy,$status.Status,$status.ActivationSuspended,$database,"Get-CriticalMailboxResources");
						write-output "(Database='$($database.Name)', Reason='Copy is not suspended, it is $($status.Status) and the activation suspended flag is $($status.ActivationSuspended)')";
					}
				}
				elseif ( -not ( Check-DatabaseRedundancyForCopyRemoval $database $serverName ) )
				{
					Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0023 -f $database,$serverName,$database,$serverName)
					write-output "(Database='$($database.Name)', Reason='Copy is critical for redundancy according to Red Alert script')";
				}
			}
			else
			{
				Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0024 -f $database);

				# For non-replicated databases, check for mailboxes.
				$tempMailboxes = Get-Mailbox -Database $database
				if ( $tempMailboxes )
				{
					$numMailboxes = ($tempMailboxes | Measure-Object).Count
					Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0025 -f $database,$numMailboxes);

					foreach ( $mbx in $tempMailboxes )
					{
						write-output "(Mailbox='$mbx', Reason='Mailbox is hosted on '$($database.Name)', which is not a replicated database. )";
					}
				}

				$tempMailboxes = Get-Mailbox -Database $database -Arbitration
				if ( $tempMailboxes )
				{
					$numMailboxes = ($tempMailboxes | Measure-Object).Count
					Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0026 -f $database,$numMailboxes);

					foreach ( $mbx in $tempMailboxes )
					{
						write-output "(Mailbox='$mbx', Reason='Arbitration Mailbox is hosted on '$($database.Name)', which is not a replicated database. )";
					}
				}
			}
		}
	}
	
	Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0027 )
}



# Gets the count of the servers in the DAG.
function Get-NodeCountOfDag( [object] $dag )
{
	$nodeCount = $dag.Servers.Count;
	Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0028 -f $nodeCount,"Get-NodeCountOfDag");
	return $nodeCount;
}



# Prints the output of GetCriticalMailboxResources
function PrintCriticalMailboxResourcesOutput( [object[]]$resources )
{
	[string]$resString = ( $resources | % { `
		if ($_.GetType() -eq [string]) 
		{
			$_
		}
		else
		{
			# every Exchange presentation object should have a Name property
			$_.Name
		}
	} )
	
	$resString;
}

# This function logs critical resources detailed information
function Log-CriticalResource([object[]] $resources) {

	foreach ( $resource in $resources )
	{
		switch ($resource.GetType().Name) {
			"String" { log-verbose ($DagCommonLibrary_LocalizedStrings.res_0029 -f $resource) }
			"Mailbox" { log-verbose ($DagCommonLibrary_LocalizedStrings.res_0030 -f $resource.Name,$resource.Database.Name) }
			"DatabaseCopyStatusEntry" { log-verbose ($DagCommonLibrary_LocalizedStrings.res_0031 -f $resource.Name,$resource.Status) }
			"MailboxDatabase" { log-verbose ($DagCommonLibrary_LocalizedStrings.res_0032 -f $resource.Name) }
			default { log-verbose ($DagCommonLibrary_LocalizedStrings.res_0029 -f $resource.Name) }
		}
	}
}

# Moves the PAM from the given server if necessary
# Supports $DagScriptTesting for whatif processing.
function Move-DagActiveManager([string]$Server)
{
	Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0033 -f $Server,"Move-DagActiveManager")
	
    # Move PAM if the target is a PAM.
	$mbxServer = get-mailboxserver $server -erroraction:silentlycontinue
	if ( $mbxServer -and $mbxServer.DatabaseAvailabilityGroup )
	{
		$dag = Get-DatabaseAvailabilityGroup $mbxServer.DatabaseAvailabilityGroup -status;
	}

    if (!$dag)
    {
		Log-Error ($DagCommonLibrary_LocalizedStrings.res_0034 -f $Server,"Move-DagActiveManager") -Stop
    }
    
	$nodeCount = Get-NodeCountOfDag $dag;
	if ( $nodeCount -lt 2 )
	{
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0035 -f $nodeCount,"Move-DagActiveManager");
		return;
	}
		
	if ( $Server -ieq $dag.PrimaryActiveManager )
	{
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0038 -f $Server,"Move-DagActiveManager")
		
		# Pick another target.
		$targetServer = $null
		$dag.Servers | where { $_.Name -ne $Server } | foreach { $targetServer = $_ }
	        
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0039 -f $dag.Name,$targetServer,"Move-DagActiveManager")
		if ( $DagScriptTesting )
		{
			write-host ($DagCommonLibrary_LocalizedStrings.res_0040 )
		}
		else
		{            
           	if (![Microsoft.Exchange.Cluster.Replay.DagTaskHelperPublic]::MovePrimaryActiveManagerRole($server)) {
                Log-Error -Stop ($DagCommonLibrary_LocalizedStrings.res_0109 -f "Move-DagActiveManager","couldn't move the cluster group")
            }
        }
    }
    else
    {
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0041 -f $dag.PrimaryActiveManager,$Server,"Move-DagActiveManager")
    }                
        
    Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0042 -f "Move-DagActiveManager")
}

# Move the critical resources off the specified server. 
# This includes Active Databases, and the Primary Active Manager.
function Move-CriticalMailboxResources([string]$Server)
{
	$moveResCmd = 
	{
		# 1. Move any mailbox database masters to other servers.
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0043 -f $Server,"Move-CriticalMailboxResources")
		Move-DagActiveCopy -MailboxServer $Server
		
		# 2. Move PAM if the server is a PAM.
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0044 -f $Server,"Move-CriticalMailboxResources")
		Move-DagActiveManager -Server $Server
		
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0045 -f $Server,"Move-CriticalMailboxResources")
	}
	
	$errorCmd = { Log-Error ($DagCommonLibrary_LocalizedStrings.res_0046 -f $Server,"Move-CriticalMailboxResources") }
	TryExecute-ScriptBlock -runCommand $moveResCmd -cleanupCommand $errorCmd -throwOnError $true
}

# Entry point of the lossless switchover function
# If both $Database and $MailboxServer are provided, it moves only that database
# If only $MailboxServer is provided, it moves all the active databases on the server
#
function Move-DagActiveCopy([string]$MailboxServer, [string]$Database)
{
	Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0047 -f $MailboxServer,$Database,"Move-DagActiveCopy")
	
    $successdb = @()
    $activedb = @()
	$sourceServer = Get-ExchangeServer | where {$_.Name -eq $MailboxServer}
    if (($Database) -and ($sourceServer))
    {
        $sourcedb = Get-MailboxDatabase -Identity $Database -Status
        Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0048 -f $sourcedb.MountedOnServer,$Database,"Move-DagActiveCopy")
        
        if ($sourcedb.MountedOnServer -eq $sourceServer.Fqdn)
        {
			Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0049 -f $sourcedb,$sourceServer,"Move-DagActiveCopy")
            if (Move-DagMasterCopy -db $sourcedb -srcServer $sourceServer)
            {
				if ( ! $DagScriptTesting )
				{
					# REVIEW: Is this 60 seconds still needed?!?
                	Sleep-ForSeconds 60
				}
                Verify-DagMasterCopy -db $sourcedb -srcServer $sourceServer -StopOnError
            }
            else
            {
				Log-Error ($DagCommonLibrary_LocalizedStrings.res_0050 -f $Database,$MailboxServer,"Move-DagActiveCopy")
            }
        }
        else
        {
            Log-Error ($DagCommonLibrary_LocalizedStrings.res_0051 -f $Database,$MailboxServer,"Move-DagActiveCopy")
        }
    }
    elseif ($sourceServer)
    {
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0052 -f $sourceServer,"Move-DagActiveCopy")
        $activedb = Get-MailboxDatabase -Server $sourceServer.Name -Status | where { ( $_.MountedOnServer -eq $sourceServer.Fqdn ) -and ( $_.ReplicationType -eq 'Remote' )}
        if ($activedb)
        {
            foreach ($sourcedb in $activedb)
            {
				Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0053 -f $sourcedb,$sourceServer,"Move-DagActiveCopy")
                if (Move-DagMasterCopy -db $sourcedb -srcServer $sourceServer) 
                {
					$successdb += $sourcedb
				}
				else
				{
					Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0054 -f $sourcedb,$sourceServer,"Move-DagActiveCopy")
				}
            }
            if ($successdb)
            {
				if ( ! $DagScriptTesting )
				{
					# REVIEW: Is this 60 seconds still needed?!?
                	Sleep-ForSeconds 60
				}
                $successdb | foreach { Verify-DagMasterCopy -db $_ -srcServer $sourceServer -StopOnError }
            }
        }
        else
        {
            Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0055 -f $MailboxServer,"Move-DagActiveCopy")
        }
    }
    else
    {
        Log-Error ($DagCommonLibrary_LocalizedStrings.res_0056 -f "Move-DagActiveCopy")
    }
    
    Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0042 -f "Move-DagActiveCopy")
}

# This function picks up the best database copy based on the switchover criteria and moves the master to it.
# Currently, it picks up the server closest to the front on the servers property as the
# most preferrable copy as long as it meets the criteria.
# If the PreferredTarget parameter is specified, it attempts to move the database only to the specified server
# and it does not try other servers.
# If the move fails, this function will attempt to "rollback" the DB to the original active server to prevent
# an extended DB outage.
function Move-DagMasterCopy ([object] $db, [object] $srcServer, [string] $preferredTarget)
{
	Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0057 -f $db,$srcServer,$preferredTarget,"Move-DagMasterCopy")
	
	[bool]$moveSuccessful = $true;
    $targetdbs = @()
	
    if ($preferredTarget)
    {
    	# $db\server does not accept a FQDN for the server.
    	$shortName = $preferredTarget  -replace "\..*$";

        $dbCopyStatus = Get-MailboxDatabaseCopyStatus $db\$shortName;
        if (Test-DagTargetCopy -copyStatus $dbCopyStatus -Lossless) {$targetdbs += $dbCopyStatus}
    }
    else
    {
        foreach ($Server in $db.Servers)
        {
            if ($Server.Name -ne $srcServer.Name)
            {
                $dbCopyStatus = Get-MailboxDatabaseCopyStatus $db\$Server
                if (Test-DagTargetCopy -copyStatus $dbCopyStatus -Lossless) {$targetdbs += $dbCopyStatus}
            }
        }
    }

    $numTargetDbs = ($targetdbs | Measure-Object).Count
    Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0058 -f $db.Servers.Count,$numTargetDbs,$db,"Move-DagMasterCopy")
    
	try
	{
		if ($targetdbs)
		{
			foreach ($targetdb in $targetdbs)
			{
				Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0059 -f $targetdb.MailboxServer,$db,$false,"Move-DagMasterCopy","Move-ActiveMailboxDatabase","-Identity","-ActivateOnServer","-SkipClientExperienceChecks","-Confirm")
				if ( $DagScriptTesting )
				{
					$moveCmd = { write-host ($DagCommonLibrary_LocalizedStrings.res_0060 -f $db,$targetdb,$false,"Move-DagMasterCopy","Move-ActiveMailboxDatabase","-Identity","-ActivateOnServer","-SkipClientExperienceChecks","-Confirm") }
				}
				else
				{
					$moveCmd = { Move-ActiveMailboxDatabase -Identity $db -ActivateOnServer $targetdb.MailboxServer -SkipClientExperienceChecks -Confirm:$false }
				}

				if (TryExecute-ScriptBlock $moveCmd)
				{
					Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0061 -f $targetdb.MailboxServer,$db,$false,"Move-DagMasterCopy","Move-ActiveMailboxDatabase","-Identity","-ActivateOnServer","-SkipClientExperienceChecks","-Confirm")
					break;
				}
				else
				{
					# Log-Verbose instead of Log-Error so that we can try another copy to move to. 
					# In any case, we will verify the move using Verify-DagMasterCopy after this function returns.
					$moveSuccessful = $false;
					Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0062 -f $targetdb.MailboxServer,$db,$false,"Move-DagMasterCopy","Move-ActiveMailboxDatabase","-Identity","-ActivateOnServer","-SkipClientExperienceChecks","-Confirm")
				}
			}
		}
		else
		{
			# When moving all the databases on a server, the script continues even with this error to move
			# other "good" databases. When moving a single database, it ends up with no ops.
			$moveSuccessful = $false;
			Log-Error ($DagCommonLibrary_LocalizedStrings.res_0063 -f $db,"Move-DagMasterCopy") -Stop
		}
    }
	finally
	{
		if (!$moveSuccessful)
		{
			&{
				# Cleanup code is run with "Continue" ErrorActionPreference
				$ErrorActionPreference = "Continue"
				Log-Error ($DagCommonLibrary_LocalizedStrings.res_0064 -f $db,"Move-DagMasterCopy")

				[string]$originalActive = $db.MountedOnServer -replace "\..*$"
				[bool]$originalMountState = $db.Mounted
				Rollback-DatabaseOnFailedMove -desiredActiveServer $originalActive -shouldBeMounted $originalMountState -db $db
				
				Log-Error ($DagCommonLibrary_LocalizedStrings.res_0065 -f $db,"Move-DagMasterCopy") -Stop
			}
		}
	
    	Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0066 -f $moveSuccessful,"Move-DagMasterCopy")
    	return $moveSuccessful
	}
}


#--------------------------------------------------------------------------------
# Check whether the specific database copy meets the criteria for a switchover or failover
# This function is called by the Move-DagActiveCopy and Execute-LossyDatabaseFailOver functions
#
# CI status is ignored for lossy moves.  The calling function will need to decide if that is suitable.
#
function Test-DagTargetCopy ([object]$copyStatus, [switch]$Lossless, [switch]$CICheck=$false)
{
    Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0067 -f $copyStatus.Name,$Lossless,$CICheck,"Test-DagTargetCopy")
    Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0068 -f $copyStatus.Name,$copyStatus.Status,$copyStatus.ContentIndexState,$copyStatus.CopyQueueLength,$copyStatus.ReplayQueueLength,"Test-DagTargetCopy")
    
    [bool]$healthy = $false
    if ($Lossless) {$allowedCopyQueueLength = 10} else {$allowedCopyQueueLength = 50}
	
    if (    ($copyStatus.Status -eq 'Healthy') `
            -and ($copyStatus.CopyQueueLength -le $allowedCopyQueueLength) `
            -and ($copyStatus.ReplayQueueLength -le 100))
    {
		if ( !$CICheck )
		{
			$healthy = $true;
		}
		else
		{
			if (($copyStatus.ContentIndexState -eq 'Healthy') -or !$Lossless)
			{
				$healthy = $true
			}
		}
    }
    
    Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0066 -f $healthy,"Test-DagTargetCopy")
    return $healthy
}

# Verify that the new active is mounted
# This function can be called with or without the srcServer property.
# If it is called with srcServer, it assumes that the database is mounted on another server.
# If it is not, it assumes that the database may be mounted on the original server.
#
# Supports $DagScriptTesting for -whatif
function Verify-DagMasterCopy ([object] $db, [object] $srcServer, [switch]$StopOnError)
{
	Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0069 -f $db,$srcServer,"Verify-DagMasterCopy")
	
    $moveddb = Get-MailboxDatabase -Identity $db -Status
    $srcServerName = $srcServer.Name
    $targetServerName = $moveddb.MountedOnServer -replace "\..*$"
	$originallyMounted = $db.Mounted
	
	if ($originallyMounted -eq $true)
	{
	    if ($moveddb.Mounted -eq $true)
	    {
	        if ($srcServerName)
	        {
				if ($srcServerName -ne $targetServerName)
				{
					Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0070 -f $db,$srcServerName,$targetServerName,"Verify-DagMasterCopy")
	            }
	            else
	            {
					if ( $DagScriptTesting )
					{
						write-host ($DagCommonLibrary_LocalizedStrings.res_0071 -f $db,$srcServerName,"Verify-DagMasterCopy")
					}
					else
					{
						Log-Error ($DagCommonLibrary_LocalizedStrings.res_0072 -f $db,$srcServerName,"Verify-DagMasterCopy") -Stop:$StopOnError
					}
	            }
	        }
	        else
	        {
	            Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0073 -f $db,$targetServerName,"Verify-DagMasterCopy")
	        }
	    }
	    else
	    {
			if ( $DagScriptTesting )
			{
				write-host ($DagCommonLibrary_LocalizedStrings.res_0074 -f $db,$targetServerName,"Verify-DagMasterCopy","Move-ActiveMailboxDatabase")
			}
			else
			{
		        Log-Error ($DagCommonLibrary_LocalizedStrings.res_0075 -f $db,$targetServerName,"Verify-DagMasterCopy","Move-ActiveMailboxDatabase") -Stop:$StopOnError
			}
	    }
    }
    elseif ($originallyMounted -eq $false)
    {
	    if ($moveddb.Mounted -eq $true)
	    {
			if ( $DagScriptTesting )
			{
				write-host ($DagCommonLibrary_LocalizedStrings.res_0076 -f $db,$targetServerName,"Verify-DagMasterCopy","Move-ActiveMailboxDatabase")
			}
			else
			{
	    		Log-Error ($DagCommonLibrary_LocalizedStrings.res_0077 -f $db,$targetServerName,"Verify-DagMasterCopy","Move-ActiveMailboxDatabase") -Stop:$StopOnError
			}
	    }
	    else
	    {
	    	if ($srcServerName)
	    	{
	    		if ($srcServerName -ne $targetServerName)
				{
					Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0078 -f $db,$srcServerName,$targetServerName,"Verify-DagMasterCopy")
				}
		        else
		        {
					if ( $DagScriptTesting )
					{
						write-host ($DagCommonLibrary_LocalizedStrings.res_0079 -f $db,$srcServerName,"Verify-DagMasterCopy")
					}
					else
					{
						Log-Error ($DagCommonLibrary_LocalizedStrings.res_0080 -f $db,$srcServerName,"Verify-DagMasterCopy") -Stop:$StopOnError
					}
		        }
	    	}
	    	else
		    {
				Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0081 -f $db,$targetServerName,"Verify-DagMasterCopy")
		    }
	    }
    }
	else
	{
		if ( $DagScriptTesting )
		{
			write-host ($DagCommonLibrary_LocalizedStrings.res_0082 -f $db,"Verify-DagMasterCopy","Verify-DagMasterCopy","get-mailboxdatabase","-status")
		}
		else
		{
			Log-Error ($DagCommonLibrary_LocalizedStrings.res_0083 -f $db,"Verify-DagMasterCopy","Verify-DagMasterCopy","get-mailboxdatabase","-status") -Stop:$StopOnError
		}
	}
    
    Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0042 -f "Verify-DagMasterCopy")
}

# This function tries to move the DB back to the specified server.
# It is only meant to be called as part of "rollback" of another move that failed.
# Any other usage of this method should be carefully evaluated.
# Typically, this should be called in a 'catch' or 'finally' block as part of error handling.
function Rollback-DatabaseOnFailedMove([string]$desiredActiveServer, [bool]$shouldBeMounted, [Object]$db)
{
	$currentDb = Get-MailboxDatabase -Identity $db -Status
	$currentActive = $currentDb.MountedOnServer -replace "\..*$"
	
	# Case 1: The active server hasn't changed so we simply need to mount the DB 
	if ($desiredActiveServer -eq $currentActive)
	{
		if ($shouldBeMounted)
		{
			Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0084 -f $db.Name,$currentActive,"Rollback-DatabaseOnFailedMove")
			if ( $DagScriptTesting )
			{
				$mountCmd = { write-host ($DagCommonLibrary_LocalizedStrings.res_0085 -f $currentDb,"Rollback-DatabaseOnFailedMove","Mount-Database","-Identity") }
			}
			else
			{
				$mountCmd = { Mount-Database -Identity $currentDb }
			}
			if (TryExecute-ScriptBlock $mountCmd)
			{
				Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0086 -f $db.Name,$currentActive,"Rollback-DatabaseOnFailedMove")
				return
			}
			Log-Error ($DagCommonLibrary_LocalizedStrings.res_0087 -f $db.Name,$currentActive,"Rollback-DatabaseOnFailedMove")
			return
		}
		else
		{
			# The DB needs to be left dismounted, but we won't explicitly dismount it in the script. 
			Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0088 -f $db.Name,$currentActive,$shouldBeMounted,$false,"Rollback-DatabaseOnFailedMove")
			return		
		}
	}
	# Case 2: The active server has changed so we need to move the DB back
	else
	{
		# We'll try the most basic move first
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0089 -f $db.Name,$currentActive,$currentDb,$desiredActiveServer,$false,"Rollback-DatabaseOnFailedMove","Move-ActiveMailboxDatabase","-Identity","-ActivateOnServer","-Confirm")

		if ( $DagScriptTesting )
		{
			$mountCmd = { write-host ($DagCommonLibrary_LocalizedStrings.res_0090 -f $currentDb,$desiredActiveServer,$false,"Rollback-DatabaseOnFailedMove","Move-ActiveMailboxDatabase","-Identity","-ActivateOnServer","-Confirm") }
		}
		else
		{
			$moveCmd = { Move-ActiveMailboxDatabase -Identity $currentDb -ActivateOnServer $desiredActiveServer -Confirm:$false }
		}
		if (TryExecute-ScriptBlock $moveCmd)
		{
			Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0091 -f $db.Name,$currentActive,"Rollback-DatabaseOnFailedMove")
			return
		}
		Log-Error ($DagCommonLibrary_LocalizedStrings.res_0092 -f $db.Name,$currentActive,"Rollback-DatabaseOnFailedMove")
		
		# Now, try the move with -SkipClientExperienceChecks in case ContentIndexing was preventing the move
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0093 -f $db.Name,$currentActive,$currentDb,$desiredActiveServer,$false,"Rollback-DatabaseOnFailedMove","Move-ActiveMailboxDatabase","-Identity","-ActivateOnServer","-SkipClientExperienceChecks","-Confirm")
		if ( $DagScriptTesting )
		{
			$mountCmd = { write-host ($DagCommonLibrary_LocalizedStrings.res_0094 -f $currentDb,$desiredActiveServer,$false,"Rollback-DatabaseOnFailedMove","Move-ActiveMailboxDatabase","-Identity","-ActivateOnServer","-SkipClientExperienceChecks","-Confirm") }
		}
		else
		{
			$moveCmd = { Move-ActiveMailboxDatabase -Identity $currentDb -ActivateOnServer $desiredActiveServer -SkipClientExperienceChecks -Confirm:$false }
		}
		if (TryExecute-ScriptBlock $moveCmd)
		{
			Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0095 -f $db.Name,$currentActive,"Rollback-DatabaseOnFailedMove")
			return
		}
		Log-Error ($DagCommonLibrary_LocalizedStrings.res_0096 -f $db.Name,$currentActive,"Rollback-DatabaseOnFailedMove")
	}
}

# Common function to call cluster.exe
# Will try to connect to $dagName or $serverName to execute the command. If it
# fails for certain reasons (such as the server being unavailable) it will
# retry on the other machines in the cluster.
#
# Parameters:
#  $dagName. Name of the DAG. Optional.
#  $serverName. Name of the server to use. Optional.
#   One of $dagName or $serverName MUST be supplied.
#  $clusterCommand. The command to execute.
#
# Returns:
#  @( $errorCode, $textOutput )

function Call-ClusterExe(
	[string] $dagName,
	[string] $serverName,
	[string] $clusterCommand
)
{
	log-verbose ($DagCommonLibrary_LocalizedStrings.res_0097 -f $dagName,$serverName,$clusterCommand,"Call-ClusterExe")

	$script:exitCode = -1;
	$script:textOutput = $null;

	$exitCode = -1;
	$textOutput = $null;

	if ( ( ! $dagName ) -and ( ! $serverName ) )
	{
		log-error -stop "Call-ClusterExe was called with neither a dag name nor a server name!"
	}

	$namesToTry = @( Get-ClusterNames -dagName $dagName -serverName $serverName )

	foreach ( $nodeName in $namesToTry )
	{

		# Start a script block so that $error.Clear doesn't affect
		# callers, and ErrorActionPreference is unchanged.
		&{
			# Simply specifying -erroraction is not enough.
			$ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue

			# Method B
			$clusCommand = "cluster.exe /cluster:$nodeName $clusterCommand 2>&1"
			log-verbose ($DagCommonLibrary_LocalizedStrings.res_0098 -f $clusCommand,$nodeName);

			# Run the command.
			$script:textOutput = invoke-expression $clusCommand -erroraction:silentlycontinue

			# We're handling the errors with the exit codes.
			$error.Clear()

			# Save the exit code.
			$script:exitCode = $LastExitCode

			# Superfluous verbose logging.
			log-verbose ($DagCommonLibrary_LocalizedStrings.res_0099 -f $script:exitCode,$script:textOutput,"Call-ClusterExe")
		}

		# Convert from $script scope to locals.
		$exitCode=$script:exitCode
		$textOutput=$script:textOutput

		log-verbose ($DagCommonLibrary_LocalizedStrings.res_0100 -f $exitCode,$textOutput,"Call-ClusterExe")

		if ( $LastExitCode -eq 1722 )
		{
			# 1722 is RPC_S_SERVER_UNAVAILABLE
			log-verbose ($DagCommonLibrary_LocalizedStrings.res_0101 -f $serverName,"Call-ClusterExe")
			continue;
		}
		elseif ( $LastExitCode -eq 1753 )
		{
			# 1753 is EPT_S_NOT_REGISTERED
			log-verbose ($DagCommonLibrary_LocalizedStrings.res_0102 -f $serverName,"Call-ClusterExe")
			continue;
		}
		elseif ( $LastExitCode -eq 1825 )
		{
			# 1825 is RPC_S_SEC_PKG_ERROR
			log-verbose ($DagCommonLibrary_LocalizedStrings.res_0103 -f $serverName,"Call-ClusterExe")
			continue;
		}
		elseif ( $LastExitCode -ne 0 )
		{
			Log-warning ($DagCommonLibrary_LocalizedStrings.res_0104 -f $LastExitCode,"Call-ClusterExe","retry-able")
			break;
		}
		else
		{
			# It returned 0.
			break;
		}
	}

	return @( $exitCode, $textOutput );
}

# Looks up the DAG and returns an array of servers to contact.
# If serverName is specified, list it first.
# If dagName is specified, then add the dag netname first,
# followed by all of the servers in the dag (excluding serverName)
function Get-ClusterNames(
	[string] $dagName,
	[string] $serverName
)
{
	$serversToTry = @( );

	# If they specified a server, use it first.
	if ( $serverName )
	{
		$serversToTry += $serverName;
	}

	if ( $dagName )
	{
		$dag = @( get-databaseavailabilitygroup $dagName -erroraction silentlycontinue );
		if ( ( ! $dag ) -or ( $dag.Length -eq 0 ) )
		{
			log-warning ($DagCommonLibrary_LocalizedStrings.res_0105 -f $dagName,"Get-ClusterNames","get-dag")
		}
		elseif ( $dag.Length -ne 1 )
		{
			log-warning ($DagCommonLibrary_LocalizedStrings.res_0106 -f $dag.Length,$dagName,"Get-ClusterNames","get-dag")
		}
		else
		{
			# If they specified a valid DAG, try the netname before any of
			# the member servers.
			$serversToTry += $dag[0].Name

			foreach ( $server in $dag[0].Servers )
			{
				# Don't add it a second time!
				if ( $server.Name -ne $serverName )
				{
					$serversToTry += $server.Name
				}
			}
		}

	}

	log-verbose ($DagCommonLibrary_LocalizedStrings.res_0107 -f $dagName,$serverName,$serversToTry,"Get-ClusterNames")
	return $serversToTry;
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
	Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0108 -f $sleepSecs)
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
	Log-Error ($DagCommonLibrary_LocalizedStrings.res_0109 -f $failedCommand,$failedMessage) -Stop:$Stop
}

# SIG # Begin signature block
# MIIaXAYJKoZIhvcNAQcCoIIaTTCCGkkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQULx3yKQkA61HHObillb8jOJAu
# EuSgghUmMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# AgphApJKAAAAAAAgMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQTAeFw0xMjAxMDkyMjI1NTlaFw0xMzA0MDkyMjI1NTlaMIGzMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRNT1BSMScwJQYD
# VQQLEx5uQ2lwaGVyIERTRSBFU046QjhFQy0zMEE0LTcxNDQxJTAjBgNVBAMTHE1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQDNY8P3orVH2fk5lFOsa4+meTgh9NFuPRU4FsFzLJmP++A8W+Gi
# TIuyq4u08krHuZNQ0Sb1h1OwenXzRw8XK2csZOtg50qM4ItexFZpNzoqBF0XAci/
# qbfTyaujkaiB4z8tt9jkxHgAeTnP/cdk4iMSeJ4MdidJj5qsDnTjBlVVScjtoXxw
# RWF+snEWHuHBQbnS0jLpUiTzNlAE191vPEnVC9R2QPxlZRN+lVE61f+4mqdXa0sV
# 5TdYH9wADcWu6t+BR/nwUq7mz1qs4BRKpnGUgPr9sCMwYOmDYKwOJB7/EfLnpVVi
# le6ff54p1s1LCnaD8EME3wZJIPJTWkSwUXo5AgMBAAGjggEJMIIBBTAdBgNVHQ4E
# FgQUzRmsYU2UZyv2r5R1WdvwWACDoTowHwYDVR0jBBgwFoAUIzT42VJGcArtQPt2
# +7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3NvZnQu
# Y29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNybDBY
# BggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNVHSUE
# DDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAURzW3zYd3IC8AqfGmd8W
# yJAGaRWaMrnzZBDSivNiKRt5j8pgqbNdxOQWLTWMRJvx9FxhpEGumou6UNFqeWqz
# YoydTVFQC6mO1L+iEMBH1U51UokeP5Zqjy5AFAZy7j9wWZJdhe2DmfqChJ7kAwEh
# E6sn1sxxSeXaHf9vPAlO1Y1m6AzJf+4xFAI3X3tp7Ik+RX8lROcfGtbFGsNK5OHx
# hJjnT/mpmKcYRuyEbOypAwr9fHpSHZxrpKgPmJKkknhcK3jjfbLH2bZwfd9bc1O/
# qtRmUEvwyTuVheXBSmWdJVhQuyBUkXk6GwdcalcorzHHn+fDHe5H/SfXf8903GXF
# PzCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEFBQAwXzETMBEG
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
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0xggSg
# MIIEnAIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQITMwAAAJ0ejSeu
# uPPYOAABAAAAnTAJBgUrDgMCGgUAoIHCMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBS/VskbQ7mSj7oVA6YnLeBm2jNm2TBiBgorBgEEAYI3AgEMMVQwUqAqgCgARABh
# AGcAQwBvAG0AbQBvAG4ATABpAGIAcgBhAHIAeQAuAHAAcwAxoSSAImh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAPgn0
# 4eN1B8sWZFkA/EtusYkT04qHzScu59jQY3dtJ/FiT5tWO6DjlRLv38TitReBB8Uq
# K/XtZwBXQ71aoSUr475Cf7qy9iGRIQDteen68s2ohFIZGfvwVmo63gMaUUvExMmt
# 8A3+oSr+Plhyaeu626Ga5ttWR+ED639dW1LP92CCUH8/XVypb6tj2nnzCssLAIE6
# NTynU9jwRMNi1krKPBcAcMy1iljeueUIsQPHVRI3QPmm08U/Q43bJg3N5cjz25CA
# NnMkxe+EN+gsY0nNWFE7DNG4fmoAgK6ZHcv7JBy0bdgMtfIdB2OxqrwmOxaAmir2
# FYKfnNCyim/GdYoApaGCAh8wggIbBgkqhkiG9w0BCQYxggIMMIICCAIBATCBhTB3
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhN
# aWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECCmECkkoAAAAAACAwCQYFKw4DAhoFAKBd
# MBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTEzMDEw
# ODA4NDY0NlowIwYJKoZIhvcNAQkEMRYEFGHWCBkDbbFjEx3p4k8MAntBu3/cMA0G
# CSqGSIb3DQEBBQUABIIBAF38c8YWlv3ptnJduYWRLZrFVfSV74Lpw+wbOktbOeQ+
# x7TKxJSWfPGFaTBUp+DXQTl6KvmzbRh7iddWJ5biIuUIiaQiHjYK+luEncz9Vws3
# SzMTcSjLSkRTSrOgrQOsdp5p8KDtD7vSdMcyTTNb/p2GMbF3ES3OfkAqsxjktQVy
# 1jKebwEldfY+7dZtaeC+YwvYhJwblEKYyy1RzGLaTVkb4sVUx1FJINWp4x/JmU0H
# 9d/cElcIgcA4YeUxW9ZARN7O7MFgYiAKonCENoKshAVrRQoIb4jpAmJ1drUnNCSC
# uFQAN/rwrk+2UpG9HWai0OYfBkvzgMJipqhwmAUSVs4=
# SIG # End signature block
