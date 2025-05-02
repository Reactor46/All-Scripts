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
	}
    return $moveSuccessful
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
# MIIawQYJKoZIhvcNAQcCoIIasjCCGq4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU2XmH8rG2gIkC04CFDPYYsGrD
# WROgghWCMIIEwzCCA6ugAwIBAgITMwAAAHD0GL8jIfxQnQAAAAAAcDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTUwMzIwMTczMjAy
# WhcNMTYwNjIwMTczMjAyWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkY1MjgtMzc3Ny04QTc2MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAoxTZ7xygeRG9
# LZoEnSM0gqVCHSsA0dIbMSnIKivzLfRui93iG/gT9MBfcFOv5zMPdEoHFGzcKAO4
# Kgp4xG4gjguAb1Z7k/RxT8LTq8bsLa6V0GNnsGSmNAMM44quKFICmTX5PGTbKzJ3
# wjTuUh5flwZ0CX/wovfVkercYttThkdujAFb4iV7ePw9coMie1mToq+TyRgu5/YK
# VA6YDWUGV3eTka+Ur4S+uG+thPT7FeKT4thINnVZMgENcXYAlUlpbNTGNjpaMNDA
# ynOJ5pT2Ix4SYFEACMHe2j9IhO21r9TTmjiVqbqjWLV4aEa/D4xjcb46Q0NZEPBK
# unvW5QYT3QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFG3P87iErvfMdr24e6w9l2GB
# dCsnMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAF46KvVn9AUwKt7hue9n/Cr/bnIpn558xxPDo+WOPATpJhVN
# 98JnglwKW8UK7lXwoy2Ooh2isywt0BHimioB0TAmZ6GmbokxHG7dxHFU8Ami3cHW
# NnPADP9VCGv8oZT9XSwnIezRIwbcBCzvuQLbA7tHcxgK632ZzV8G4Ij3ipPFEhEb
# 81KVo3Kg0ljZwyzia3931GNT6oK4L0dkKJjHgzvxayhh+AqIgkVSkumDJklct848
# mn+voFGTxby6y9ErtbuQGQqmp2p++P0VfkZEh6UG1PxKcDjG6LVK9NuuL+xDyYmi
# KMVV2cG6W6pgu6W7+dUCjg4PbcI1cMCo7A2hsrgwggTsMIID1KADAgECAhMzAAAA
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
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBKkwggSl
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAAAymzVMhI1xOFV
# AAEAAADKMAkGBSsOAwIaBQCggcIwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFJhR
# dXQJjKpp9L83rbOK5a2yg5q4MGIGCisGAQQBgjcCAQwxVDBSoCqAKABEAGEAZwBD
# AG8AbQBtAG8AbgBMAGkAYgByAGEAcgB5AC4AcABzADGhJIAiaHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQBSsdoEd1if
# ABi+HLKjz8AMu5uIM7wUYoQqF7u97bLyXM1YrxH0kii9nwZdXe7lE4Sq75nq4hW7
# 6iJqGOETHrV7aBN1gAPrPr7tu9Q3cSlu+C6PLUF4VzlaNxLBGawowFQKV3g8oUJL
# dJR1oVHiaUnM2+y29zl5UVfLAtsCzC/UoI010vbkZYZABBQ+fP1A4z7Awe8981G3
# 5AHGpC2RigsH8JjEkKeCFhbxDWNxK3zmVcCDqPIG/2OnXkODnqGwjIGzJNk89ety
# ptywTVyoZuipTkolVIEFdrCtH8pKdRIPVGY8+zOmwBsjSu9Emf4Ssu4YUxRHz0UD
# 94D2s9HaHlKnoYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jv
# c29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAHD0GL8jIfxQnQAAAAAAcDAJBgUrDgMC
# GgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MTUwNDEwMDI1NzM1WjAjBgkqhkiG9w0BCQQxFgQUSjO0UdTe8Hu78bP92ToHc4HG
# vJ0wDQYJKoZIhvcNAQEFBQAEggEAkG9Ul0mJes8AtYtN6M/Z3PTZ0Khpiy07JUF6
# ODJ0kw71P/AZ+MjZxE6tCtYjH3jNB+lwhDOZ5QP11ELhRtjUjdYh0MzzhalCur5h
# 9w3O36ReRVKPKi9Lqe7kg0PSIkM4YD3pibtw+gXY/TNswU1uRwbpzQx5HRBaKn1L
# RPdL0jjka8ne6OQJjk7NQJF6jGCONpMPqAXeuzVMeDGmi7sfnETFxxNQ86sHbpC/
# zI15G+qyJwHdhK67jESYCNpxh0mW03h1S7BtVbmgk760pNN69k01w1bjhqDD6Sjf
# 8TJ8y40Ggz8W5gdTsG75MZB+wd31Z6LQ/chqWDXlxV10Sjusag==
# SIG # End signature block
