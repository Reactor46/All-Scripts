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
# -Any target replicated mailbox databases are suspended or Misconfigured.
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

	    $dag = Get-DatabaseAvailabilityGroup $mbxServer.DatabaseAvailabilityGroup;
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
            $dag = Get-DatabaseAvailabilityGroup $mbxServer.DatabaseAvailabilityGroup -status;
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
	
	if ( $numDatabases -gt 0 )
	{
        $copyStatuses = Get-MailboxDatabaseCopyStatus -Server $serverName
		foreach ( $database in $databases )
		{
            Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0111 -f $database.Name,"Get-CriticalMailboxResources",$database.AutoDagExcludeFromMonitoring);
            
			if ( $database.ReplicationType -eq 'Remote' )
			{
                if ($database.AutoDagExcludeFromMonitoring -eq $true)
                {                    
                    Continue
                }
            
				# Make sure replicated databases aren't hosted on the server.
				# And if they aren't hosted, they need to be suspended.
                $status = $copyStatuses | ?{$_.DatabaseName -eq $database.Name} | Select-Object -First 1
                if(!$status)
                {
				    $status = get-mailboxdatabasecopystatus $database\$serverName
                }

				Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0018 -f $status.ActiveCopy,$status.Status,$status.ActivationSuspended,$database,"Get-CriticalMailboxResources");

				if ( $status.ActiveCopy )
				{
					Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0019 -f $status.ActiveCopy,$status.Status,$status.ActivationSuspended,$database,"Get-CriticalMailboxResources");
					write-output "(Database='$($database.Name)', Reason='Copy is active')";
				}
                else
		        {
				    # Check if the active is copy is mounted
			        $activeServerName = $status.ActiveDatabaseCopy
                    if ($activeServerName)
                    {
                        $activeStatus = Get-MailboxDatabaseCopyStatus $database\$activeServerName
			            if ($activeStatus.Status -ine 'Mounted')
                        {
			                Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0110 -f $database,$activeServerName,"Get-CriticalMailboxResources");
			                write-output "(Database='$($database.Name)', Reason='Active is dismounted on $activeServerName')";
		                }
                    }
		        }
				
                if ( $status.Status -eq 'ServiceDown' )
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
				# O15 372945: If we add single-copy check for non-replicated databases, then we break new DAG deployment
				# scenario. Consider: If we have only 1 copy of a DB and need to reinstall, we should not block due to 
				# this check. It would be great if deployment process could mark up the DAG state as "not-yet-fully-deployed"
				# so that scripts like this could be smarter and safer.
				#
				else
                {
                    $result = IsDatabaseReadyForCopyRemoval -DatabaseGuid $database.Guid -AtleastNCopies $AtLeastNCriticalCopies -Target $serverName
                    if (!$result.Success)
                    {
					    Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0023 -f $database,$serverName,$database,$serverName)
					    write-output "(Database='$($database.Name)', Reason='$($result.Message)')";
				    }
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

function IsDatabaseReadyForCopyRemoval ([Guid] $databaseGuid, [int] $AtLeastNCopies, [string] $Target)
{
	[string]$version = [Microsoft.Exchange.Diagnostics.BuildVersionConstants]::RegistryVersionSubKey;
	$setupRegistryPath = Get-ItemProperty -path "HKLM:SOFTWARE\Microsoft\ExchangeServer\$version\Setup"
    $replPath = Join-Path $setupRegistryPath.MsiInstallPath "Bin\Microsoft.Exchange.Cluster.Replay.dll"	
    [System.Reflection.Assembly]::LoadFrom($replPath) | Out-Null
    
    $dbHealthValidater = New-Object -TypeName Microsoft.Exchange.Cluster.Replay.DatabaseHealthValidationRunner -ArgumentList $Target
    $dbHealthValidater.Initialize()
    $validaterType = $dbHealthValidater.GetType()
    $sickDatabases = $null
    [hashtable]$result = @{}
    $result.Success = $true
    $result.Message = ""    
    
	# Check if the new RunDatabaseServerBeginMaintenanceChecks method exists, in order to handle cross-version issues
	$maintenanceMember = $validaterType.GetMember("RunDatabaseServerBeginMaintenanceChecks")
	
	if ($maintenanceMember)
	{
		if ($AtLeastNCopies)
		{
			$sickDatabases = $dbHealthValidater.RunDatabaseServerBeginMaintenanceChecks($databaseGuid, $AtLeastNCopies)
		}
		else
		{
			$sickDatabases = $dbHealthValidater.RunDatabaseServerBeginMaintenanceChecks($databaseGuid)
		}
	}
	else
	{	
		# Only need this block to handle cross-version issues
		#
		# Lowest build deployed in production is 15.00.0980.008, and these parameters ($AtLeastNCopies) were added in 15.00.0879.000
		#
        if ($AtLeastNCopies)
        {
            $sickDatabases = $dbHealthValidater.RunDatabaseRedundancyChecksForCopyRemoval($false, $databaseGuid, $false, $true, $AtLeastNCopies)
        }
        else
        {
            $sickDatabases = $dbHealthValidater.RunDatabaseRedundancyChecksForCopyRemoval($false, $databaseGuid, $false, $true)
        }
    }
	
    foreach ($sickDatabase in $sickDatabases)
    {
        if ($sickDatabase.IsValidationSuccessful -eq $false)
        {
            $result.Success = $false
            $result.Message = $sickDatabase.ErrorMessageWithoutFullStatus            
            break
        }
    }
    
    return $result
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
function Move-CriticalMailboxResources(
	[string] $Server, 
	[string] $MoveComment = "DagCommonLibrary:Move-CriticalMailboxResources",
	[string] $Force = 'false')
{
    $msg = ""
	$moveActCmd =
	{
		# 1. Move any mailbox database masters to other servers.
		$Error.Clear()
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0043 -f $Server,"Move-CriticalMailboxResources")
		if ($Force -eq 'true')
		{
   			Move-ActiveMailboxDatabase -Server $Server -Confirm:$false -MoveComment $MoveComment -SkipMoveSuppressionChecks
		}
		else
		{
			Move-ActiveMailboxDatabase -Server $Server -Confirm:$false -MoveComment $MoveComment
		}

		foreach ($er in $Error)
		{
			Log-Verbose "[Error] $er"            
		}
	}
	
    $msg += TryExecute-ScriptBlockReturnErrorMessage -runCommand $moveActCmd -throwOnError $false -silentonerrors $true

	$moveResCmd = 
	{		
		# 2. Move PAM if the server is a PAM.
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0044 -f $Server,"Move-CriticalMailboxResources")
		Move-DagActiveManager -Server $Server		
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0045 -f $Server,"Move-CriticalMailboxResources")
	}
	
	$errorCmd = { Log-Error ($DagCommonLibrary_LocalizedStrings.res_0046 -f $Server,"Move-CriticalMailboxResources") }
	$msg += TryExecute-ScriptBlockReturnErrorMessage -runCommand $moveResCmd -cleanupCommand $errorCmd -throwOnError $true
    
    Return $msg
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

# This is only called by obsolete but still existing switchover workflows. 
# Please don't add more calls without checking with exhacore.
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
    
	$Error.Clear()
	try
	{
		if ($targetdbs)
		{
			foreach ($targetdb in $targetdbs)
			{
				Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0059 -f $targetdb.MailboxServer,$db,$false,"Move-DagMasterCopy","Move-ActiveMailboxDatabase","-Identity","-ActivateOnServer","-SkipClientExperienceChecks","-Confirm")
				if ( $DagScriptTesting )
				{
					$moveCmd = { write-host ($DagCommonLibrary_LocalizedStrings.res_0060 -f $db,$targetdb,$false,"Move-DagMasterCopy", "Move-ActiveMailboxDatabase", "-Identity","-ActivateOnServer","-SkipClientExperienceChecks","-Confirm") }
				}
				else
				{
					$moveCmd = { 
						Move-ActiveMailboxDatabase -Identity $db -ActivateOnServer $targetdb.MailboxServer `
							-SkipClientExperienceChecks -MoveComment "Move-DagMasterCopy" -Confirm:$false 
						}
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
	catch
	{
		# We need to log all new error messages
		if ($Error.Count -gt 0 )
		{
			[string] $error_message = ""
			foreach ( $err in $Error )
			{
				$error_message += $err.ToString() + "`n"
			}
			$Error.Clear()
			write-Error $error_message -ErrorAction:Continue
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
		# Does this make sense? Shouldn't this logic exist in PAM?
		# Regardless, I think it is dead so won't try to do much with it.
		# I collapsed the 2 move commands into one, since there's no point avoiding -SkipClientExperienceChecks 
		# when you are just going to use it if the move fails.
		Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0089 -f $db.Name,$currentActive,$currentDb,$desiredActiveServer,$false,"Rollback-DatabaseOnFailedMove","Move-ActiveMailboxDatabase","-Identity","-ActivateOnServer","-Confirm")

		if ( $DagScriptTesting )
		{
			$moveCmd = { write-host ($DagCommonLibrary_LocalizedStrings.res_0090 -f $currentDb,$desiredActiveServer,$false,"Rollback-DatabaseOnFailedMove", "Move-ActiveMailboxDatabase","-Identity","-ActivateOnServer","-Confirm") }
		}
		else
		{
			$moveCmd = 
			{ 
				Move-ActiveMailboxDatabase -Identity $currentDb -ActivateOnServer $desiredActiveServer `
					-Confirm:$false -MoveComment "Rollback-DatabaseOnFailedMove" -SkipMoveSuppressionChecks -SkipClientExperienceChecks
			}
		}
		if (TryExecute-ScriptBlock $moveCmd)
		{
			Log-Verbose ($DagCommonLibrary_LocalizedStrings.res_0091 -f $db.Name,$currentActive,"Rollback-DatabaseOnFailedMove")
			return
		}
		Log-Error ($DagCommonLibrary_LocalizedStrings.res_0092 -f $db.Name,$currentActive,"Rollback-DatabaseOnFailedMove")
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
				else
				{
					Log-Verbose $_
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
function TryExecute-ScriptBlockReturnErrorMessage ([ScriptBlock]$runCommand, [ScriptBlock]$cleanupCommand={}, [bool]$throwOnError=$false, [bool]$silentOnErrors=$false)
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
				
				$strToLog = ""
                if($_.Exception)
                {
                    $strToLog += "[" + $_.Exception.GetType().Name + "] "
                }
                $strToLog += $_.ToString()
				
				if (!$silentOnErrors)
				{
					Log-ErrorRecord $strToLog
				}
				else
				{
					Log-Verbose $strToLog
                    Write-Output $strToLog
				}
				
				# Run the cleanup scriptblock
				$ignoredObjects = @(&$cleanupCommand)
			}
			
			if ($throwOnError)
			{
				throw
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
	Write-Debug "$timeStamp $msg"
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

function Is-GlobalVariableDefined( [string] $varName )
{
	return get-variable -scope global -name $varName -ErrorAction SilentlyContinue;
}

function Is-VariableDefined( [string] $varName )
{
	test-path variable:$varName;
}

function Get-DatacenterScriptDirectory
{
	#[string]$filePath = $RoleDatacenterPath
	[string]$filePath = $null
	if ( Is-GlobalVariableDefined "RoleDatacenterPath")
	{
		$filePath = $RoleDatacenterPath
	}

	if (!$filePath)
	{
		# For testing scenarios, $RoleDatacenterPath is not defined
		
		# Try to find the Datacenter directory
		[string]$version = [Microsoft.Exchange.Diagnostics.BuildVersionConstants]::RegistryVersionSubKey
		$installPath = (Get-ItemProperty -path "HKLM:SOFTWARE\Microsoft\ExchangeServer\$version\Setup").MsiInstallPath.TrimEnd("\")
		$datacenterPath = $installPath + '\Datacenter'
		if (Test-Path $datacenterPath)
		{
			$filePath = $datacenterPath
		}
		else
		{
			# Get the current script name. The method is different if the script is
			# executed or if it is dot-sourced, so do both.
			$thisScriptName = $myinvocation.scriptname
			if ( ! $thisScriptName )
			{
				$thisScriptName = $myinvocation.MyCommand.Path
			}

			$filePath = Split-Path -Parent $thisScriptName;
		}
	}
	return $filePath
}

function Initialize-RoleDatacenterPathVariable
{
	if (!(Is-GlobalVariableDefined "RoleDatacenterPath"))
	{
		$global:RoleDatacenterPath = Get-DatacenterScriptDirectory
	}
}

Initialize-RoleDatacenterPathVariable

# If DatacenterLogHelper is available dot source the file
if (Test-Path "$RoledatacenterPath\DatacenterLogHelper.ps1")
{
	. "$RoledatacenterPath\DatacenterLogHelper.ps1"
}
elseif (Test-Path ".\DatacenterLogHelper.ps1")
{
	. ".\DatacenterLogHelper.ps1"
}

# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUJiYuE7STgLTsQzMVTW09v3he
# +dKgghhkMIIEwzCCA6ugAwIBAgITMwAAAJmqxYGfjKJ9igAAAAAAmTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI4
# WhcNMTcwNjMwMTkyMTI4WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# Ojk4RkQtQzYxRS1FNjQxMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAipCth86FRu1y
# rtsPu2NLSV7nv6A/oVAcvGrG7VQwRn+wGlXrBf4nyiybHGn9MxuB9u4EMvH8s75d
# kt73WT7lGIT1yCIh9VC9ds1iWfmxHZtYutUOM92+a22ukQW00T8U2yowZ6Gav4Q7
# +9M1UrPniZXDwM3Wqm0wkklmwfgEEm+yyCbMkNRFSCG9PIzZqm6CuBvdji9nMvfu
# TlqxaWbaFgVRaglhz+/eLJT1e45AsGni9XkjKL6VJrabxRAYzEMw4qSWshoHsEh2
# PD1iuKjLvYspWv4EBCQPPIOpGYOxpMWRq0t/gqC+oJnXgHw6D5fZ2Ccqmu4/u3cN
# /aAt+9uw4wIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFHbWEvi6BVbwsceywvljICto
# twQRMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBABbNYMMt3JjfMAntjQhHrOz4aUk970f/hJw1jLfYspFpq+Gk
# W3jMkUu3Gev/PjRlr/rDseFIMXEq2tEf/yp72el6cglFB1/cbfDcdimLQD6WPZQy
# AfrpEccCLaouf7mz9DGQ0b9C+ha93XZonTwPqWmp5dc+YiTpeAKc1vao0+ru/fuZ
# ROex8Zd99r6eoZx0tUIxaA5sTWMW6Y+05vZN3Ok8/+hwqMlwgNR/NnVAOg2isk9w
# ox9S1oyY9aRza1jI46fbmC88z944ECfLr9gja3UKRMkB3P246ltsiH1fz0kFAq/l
# 2eurmfoEnhg8n3OHY5a/Zzo0+W9s1ylfUecoZ4UwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBK4wggSqAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBwjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUTt+M3d5J311bMQsTUheDwXRvXmUwYgYKKwYB
# BAGCNwIBDDFUMFKgKoAoAEQAYQBnAEMAbwBtAG0AbwBuAEwAaQBiAHIAYQByAHkA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBAAkgDM9vUJ8Ed9vBdz8Tsb5+gC2kPlUAq1KMt7swAESf
# L+rUKNpoHzNFsD4P/ED/5idM4IC3g7zaLJHptoEKpXTwy/AEN6ONs913obEkN6WE
# j3YVLhzjvgj85tbGtbq8dEoyD0cb3IOzYHIuc29PlS1XiuieudlnWFd8FA/zuna3
# 2Zb/cVsxPCy4oJnMx6El1VtGP/FNGFHg4Vuv8xb5NDtsyNsjV5P7vHv0zrMRtNmp
# CTPbyhbgl39lgYM6LDv8497rRTHlTLjLuQN2X90i2uJl4cK4Bvaze+RzW4DE61KM
# z8jAj2H0uCYpjRqC7PHo6mY8fBw58q5s9UBfTu/XWmehggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAA
# marFgZ+Mon2KAAAAAACZMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMxODQ1MDBaMCMGCSqGSIb3DQEJ
# BDEWBBR2ifnpnXRJYL8edPRhNFUViYzEhjANBgkqhkiG9w0BAQUFAASCAQBnyUxd
# snSMgZF+xtmf9JZ+h0LchhkRtPi0kdb6YWjm96+/dFRcSAEDak2NvM0lukHTtEQg
# +lUzZuXajnwgYlxfrgF3xCHMLUVYKqcwWw0EGSmLilGcTqNAd9GEAaCwh9+PjXzp
# w7CEg0tE89HsIv4n8a9ToBGjvYL2hjbXePyTVNzdvdGJqZ+Z6oxhgq5ApN1SJDsk
# QPsOZxypxfGOdBtiCjO7F7c/4C1h90+FIFZjbsQ/iWMdaROfTS+g2EAE6AYXKF4C
# aM7UW5+0qz9UC1nEVDWr3VFT4uDSuJhiS7XvZ4SZRuPWJMyL6n7TPjpFG/DPUTYI
# F/xTsCxmfKy2qRbQ
# SIG # End signature block
