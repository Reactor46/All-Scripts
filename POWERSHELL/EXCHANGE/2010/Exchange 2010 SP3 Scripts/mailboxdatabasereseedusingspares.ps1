<#
.EXTERNALHELP MailboxDatabaseReseedUsingSpares-help.xml
#>

#
# Copyright (c) 2010 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

# This script is used to validate the safety of the environment, before proceeding to swap failed database copy
# to a spare disk and reseed.
#
# Mailbox Database Reseed
#
# If both $MailboxServerName and $MailboxDatabaseName are provided, it tries to reseed the database on the server.
#
param(
	[string]	$MailboxServerName,
	[string]	$MailboxDatabaseName,
	[string]	$NotifyEmail,
	[switch]	$SingleActiveCopyOK=$false,
	[switch]	$ReseedConfirm=$true,
	[switch]	$CatalogOnly=$false,
	[switch]	$Verbose=$false,
	[int]		$SpareDiskToUse=$null,
	# Test Hooks
	[switch]	$SkipSpareCheck=$false,
	[switch]	$WhatIf=$false,
	[int]		$SpareDriveOverride=$null,
	[int]		$SpareDriveReserveTimeOverride=(3*60*60),
	[int]		$WaitBeforeReseed,
	[int]		$ReseedCount=5)

#
# Globals
#
[int]$script:NumberOfHealthyPassiveCopiesRequired = 1
[int]$script:RequiredLogsCopied = 3
[int]$script:RequiredLogsReplayed = 3
[int]$script:MaxWorkflowReseedCount = 1
[int]$script:ReseedIntervalInDays = 2
[int]$script:SleepSeconds = 60
[string]$script:VolumeHistoryFile = "VolumeHistory.xml"
[string]$script:DatabaseHistoryFile = "DatabaseHistory.xml"
[string]$script:DatabasePath = ""
[string]$script:LockPath = "C:\LocalFiles\ReseedLocks"
[switch]$script:IsOperator = $true

#
# Helper function
#

Import-LocalizedData -BindingVariable MailboxDatabaseReseedUsingSpares_LocalizedStrings -FileName MailboxDatabaseReseedUsingSpares.strings.psd1
Set-StrictMode -Version 2.0

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ea SilentlyContinue
Add-PSSnapin Microsoft.Exchange.Management.Powershell.Setup -ea SilentlyContinue
Add-PSSnapin Microsoft.Exchange.Management.Powershell.Support -ea SilentlyContinue
Add-PSSnapin Microsoft.Exchange.Management.Powershell.CentralAdmin -ea SilentlyContinue

# Get the current script name. The method is different if the script is
# executed or if it is dot-sourced, so do both.
$thisScriptName = $myinvocation.scriptname
if ( ! $thisScriptName )
{
	$thisScriptName = $myinvocation.MyCommand.Path
}

# Many of the script libraries already use $DagScriptTesting
if ( $whatif )
{
	$DagScriptTesting = $true;
}

# If RoleDatacenterPath is not defined use current directory
if ( !( Test-Path variable:\RoleDatacenterPath ) )
{
	$RoleDatacenterPath = split-path $myInvocation.MyCommand.Path;
}

# Load some of the common functions.
# The common lib doesn't use clean practices so we have to avoid strict mode
Set-StrictMode -Off

. "$RoleDatacenterPath\DatacenterDiskCommonLibrary.ps1";

Set-StrictMode -Version 2.0

$CopyStatusType = [Microsoft.Exchange.Management.SystemConfigurationTasks.CopyStatus]

$validEventMap = @{
	#<data name="DatabaseNotPresentAfterReplay">
	#	<stringvalue>The database file wasn't found after log replay. The copy will be set to failed. Database: '%1'. File Path: '%2'.</stringvalue>
	#	<eventmessageid>2106</eventmessageid>
	#Reseed is needed.
	"2106" = $null;
	
	#<data name="IncrementalReseedPrereqError">
	#	<stringvalue>Incremental seeding of database %1 was not initiated due to an unsatisfied prerequisite. The copy must be reseeded. Error: %2</stringvalue>
	#	<eventmessageid>3147</eventmessageid>
	#Reseed is needed.
	"3147" = $null;	
	
	#<data name="LogFileGapFound">
	#	<stringvalue>The log file %2 for %1 is missing on the production copy. Continuous replication for this database is blocked. If you removed the log file, please replace it. If the log is lost, the passive copy will need to be reseeded using the Update-MailboxDatabaseCopy cmdlet in the Exchange Management Shell.</stringvalue>
	#	<eventmessageid>2059</eventmessageid>
	#Reseed is needed. And a bug.
	"2059" = $null;

	#<data name="SeedInstancePrepareFailed">
	#	<stringvalue>[Seed Manager] Seed request for database '%1' encountered an error while running prerequisite checks. The error that occurred was: %2</stringvalue>
	#	<eventmessageid>4025</eventmessageid>
	#Reseed is needed.
	"4025" = $null;

	#<data name="SeedInstanceInProgressFailed">
	#	<stringvalue>[Seed Manager] Seed request for database '%1' encountered an error during seeding. The error that occurred was: %2</stringvalue>
	#	<eventmessageid>4026</eventmessageid>
	#Reseed is needed.
	"4026" = $null;

	#<data name="SeedInstanceCancelled">
	#	<stringvalue>Seed request for database '%1' has been cancelled.</stringvalue>
	#	<eventmessageid>4027</eventmessageid>
	#Reseed is needed.
	"4027" = $null;

	#<data name="SeedInstanceAnotherError">
	#	<stringvalue>[Seed Manager] Seed request for database '%1' encountered another error. The error that occurred was: %2</stringvalue>
	#	<eventmessageid>4083</eventmessageid>
	#Reseed is needed.
	"4083" = $null;

	#<data name="FileCheckError">
	#	<stringvalue>The Microsoft Exchange Replication Service encountered an error while inspecting the logs and database for %1 on startup. The specific error code returned is: %2.</stringvalue>
	#	<eventmessageid>2070</eventmessageid>
	#Reseed for these exceptions only:  TBD  - Shuab to propose some:
	"2070" = @( `
		'^FileCheckLogfileSignatureException$', `
		'^FileCheckDatabaseLogfileSignatureException$', `
		'^FileCheckEDBMissingException$', `
		'^FileCheckLogfileGenerationException$', `
		'^FileCheckLogfileCreationTimeException$', `
		'^FileCheckLogfileMissingException$', `
		'^FileCheckRequiredLogfileGapException$', `
		'^IsamMishandlingException$', `
		'^IsamStateException$');

	#<data name="IncrementalReseedFailedError">
	#	<stringvalue>Incremental seeding on database %1 encountered an error and a full reseed is required. Reason: %2</stringvalue>
	#	<eventmessageid>3145</eventmessageid>
	#Reseed for log missing case only.
	"3145" = @( '^ReseedCheckMissingLogfileException$' );

	#<data name="IsamException">
	#	<stringvalue>The Microsoft Exchange Replication Service encountered an unexpected Extensible Storage Engine (ESE) exception in database '%1'. The ESE exception is %2 (%3).</stringvalue>
	#	<eventmessageid>2097</eventmessageid>
	"2097" = @( '^IsamMishandlingException$', '^IsamStateException$' );

	#<data name="ConfigurationCheckerFailedPermanent">
	#	<stringvalue>The Microsoft Exchange Replication Service encountered an error while attempting to start a replication instance for %1. The error is not transient, and administrator intervention may be needed. The specific error returned is: %2.</stringvalue>
	#	<eventmessageid>2101</eventmessageid>
	"2101" = @( '^IsamMishandlingException$', '^IsamStateException$' );
	
	#<data name="LogReplayMapiException">
	#	<stringvalue>The Microsoft Exchange Replication Service encountered an unexpected exception in log replay for database '%1'. The exception is: %2</stringvalue>
	#	<eventmessageid>4057</eventmessageid>
	"4057" = @( '^IsamMishandlingException$', '^IsamStateException$' ) }

function ValidateEvent([string]$eventId,[Microsoft.Exchange.Rpc.Cluster.ExtendedErrorInfo]$errorException)
{
	if ( $validEventMap.Contains($eventId) )
	{
		if ( $validEventMap[$eventId] )
		{
			if ( $errorException )
			{				
				foreach($expression in $validEventMap[$eventId])
				{
					if ( $errorException.FailureException.GetType().Name -match $expression )
					{
						return $true
					}
					else
					{
						$innerException = $errorException.FailureException.InnerException;
						while ( $innerException )
						{
							if ( $innerException.GetType().Name -match $expression )
							{
								return $true
							}
							else
							{
								$innerException = $innerException.InnerException
							}							
						}
					}
				}
				return $false
			}
			else
			{
				return $false
			}
		}
		else
		{
			return $true
		}		
	}
	else 
	{ 
		return $false 
	}
}

# function to execute reseed checks and proceed to reseed
#
function Execute-Reseeding
{
    Log-Verbose "Execute-Reseeding -MailboxServer $MailboxServerName -Database $MailboxDatabaseName -SingleActiveCopyOK $SingleActiveCopyOK -ReseedConfirm $ReseedConfirm -CatalogOnly $CatalogOnly -Verbose $Verbose"
	
	$global:WorkflowUrgent = $true
	
    $problemcopy = @()
    $server = Get-ExchangeServer | where {$_.Name -ieq $MailboxServerName}
    $problemdb = Get-MailboxDatabase -Status | where {$_.Name -ieq $MailboxDatabaseName}
	
	# If 1)the MailboxServer value is not specified or invalid or 2)the Database value is specified but invalid, we log an error
    #
    if ((!$server) -or (!$MailboxDatabaseName) -or (!$problemdb))
    {
        #Log-Error "No valid database or target mailbox server name for reseeding was provided. Specified database name: '$MailboxDatabaseName', mailbox server name: '$MailboxServerName'"
		Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0001 -f $MailboxDatabaseName,$MailboxServerName)

		$global:WorkflowUrgent = $false
        Return
    }
	
    # Checking whether the specified database copy is Failed or Suspended 
    #    
	#Log-Verbose "Get-MailboxDatabaseCopyStatus $problemdb\$server"
	Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0049 -f $problemdb,$server,"Get-MailboxDatabaseCopyStatus")
    $copyStatus = Get-MailboxDatabaseCopyStatus -ExtendedErrorInfo $problemdb\$server               
            
    # For CatalogOnly, we'll process for any state
    #
    if (($CatalogOnly) `
	-or ($copyStatus.Status -eq $CopyStatusType::Failed) `
	-or ($copyStatus.Status -eq $CopyStatusType::Suspended) `
	-or ($copyStatus.Status -eq $CopyStatusType::FailedAndSuspended))
    {
		if ( (!$CatalogOnly) -and `
			( ($copyStatus.Status -eq $CopyStatusType::Failed) -or `
			($copyStatus.Status -eq $CopyStatusType::FailedAndSuspended) ) )
		{
			if ( ValidateEvent $copyStatus.ErrorEventId $copyStatus.ExtendedErrorInfo )
			{
				$problemcopy += $copyStatus
			}
			else
			{
				#Log-Error "Skipping reseed of $problemdb\$server, because it is not applicable to the failure experienced by the database copy."
				Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0003 -f $problemdb,$server)
				$global:WorkflowUrgent = $false
				Return
			}
		}
		else
		{
			$problemcopy += $copyStatus
		}
    }
    else
    {
        #Log-Warning "Mailbox database '$MailboxDatabaseName' is neither 'Failed' nor 'Suspended' status on server $MailboxServerName. Reseeding skipped."
		Log-Warning ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0004 -f $MailboxDatabaseName,$MailboxServerName)
		$global:WorkflowUrgent = $false
        Return
    }
	
	# Verify if Database is sufficiently healthy for reseed if SingleActiveCopyOK=$false
	#
	if (!$SingleActiveCopyOK)
	{
		if (Get-DatabaseHealth $problemdb)
		{
			Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0015 -f $MailboxDatabaseName)
		}
		else
		{
			Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0041 -f $MailboxDatabaseName,$script:NumberOfHealthyPassiveCopiesRequired)
			$global:WorkflowUrgent = $true
			Return
		}
	}
	else
	{
		Log-Warning ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0042 -f $MailboxDatabaseName,"-SingleActiveCopyOk")
	}
	
	$count = Get-ReseedCount	
	# Verify number of times reseed script ran
	if ($script:IsOperator)
	{
		if ($count -ge $ReseedCount)
		{
			Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0043 -f $ReseedCount,$script:ReseedIntervalInDays,"-ReseedCount")
			$global:WorkflowUrgent = $true	
			Return
		}
	}
	else
	{
		if ($count -ge $script:MaxWorkflowReseedCount)
		{	
			Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0044 -f $script:ReseedIntervalInDays)
			$global:WorkflowUrgent = $true	
			Return
		}
	}
	
	$sourceDisk = 0
	$spareDisk = 0
	# Verify Spare Allocation
	if (!$SkipSpareCheck)
	{
		# Dot source the datacenter disk common stuff
		if (Test-Path "$RoleDatacenterPath\DatacenterDiskCommonLibrary.ps1")
		{
			Set-StrictMode -Off
			Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0045 -f $RoleDatacenterPath,"Dot-Source")
			. "$RoleDatacenterPath\DatacenterDiskCommonLibrary.ps1"
			Set-StrictMode -Version 2.0
		}
		else
		{
			Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0019 -f $RoleDatacenterPath)
			$global:WorkflowUrgent = $true
			Return
		}
		$sourceDisk = Get-DatabaseDisk
		if ($sourceDisk -eq 0)
		{
			Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0020 -f $MailboxDatabaseName)
			$global:WorkflowUrgent = $true
			Return
		}
		
		# Verify Spare Allocation
		$spareHash = Reserve-SpareVolume -mailboxDatabaseName $MailboxDatabaseName
		$spareDisk = Get-DiskFromVolume -Volume $($spareHash.reserveVolume)
		$spareMount = $spareHash.reserveMount
		if ($spareDisk -le 0)
		{
			Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0046 )
			$global:WorkflowUrgent = $true
			Return
		}
	}
	else
	{	
		if(!$ReseedConfirm)
		{	
			Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0047 -f "-SkipSpareCheck")
			$global:WorkflowUrgent = $false
			Return
		}
	}

	# Verify whether to proceed to Reseed or return with summary
	#
	if ($ReseedConfirm)
	{
        ## Correct this one
		Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0048 -f
                    $spareDisk,$MailboxDatabaseName,$spareDisk,$MailboxServerName,
                    "pre-checks","-ReseedConfirm","pre-checks",
        	        "MailboxDatabaseReseedUsingSpares.ps1 -ReseedConfirm:`$false -MailboxServerName $MailboxServerName -MailboxDatabaseName $MailboxDatabaseName",
                    "Get-Help MailboxDatabaseReseedUsingSpares.ps1")
		
		$global:WorkflowUrgent = $false
		Return
	}

	# (This used to be in Swap-DatabaseDiskToSpare)
	Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0049 -f $MailboxDatabaseName,$MailboxServerName,"Remove-MailboxDatabaseCopy")		
	Remove-MailboxDatabaseCopy "$MailboxDatabaseName\$MailboxServerName" -Confirm:$false

	# Spare manipulation. This will either allocate a spare (if there isn't a database disk),
	# or it will swap the database disk with a spare disk.
	if (Swap-DatabaseDiskToSpare -SourceDisk $sourceDisk -SpareDisk $spareDisk -SpareMount $spareMount -mailboxDatabaseName $MailboxDatabaseName )
	{
		Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0023 -f $MailboxDatabaseName,$sourceDisk,$spareDisk)
	}
	else
	{
		Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0024 -f $sourceDisk,$spareDisk,$spareMount)

		$global:WorkflowUrgent = $true
		Return
	}

	# (This used to be in Swap-DatabaseDiskToSpare)
	# Add-MailboxDatabaseCopy seems to require some time after Remove-MailboxDatabaseCopy hence the sleep seconds
	# Sleep-ForSeconds $script:SleepSeconds
	Log-Verbose "Add-MailboxDatabaseCopy $MailboxDatabaseName -MailboxServer $MailboxServerName -SeedingPostponed"
	Add-MailboxDatabaseCopy $MailboxDatabaseName -MailboxServer $MailboxServerName -SeedingPostponed

	# If Database history file doesn't exist create one
	if (!(Test-Path "$script:DatabaseMountPath\$script:DatabaseHistoryFile"))
	{
		if (!(Create-DatabaseHistoryXML -vhFilePath "$script:DatabaseMountPath\$script:VolumeHistoryFile" -dbFolderPath $script:DatabaseMountPath -mailboxDatabaseName $MailboxDatabaseName ))
		{
			Return
		}
	}

    # Now to reseeding
    # Leaving loop here, since it might be useful later when there are multiple DBs that need reseeding for PF changes
    foreach ($copy in $problemCopy)
    {
		$reseedSuccess = $true
        $problemdb = Get-MailboxDatabase -Identity $copy.DatabaseName -Status
        if ($problemdb.Mounted)
        {
			if ($WaitBeforeReseed)
			{
				log-verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0026 -f $WaitBeforeReseed)
				sleep-ForSeconds $WaitBeforeReseed
			}
            if ($CatalogOnly)
            {
                if (!(Reseed-DagPassiveCopy -targetServer $server.Name -db $problemdb -CatalogOnly))
                {                    
					$reseedSuccess = $false					
                }
            }
            else
            {
                if (Reseed-DagPassiveCopy -targetServer $server.Name -db $problemdb)
                {
					log-verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0027 -f $script:SleepSeconds)
                    Sleep-ForSeconds $script:SleepSeconds
                    if (!(Get-DatabaseCopyHealth -DatabaseCopy "$problemdb\$server"))
					{
						$reseedSuccess = $false
					}					
                }
                else
                {                    
					$reseedSuccess = $false
                }
            }
			if ($reseedSuccess -eq $true)
			{
				Update-LogHistoryXML -Status "Complete" -Action "Reseed" -MountPoint $script:DatabaseMountPath -Path "$script:DatabaseMountPath\$script:DatabaseHistoryFile" -mailboxDatabaseName $MailboxDatabaseName| Log-Verbose
				if ( $sourceDisk -ne -1 )
				{
					Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0028 -f $sourceDisk)
					Update-LogHistoryXML -Status "Spare" -Action "Reseed" -MountPoint $spareMount -Path "$spareMount\$script:VolumeHistoryFile" -mailboxDatabaseName $MailboxDatabaseName | Log-Verbose
				}
				Send-ReseedEmail -MessageBody ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0050 -f $MailboxDatabaseName,$MailboxServerName)				
				$global:WorkflowUrgent = $false
			}
			else
			{			
				Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0029 -f $problemdb,$MailboxServerName)
				if ( $sourceDisk -ne -1 )
				{
					Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0030 -f $sourceDisk)
				}
				Update-LogHistoryXML -Status "Failed" -Action "Reseed" -MountPoint $script:DatabaseMountPath -Path "$script:DatabaseMountPath\$script:DatabaseHistoryFile" -mailboxDatabaseName $MailboxDatabaseName | Log-Verbose
				Send-ReseedEmail -MessageBody ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0051 -f $MailboxDatabaseName,$MailboxServerName)
				$global:WorkflowUrgent = $true
			}
        }
        else
        {
            Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0031 -f $problemdb)
			$global:WorkflowUrgent = $true
        }		
    }
	Return
}

# Function which initializes variables and calls Execute-Reseed
#
function Reseed-Wrapper
{

	# Use an external file for locks
	# TODO: Use steve's library for locks after his check-in for Diagnose-Replication.ps1 
	if (!(Test-Path "$script:LockPath\$MailboxDatabaseName"))
	{
		New-Item "$script:LockPath\$MailboxDatabaseName" -type file -Force
	}	
	$lockFile = New-Object System.IO.FileInfo "$script:LockPath\$MailboxDatabaseName"
	$lockStream = $lockFile.Open( [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None )
	if (!(Test-Path variable:\lockStream))
	{
		Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0052 -f $MailboxDatabaseName)
		$global:WorkflowUrgent = $false
		Return
	}		
	
	Execute-Reseeding
	
	$lockStream.close()
}

# The worker function for reseeding
#
#
function Reseed-DagPassiveCopy ([string]$targetServer, [object]$db, [switch]$CatalogOnly = $false)
{
    $script:returnCode = $true
    $script:exception = @()
    # Run the following block with the trap in a separate script block so that we can change
    # ErrorActionPerefence and the trap doesn't affect other places
    try
    {		
        if ($CatalogOnly)
        {
			# This is to trigger some traffic on an empty database, so the copy can be provably healthy
			# TODO: Test-Mailflow is still a little wacky and doesn't seem to give consistent results.
			start-job -scriptblock {sleep 10; 1..30 | %{test-mailflow -targetdatabase $MailboxDatabaseName}}
            Update-MailboxDatabasecopy $db\$targetServer -DeleteExistingFiles -ManualResume -Force -confirm:$false -CatalogOnly
        }
        else
        {
            $status = Get-MailboxDatabaseCopyStatus $db\$targetServer
            Suspend-MailboxDatabaseCopy $db\$targetServer -Confirm:$false -SuspendComment "Suspended by Reseed-DagPassiveCopy as part of the reseed workflow."
			# This is to trigger some traffic on an empty database, so the copy can be provably healthy
			# TODO: Test-Mailflow is still a little wacky and doesn't seem to give consistent results.
			start-job -scriptblock {sleep 10; 1..30 | %{test-mailflow -targetdatabase $MailboxDatabaseName}}
            Update-MailboxDatabasecopy $db\$targetServer -DeleteExistingFiles -ManualResume -Force -confirm:$false 
            if ($status.ActivationSuspended)
            {
                Resume-MailboxDatabaseCopy $db\$targetServer -Confirm:$false -ReplicationOnly:$true
            }
            else
            {
                Resume-MailboxDatabaseCopy $db\$targetServer -Confirm:$false
            }
        }		
    }
    catch {
        foreach ($failure in $_)
        {
            # Trim the message so it will not display the "ErrorActionPreference is set to Stop" message
            #
            $failedMessage = $failure.ToString()
            if ($failedMessage.IndexOf("ErrorActionPreference") -ne -1)
            {
                $failedMessage = $failedMessage.Substring($failedMessage.IndexOf("set to Stop: ") + 13)
            }
            $failedMessage = $failedMessage -replace "`r"
            $failedMessage = $failedMessage -replace "`n"
            $failedPosition = $failure.InvocationInfo.PositionMessage
            # The "PositionMessage" includes formatting along the lines of "\n  At line:23 char:37 ..." so we don't add extra.
            #Log-Error "Failed with '$failedMessage' $failedPosition"
			Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0010 -f $failedMessage,$failedPosition)

			$global:WorkflowUrgent = $true
        }
    }
    return $script:returnCode
}

# Return number of times the database was reseeded in the specified number of days
#
function Get-ReseedCount ([int]$LastDays = $script:ReseedIntervalInDays)
{	
	if (!(Test-Path "$script:DatabaseMountPath\$script:DatabaseHistoryFile"))
	{
		Log-Warning ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0053 -f $script:DatabaseMountPath,$script:DatabaseHistoryFile,"Get-ReseedCount")
		Return 0
	}
	Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0054 -f $MailboxDatabaseName,$LastDays,"Get-ReseedCount")
	[xml]$xml = Get-Content "$script:DatabaseMountPath\$script:DatabaseHistoryFile"	
	$lastDate = [DateTime]::Now.AddDays(-$LastDays)
	$events = @($xml.VolumeHistory.Event)
	$reseedEvents = @($events | where { ($_.Action -eq "Reseed") -and ($_.Status -eq "Complete") -and ([System.DateTime]$_.Time -ge $lastDate) })		
	Return $reseedEvents.Count
}

# Verify the database is sufficiently healthy for reseed
# Returns true if there is alteast NumberOfHealthyPassiveCopiesRequired number of healthy passive copy of the Database
function Get-DatabaseHealth([Microsoft.Exchange.Data.Directory.SystemConfiguration.MailboxDatabase] $Database)
{	
	$NumberOfHealthyPassiveCopies = 0
	foreach ( $DatabaseCopy in $Database.DatabaseCopies )
	{
		if ( ( Get-DatabaseCopyHealth $DatabaseCopy ) )
		{       	
			$NumberOfHealthyPassiveCopies++
		}
	}	
	Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0055 -f $Database,$NumberOfHealthyPassiveCopies,"Get-DatabaseHealth")	
	
	return $NumberOfHealthyPassiveCopies -ge $script:NumberOfHealthyPassiveCopiesRequired	
}

# This function returns $true if a specific database copy is sufficiently healthy.
# Sufficiently healthy is defined as the copy status returning as "Healthy" and
# LogsCopiedSinceInstanceStart should be -ge than $script:RequiredLogsCopied
# LogsReplayedSinceInstanceStart should be -ge than $script:RequiredLogsReplayed
function Get-DatabaseCopyHealth($DatabaseCopy)
{		
	Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0056 -f $DatabaseCopy,"Get-DatabaseCopyHealth")
	$copyStatus = Get-MailboxDatabaseCopyStatus $DatabaseCopy
	if (($copyStatus.ActiveCopy -eq $false) -and ($copyStatus.Status -eq $CopyStatusType::Healthy))
	{
		if (($copyStatus.LogsReplayedSinceInstanceStart -ge $script:RequiredLogsReplayed) -and `
			($copyStatus.LogsCopiedSinceInstanceStart -ge $script:RequiredLogsCopied))
		{
			Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0057 -f $DatabaseCopy,"Get-DatabaseCopyHealth")
			return $true
		}
		else
		{
			Log-Warning ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0058 -f $DatabaseCopy,$script:RequiredLogsCopied,$script:RequiredLogsCopied,"Get-DatabaseCopyHealth","Get-MailboxdatabaseCopyStatus")
			return $false
		}
	}	
	else
	{
		Log-Warning ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0059 -f $DatabaseCopy,"Get-DatabaseCopyHealth")
		return $false
	}	
}


# Function to send email, takes status success/failure message as input
function Send-ReseedEmail ([string]$MessageBody)
{
	if ($NotifyEmail)
	{		
		if (Test-Path "$RoleDatacenterPath\DatacenterSvcEngCommonLibrary.ps1")
		{
			Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0038 -f $RoleDatacenterPath)
			. "$RoleDatacenterPath\DatacenterSvcEngCommonLibrary.ps1"
			Send-NotificationMail -tos @($NotifyEmail) -title "Reseed Notification" -body $MessageBody
		}
		else
		{
			Log-Warning ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0039 -f $NotifyEmail)
		}
	}
}

# Sleep for the specified duration (in seconds)
function Sleep-ForSeconds ( [int]$sleepSecs )
{
	#Log-Verbose "Sleeping for $sleepSecs seconds..."
	Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0013 -f $sleepSecs)
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
    if ($Verbose)
    {
		$timeStamp = Get-CurrentTimeString
		Write-Verbose "$timeStamp $msg"
	}
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

if ($MailboxDatabaseName)
{
	$db = Get-mailboxdatabase | where {$_.Name -ieq $MailboxDatabaseName}
	if ( !$db )
	{
		log-error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0060 -f $MailboxDatabaseName) -stop
	}
	$script:DatabasePath = Split-Path $db.EdbFilePath -Parent
}

$script:DatabaseMountPath = Split-Path $DatabasePath -Parent

if ($SpareDiskToUse)
{	
	Log-Warning ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0061 -f $SpareDiskToUse,"-SpareDiskToUse")
    
	$prompt = Read-Host ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0062 -f $SpareDiskToUse)
	if ($prompt -ne "Y")
	{
		Exit
	}
}

if (Test-Path variable:\WorkflowMailboxDatabaseName)
{
	$script:IsOperator = $false
}

Reseed-Wrapper

# SIG # Begin signature block
# MIIafQYJKoZIhvcNAQcCoIIabjCCGmoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUbS4aF+JaNsU1/T8AoOwR1FxV
# tJGgghUmMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0xggTB
# MIIEvQIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQITMwAAAJ0ejSeu
# uPPYOAABAAAAnTAJBgUrDgMCGgUAoIHjMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBQA28xNhXSinDSL5sO3iVyAbcIgETCBggYKKwYBBAGCNwIBDDF0MHKgSoBIAE0A
# YQBpAGwAYgBvAHgARABhAHQAYQBiAGEAcwBlAFIAZQBzAGUAZQBkAFUAcwBpAG4A
# ZwBTAHAAYQByAGUAcwAuAHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNv
# bS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAo2IkWpBc9VmeQtx9bfBQGIVf
# wA3SWURbR3Nh6PAXbk6JQ4oS5fVTxZwElNaeupA/gLPv4eVY1ZiISJxDOtmfR42g
# 4YjNMLUOmazrzl77n+SLfsZXLDHDLG10FCw0TKvi/+u3qhXnMfXVq7DZ9f+7qX8Q
# YJabFXCl/55l23li6yVertmV2Q4QVYJD128GvXuSZ0jU6ug7zvEvUAatrDRmvN04
# SAdCn2Cm1JEEaxBw0WrIVu+Hb5nqgDKL9qLLBlvXRA1lysHyJ67Q3QkjVjmY3a9C
# uhUcgV+rf0I6Gdxoyga4xSkZZmePD5ytiF9FfX00z4YoNNnCcTUio2clPUSPPqGC
# Ah8wggIbBgkqhkiG9w0BCQYxggIMMIICCAIBATCBhTB3MQswCQYDVQQGEwJVUzET
# MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
# TWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1T
# dGFtcCBQQ0ECCmECjkIAAAAAAB8wCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzEL
# BgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTEzMDEwODA4NDY1NlowIwYJKoZI
# hvcNAQkEMRYEFEwmPpURrbYpwQ2sPLZlB1+j44pYMA0GCSqGSIb3DQEBBQUABIIB
# AG1TLqR9CiBA/Q0GDSomXqfpCrj7u+22bRMhEMg6mnxOfuLDrQE0MA3W9Ui80Wty
# tOLY08BF57HiEtFJvJ0/cGOqArw5zdXPkP6dXTlfWK3zMvCdPzCfLhjgIzSE3w2F
# 9K/+fuE2fsLG20762271PFgZaL18WgItGkmiOYaSafoBCi02lxUfu193pUhBU5uB
# xGo3NgTFVufxD6ygypGyfG6JJVHAbeaLbDbmJBGC19wd3GAWMlV62nj/0f6XOVhn
# zhYYyO3LG3y7tiyArpZOPjSIfMUgAYUlKQyMQxwaG4eHqWoRhCBPSmAVJXv39nf2
# OvuFlE/rRBFjBessL/BxINk=
# SIG # End signature block
