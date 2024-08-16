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
# MIIa4gYJKoZIhvcNAQcCoIIa0zCCGs8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUbS4aF+JaNsU1/T8AoOwR1FxV
# tJGgghWCMIIEwzCCA6ugAwIBAgITMwAAAHD0GL8jIfxQnQAAAAAAcDANBgkqhkiG
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
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBMowggTG
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAAAymzVMhI1xOFV
# AAEAAADKMAkGBSsOAwIaBQCggeMwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFADb
# zE2FdKKcNIvmw7eJXIBtwiARMIGCBgorBgEEAYI3AgEMMXQwcqBKgEgATQBhAGkA
# bABiAG8AeABEAGEAdABhAGIAYQBzAGUAUgBlAHMAZQBlAGQAVQBzAGkAbgBnAFMA
# cABhAHIAZQBzAC4AcABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL2V4
# Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQCCmMnqAlyPr/DqgO5qWSqdY6g2W4SC
# t2lTCaICPp0QDNFhq5iknfVS5Auq8iB1GmFAWuSOQRUP5VsuwGKWvKzW7u8avKDx
# Ah1GVGU5xksA8Vk4WEd1w9XT8bYBk5zrdzvzUGzPzwdFvnzpyMERvR35y6hgqztX
# IDfGFn71RWyxmg5ijhJ4wrUKIyeWEuvY61belURoNXKH+3n1KHtjAivabsKnRxDy
# DqWLckbufwGrYZT6pdb1vr28BZkS3pen6CDGbJS4o5P1aGGxMSpJENw+xGDgn65P
# wG8O8m/UJzVjC49VwSEIkDdxP2QBBxDQ7fD4qpZesMHFILOyZuzr59iBoYICKDCC
# AiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQQITMwAAAHD0GL8jIfxQnQAAAAAAcDAJBgUrDgMCGgUAoF0wGAYJKoZIhvcN
# AQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTUwNDEwMDI1NzQzWjAj
# BgkqhkiG9w0BCQQxFgQUhTklsu+raD5sapiu5QzgqdAOLKowDQYJKoZIhvcNAQEF
# BQAEggEAWvqWZwpQc9yldQbeQuh9Xlbl/DvY6cxuQAECui7jYFGD0vQ9QrzBc2hj
# B2FjfW/vOJs4g4dYG8yJTDPsD55F3oFjeHeD9vTzZWmJRcxjgByVP+fXQsDOySbG
# KCF3oXPNVhE/y3AR3w7SnTmqKLXGUL6YwS2lR3N8LAgHiRP3EZDsGJNteA6d35Tr
# DXDFD8X2ucwXbHqRnKcIREW6Lq95IHIt/LnQQ2xW8CtmWFViWJSmBsjW/I/slzXt
# pqNJR5Uw0Lsij0dAWSbXEP9GmulQyQrzar90Q84g0H8eC0MazJHjpVKY/lEJ5PsR
# IXoNUhOnaoBBCWidvlqpe4mF6yI44w==
# SIG # End signature block
