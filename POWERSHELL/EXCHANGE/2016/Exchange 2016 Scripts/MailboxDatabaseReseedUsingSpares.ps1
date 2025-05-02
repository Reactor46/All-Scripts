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
	[switch]	$ReseedConfirm=$false,
	[switch]	$Verbose=$false,
	[int]		$SpareDiskToUse=$null,
	# Test Hooks
	[switch]	$SkipSpareCheck=$false,
	[switch]	$WhatIf=$false,
	[int]		$SpareDriveOverride=$null,
	[int]		$WaitBeforeReseed,
	[int]		$ReseedCount=5)

Import-LocalizedData -BindingVariable MailboxDatabaseReseedUsingSpares_LocalizedStrings -FileName MailboxDatabaseReseedUsingSpares.strings.psd1
Set-StrictMode -Version 2.0

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ea SilentlyContinue
Add-PSSnapin Microsoft.Exchange.Management.Powershell.Setup -ea SilentlyContinue
Add-PSSnapin Microsoft.Exchange.Management.Powershell.Support -ea SilentlyContinue
Add-PSSnapin Microsoft.Exchange.Management.Powershell.CentralAdmin -ea SilentlyContinue

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
[string]$script:DatabaseMountPath = ""
[switch]$script:IsOperator = $true
$script:dbcopyInfo = "DatabaseCopy: $MailboxDatabaseName\$MailboxServerName"
$script:eventMsg = ""
$eventIDHash = @{"Started" = "400410E8"; "PrereqFailed" = "C00410E9"; "NoSpareDisk" = "C00410EA";
				"ReseedFailed" = "C00410EB"; "DiskSwapFailed" = "C00410E9"; "ReseedSuccess" = "400410EC";}
$script:eventID = $eventIDHash.PrereqFailed
$CopyStatusType = [Microsoft.Exchange.Management.SystemConfigurationTasks.CopyStatus]
[string]$version = [Microsoft.Exchange.Diagnostics.BuildVersionConstants]::RegistryVersionSubKey

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

$ExchangeScriptsPath = Join-Path (Get-ItemProperty -path "HKLM:SOFTWARE\Microsoft\ExchangeServer\$version\Setup").MsiInstallPath "Scripts"

# If RoleDatacenterPath is not defined use current directory
if ( !( Test-Path variable:\RoleDatacenterPath ) )
{
	$RoleDatacenterPath = split-path $myInvocation.MyCommand.Path;
}

# Load some of the common functions.
# The common lib doesn't use clean practices so we have to avoid strict mode
Set-StrictMode -Off

. "$RoleDatacenterPath\DatacenterSpareDiskManipulation.ps1";
. "$RoleDatacenterPath\DatacenterLockHelper.ps1";
. "$ExchangeScriptsPath\CheckDatabaseRedundancy.ps1"

Set-StrictMode -Version 2.0

$ErrorActionPreference = "Stop"

# Function which initializes variables and calls Execute-Reseed
#
function Reseed-Wrapper
{
	try
	{	
		$lock = AllocateLock "MailboxDatabaseReseedUsingSpares.ps1"				
		Execute-Reseeding
	}
	catch
	{
		$script:eventMsg += "$_"
		Write-HAAppLogErrorEvent $script:eventID 1 @($script:dbCopyInfo, $script:eventMsg)
		throw $_
	}
	finally
	{
		FreeLock $lock
	}
}

# function to execute reseed checks and proceed to reseed
#
function Execute-Reseeding
{	
    Log-Verbose "Execute-Reseeding -MailboxServer $MailboxServerName -Database $MailboxDatabaseName -SingleActiveCopyOK $SingleActiveCopyOK -ReseedConfirm $ReseedConfirm -Verbose $Verbose"	
	
    $server = Get-ExchangeServer $MailboxServerName
    $problemdb = Get-MailboxDatabase $MailboxDatabaseName -Status
	
	# If 1)the MailboxServer value is not specified or invalid or 2)the Database value is specified but invalid, we log an error
    #
    if ((!$server) -or (!$MailboxDatabaseName) -or (!$problemdb))
    {
        #Log-Error "No valid database or target mailbox server name for reseeding was provided. Specified database name: '$MailboxDatabaseName', mailbox server name: '$MailboxServerName'"
		Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0001 -f $MailboxDatabaseName,$MailboxServerName)		
        Return
    }
	
	# Load values for script level globals
	$script:DatabasePath = Split-Path $problemdb.EdbFilePath -Parent
	$script:DatabaseMountPath = Split-Path $script:DatabasePath -Parent
	
    # Checking whether the specified database copy is FailedAndSuspended     
	Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0049 -f $problemdb,$server,"Get-MailboxDatabaseCopyStatus")
    $copyStatus = Get-MailboxDatabaseCopyStatus -ExtendedErrorInfo $problemdb\$server

    if ($copyStatus.Status -eq $CopyStatusType::FailedAndSuspended)
    {
		$script:dbCopyInfo += "`nOriginalErrorEventID: " + $copyStatus.ErrorEventId + "`nErrorMessage: " + $copyStatus.ErrorMessage + "`nException: " + $copyStatus.ExtendedErrorInfo
		Write-HAAppLogInformationEvent $eventIDHash.Started 1 @($script:dbCopyInfo)
    }
    else
    { 
        #Log-Warning "Mailbox database '$MailboxDatabaseName' is not 'FailedAndSuspended' status on server $MailboxServerName. Reseeding skipped."
		Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0004 -f $MailboxDatabaseName,$MailboxServerName,$copyStatus.Status)		
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
			Return
		}
	}
	else
	{
		if ($count -ge $script:MaxWorkflowReseedCount)
		{	
			Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0044 -f $script:ReseedIntervalInDays)			
			Return
		}
	}
	
	$script:eventID = $eventIDHash.NoSpareDisk
	$sourceDisk = 0
	$spareDisk = 0
	# Verify Spare Allocation
	if (!$SkipSpareCheck)
	{
		$sourceDisk = Get-DatabaseDisk -DatabasePath $script:DatabasePath
		$script:eventMsg += "`nSource Disk: $sourceDisk "
		if ($sourceDisk -le 0)
		{
			Log-Warning ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0020 -f $MailboxDatabaseName)
			# This disk probably crashed, setting this -1 so Swap-DatabaseDiskToSpare allocates a new disk for this database
			$sourceDisk = -1
			$script:eventMsg += "`nCould not locate database disk. Disk is probably failed. Reseeding to spare disk."
		}
		
		# Verify Spare Allocation
		$spareHash = Reserve-SpareVolume -mailboxDatabaseName $MailboxDatabaseName
		$spareDisk = Get-DiskFromVolume -Volume $($spareHash.reserveVolume)
		$spareMount = $spareHash.reserveMount
		Log-Verbose "Total number of spare disks found in $MailboxServerName = ($spareHash.numberOfSpares)"
		$script:eventMsg += "`nTarget Spare Disk: $spareDisk `nSpare disk mountpoint: $spareMount `nTotal number of spares found = " + $spareHash.numberOfSpares
		if ($spareDisk -le 0)
		{
			Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0046 )			
			Return
		}
	}
	else
	{	
		if(!$ReseedConfirm)
		{	
			Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0047 -f "-SkipSpareCheck")			
			Return
		}
	}
	$script:eventID = $eventIDHash.DiskSwapFailed

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
		Return
	}

	# Spare manipulation. This will either allocate a spare (if there isn't a database disk),
	# or it will swap the database disk with a spare disk.
	if (Swap-DatabaseDiskToSpare -SourceDisk $sourceDisk -SpareDisk $spareDisk -SpareMount $spareMount -databaseMountPath $script:DatabaseMountPath -mailboxDatabaseName $MailboxDatabaseName )
	{
		$script:eventMsg += "`nSwap to spare disk succeeded...`nNew Database disk: $spareDisk `nQuarantined database disk: $sourceDisk"
		Log-Verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0023 -f $MailboxDatabaseName,$sourceDisk,$spareDisk)
	}
	else
	{
		Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0024 -f $sourceDisk,$spareDisk,$spareMount)		
		Return
	}

	# If Database history file doesn't exist create one
	if (!(Test-Path "$script:DatabaseMountPath\$script:DatabaseHistoryFile"))
	{
		Create-DatabaseHistoryXML -vhFilePath "$script:DatabaseMountPath\$script:VolumeHistoryFile" -dbFolderPath $script:DatabaseMountPath -mailboxDatabaseName $MailboxDatabaseName 		
	}
	
	$script:eventID = $eventIDHash.ReseedFailed

    # Now to reseeding
	$reseedSuccess = $true
    if (!$problemdb.Mounted)
    {
		Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0031 -f $problemdb)	
	}		
	
	if ($WaitBeforeReseed)
	{
		log-verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0026 -f $WaitBeforeReseed)
		sleep-ForSeconds $WaitBeforeReseed
	}            
    if (Reseed-DagPassiveCopy -targetServer $server.Name -db $problemdb)
    {
		log-verbose ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0027 -f $problemdb)					
    }
    else
    {                    
		$reseedSuccess = $false
    }
    
	if ($reseedSuccess -eq $true)
	{
		Update-LogHistoryXML -Status "Complete" -Action "Reseed" -MountPoint $script:DatabaseMountPath `
			-Path "$script:DatabaseMountPath\$script:DatabaseHistoryFile" -mailboxDatabaseName $MailboxDatabaseName| Log-Verbose			
		Send-ReseedEmail -MessageBody ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0050 -f $MailboxDatabaseName,$MailboxServerName)			
		Write-HAAppLogInformationEvent $eventIDHash.ReseedSuccess 1 @($script:dbCopyInfo)
	}
	else
	{
		$script:eventMsg += "`nUpdate-MailboxdatabaseCopy Failed. Souce disk: $sourceDisk has been left offline with the original databasecopy files."
		Update-LogHistoryXML -Status "Failed" -Action "Reseed" -MountPoint $script:DatabaseMountPath `
			-Path "$script:DatabaseMountPath\$script:DatabaseHistoryFile" -mailboxDatabaseName $MailboxDatabaseName | Log-Verbose				
		Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0029 -f $problemdb,$MailboxServerName)			
	} 
	Return
}

# The worker function for reseeding
#
#
function Reseed-DagPassiveCopy ([string]$targetServer, [object]$db)
{
    $script:returnCode = $true
    $script:exception = @()
    # Run the following block with the trap in a separate script block so that we can change
    # ErrorActionPerefence and the trap doesn't affect other places
    try
    {   
        $status = Get-MailboxDatabaseCopyStatus $db\$targetServer
		# This is to trigger some traffic on an empty database, so the copy can be provably healthy
		# TODO: Test-Mailflow is still a little wacky and doesn't seem to give consistent results.
		start-job -scriptblock {sleep 10; 1..30 | %{test-mailflow -targetdatabase $MailboxDatabaseName}}
        Update-MailboxDatabasecopy $db\$targetServer -Force -confirm:$false
    }
    catch {
        $script:returnCode = $false
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
			$script:eventMsg += "`nReseeding passive copy failed: $failedMessage"
			Log-Error ($MailboxDatabaseReseedUsingSpares_LocalizedStrings.res_0010 -f $failedMessage,$failedPosition) -Stop	
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
# MIIdyQYJKoZIhvcNAQcCoIIdujCCHbYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU4JPPYjVLqYL+DOftlELRkiZ1
# CRugghhkMIIEwzCCA6ugAwIBAgITMwAAAJvgdDfLPU2NLgAAAAAAmzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI5
# WhcNMTcwNjMwMTkyMTI5WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjcyOEQtQzQ1Ri1GOUVCMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAjaPiz4GL18u/
# A6Jg9jtt4tQYsDcF1Y02nA5zzk1/ohCyfEN7LBhXvKynpoZ9eaG13jJm+Y78IM2r
# c3fPd51vYJxrePPFram9W0wrVapSgEFDQWaZpfAwaIa6DyFyH8N1P5J2wQDXmSyo
# WT/BYpFtCfbO0yK6LQCfZstT0cpWOlhMIbKFo5hljMeJSkVYe6tTQJ+MarIFxf4e
# 4v8Koaii28shjXyVMN4xF4oN6V/MQnDKpBUUboQPwsL9bAJMk7FMts627OK1zZoa
# EPVI5VcQd+qB3V+EQjJwRMnKvLD790g52GB1Sa2zv2h0LpQOHL7BcHJ0EA7M22tQ
# HzHqNPpsPQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFJaVsZ4TU7pYIUY04nzHOUps
# IPB3MB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBACEds1PpO0aBofoqE+NaICS6dqU7tnfIkXIE1ur+0psiL5MI
# orBu7wKluVZe/WX2jRJ96ifeP6C4LjMy15ZaP8N0OckPqba62v4QaM+I/Y8g3rKx
# 1l0okye3wgekRyVlu1LVcU0paegLUMeMlZagXqw3OQLVXvNUKHlx2xfDQ/zNaiv5
# DzlARHwsaMjSgeiZIqsgVubk7ySGm2ZWTjvi7rhk9+WfynUK7nyWn1nhrKC31mm9
# QibS9aWHUgHsKX77BbTm2Jd8E4BxNV+TJufkX3SVcXwDjbUfdfWitmE97sRsiV5k
# BH8pS2zUSOpKSkzngm61Or9XJhHIeIDVgM0Ou2QwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBM8wggTLAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB4zAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUH0cpBGtadEJECuHwxR0vdJyO6YwwgYIGCisG
# AQQBgjcCAQwxdDByoEqASABNAGEAaQBsAGIAbwB4AEQAYQB0AGEAYgBhAHMAZQBS
# AGUAcwBlAGUAZABVAHMAaQBuAGcAUwBwAGEAcgBlAHMALgBwAHMAMaEkgCJodHRw
# Oi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIB
# AEEnP3Cl2RVH0vlGVEBk5T/FgG3C9GTxs9C4Vtr1PovXesrMVUG9lLoUe1SUiR4G
# amKz+a2kVGwFQPGZul8Fs4+fCsfeMIhOzoMHZ5utpQ8xbJMmq2fiiGsokavkgIVO
# XzDE6ed0dcl3IraF48xRkLJGFzfZOFXjzV+Vqkyi1F5Qvai6x6n7YGNb9wDNJv/T
# nXWpi/5FS+WSdr5TGg0zEzXlzIRjOxeK6qNtD6FN53O3FFchOnMyB3JR3fUnqLuW
# kUGG+9OixyExCsEpyxLsUIUvw3/LU3jHasSblZHRG4axQNKpVZSrh5+TGAZG2WFP
# 1diwzAODK1o/XWitaVtuswihggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEw
# gY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UE
# AxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAm+B0N8s9TY0uAAAAAACb
# MAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3
# DQEJBTEPFw0xNjA5MDMxODQ1MDJaMCMGCSqGSIb3DQEJBDEWBBRKrgAIh+lYf4d6
# TT/jO6FgS3x3pjANBgkqhkiG9w0BAQUFAASCAQAxf65BTVBLGMvNo+5FZC+EgkIe
# RDc+5+cEtZbBTUUVyzx8TVdQuNjUOhUmrJXu6S/HqVwLcbFMowGkNouXVNkwtF4W
# Ov3TtL5gh9j7G5pajNWon0hyTw+ElOtxEsrj1vyXhR6xl5EFdWjkKn2vKqxou/MK
# yK3k2QBmmOogRs0sZnZ4Ppj4GaNaQ2VGWzaEKwy3isiWgIN3ELopBiFO9TGTMAp3
# gEv5ZSiDb3ygux6m26CdOzWpn2thF4lwspm05VUS5zsAj5TInm7CiUJu536J9KB6
# zEzS6iFFzlem4vIMJqnlSXPRwsl44QINuqpkYgsNqlEIm4sjxslwvUXEuJwX
# SIG # End signature block
