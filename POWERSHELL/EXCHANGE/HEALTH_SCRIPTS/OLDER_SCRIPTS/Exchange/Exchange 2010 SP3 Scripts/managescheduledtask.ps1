# Copyright (c) Microsoft Corporation. All rights reserved.

param(
	[Parameter(ParameterSetName="Install",Mandatory=$true)]
	[switch] $Install,
	
	[Parameter(ParameterSetName="Install")]
	[switch] $DeleteExisting = $false,
	
	[Parameter(Mandatory=$true)]
	[System.Management.Automation.ValidateNotNullOrEmptyAttribute()]
	[string] $ServerName,
	
	[Parameter(Mandatory=$true)]
	[System.Management.Automation.ValidateNotNullOrEmptyAttribute()]
	[string] $TaskName,
	
	[Parameter(ParameterSetName="Install",Mandatory=$true)]
	[System.Management.Automation.ValidateNotNullOrEmptyAttribute()]
	[string] $PsScriptPath,
	
	[Parameter(ParameterSetName="Install")]
	[System.Management.Automation.ValidateNotNullOrEmptyAttribute()]
	[string] $PsScriptArgs,
	
	[Parameter(ParameterSetName="Uninstall",Mandatory=$true)]
	[switch] $Uninstall,
	
	# Disable a scheduled task without unregistering it
	[Parameter(ParameterSetName="Disable",Mandatory=$true)]
	[switch] $Disable,
	
	# Re-enable a pre-configured scheduled task and run it
	[Parameter(ParameterSetName="Enable",Mandatory=$true)]
	[switch] $Enable,
	
	# Returns $true or $false based on whether the specified task exists on the given server
	[Parameter(ParameterSetName="TestExistence",Mandatory=$true)]
	[switch] $TestExistence,
	
	# True if run from Exchange Setup (will cause the script to log to 
	# the Exchange setup logs)
	[Parameter(Mandatory=$false)]
	[switch] $RunFromSetup = $false
)

Set-StrictMode -Version 2.0
Import-LocalizedData -BindingVariable ManageScheduledTask_LocalizedStrings -FileName ManageScheduledTask.Strings.psd1

#-------------------
# Script variables #
#-------------------

$script:exeOutput = $null
$script:taskQueried = $null
$script:writeToExchangeSetupLog = $false
$script:exLogDir = "$env:SYSTEMDRIVE\ExchangeSetupLogs"


# Returns the name of the user this script is running as: eg: SDIMEB-DOM\Administrator
function Get-CurrentUserName
{
	[System.Security.Principal.WindowsIdentity]$id = [System.Security.Principal.WindowsIdentity]::GetCurrent()

	# E14:398785: to support a case in which an ampersand is appeared in the NetBIOS domain name.
	[System.String]$fixedName = $id.Name.Replace('&','&amp;')
	return $fixedName
}

# 
function Get-TaskXml(
	[Parameter(Mandatory=$true)] [string] $powershellCommand
)
{
	# XML describing the scheduled task to create
	#
	# Parameters:
	# {0} : current date computed using: (get-date).ToString("yyyy-MM-ddTHH:mm:ss")
	# {1} : user name
	# {2} : start time of the task
	# {3} : windows directory path: eg: $env:windir = D:\Windows
	# {4} : powershell command to execute: eg: D:\test.ps1 -MonitoringContext -ErrorAction:Stop
	[string]$xmlContent = 
'<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <!-- <Date>2010-01-15T03:00:00</Date> -->
	<Date>{0}</Date>
    <!-- <Author>SDIMEB-DOM\administrator</Author> -->
	<Author>{1}</Author>
    <Description>Database redundancy monitoring task.</Description>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <Repetition>
        <Interval>PT1H</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
      <!-- <StartBoundary>2010-01-15T03:00:00</StartBoundary> -->
	  <StartBoundary>{2}</StartBoundary>
      <ExecutionTimeLimit>PT2H</ExecutionTimeLimit>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>NT AUTHORITY\SYSTEM</UserId>
<!--      <LogonType>Password</LogonType>    -->
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>Queue</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>false</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>true</WakeToRun>
    <ExecutionTimeLimit>PT2H</ExecutionTimeLimit>
    <Priority>7</Priority>
	<!-- <RestartOnFailure>
      <Interval>PT5M</Interval>
      <Count>5</Count>
    </RestartOnFailure> -->
  </Settings>
  <Actions Context="Author">
    <Exec>
      <!-- <Command>D:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe</Command> -->
	  <Command>{3}\System32\WindowsPowerShell\v1.0\powershell.exe</Command>
      <!-- <Arguments>-version 2.0 -NonInteractive -WindowStyle Hidden -command "D:\test.ps1 -MonitoringContext -ErrorAction:Continue"</Arguments> -->
	  <Arguments>-version 2.0 -NonInteractive -WindowStyle Hidden -command "{4}"</Arguments>
      <!-- <WorkingDirectory>D:\</WorkingDirectory> -->
    </Exec>
  </Actions>
</Task>'

	[DateTime]$currentDate = Get-Date
	$currentDateString = $currentDate.ToString("yyyy-MM-ddTHH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
	$userName = Get-CurrentUserName
	[DateTime]$newDate = $currentDate.AddMinutes(2)
	$newDateString = $newDate.ToString("yyyy-MM-ddTHH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)
	
	Log-Verbose "Get-TaskXml: `$currentDateString=$currentDateString, `$userName=$userName, `$newDateString=$newDateString, WindowsDir=$($env:windir), `$powershellCommand=$powershellCommand"
	[string]$returnXmlString = [string]::Format($xmlContent, $currentDateString, $userName, $newDateString, $env:windir, $powershellCommand)
	return $returnXmlString
}

# Returns $true or $false based on whether the specified task exists on the given server
function Is-TaskExisting(
	[Parameter(Mandatory=$true)] [string] $_serverName,
	[Parameter(Mandatory=$true)] [string] $_taskName
)
{
	Log-Verbose "Is-TaskExisting: Entering: `$_serverName=$_serverName, `$_taskName=$_taskName"
	
	# Create the task scheduler COM object
	$st = New-Object -ComObject "Schedule.Service"
	# connect to remote server
	$st.Connect($_serverName, $null, $null, $null)
	$rootFolder = $st.GetFolder("\")
	$script:taskQueried = $null
	
	$checkCommand = 
	{
		$script:taskQueried = $rootFolder.GetTask($_taskName)
	}
	[bool]$taskFound = TryExecute-ScriptBlock -runCommand $checkCommand -silentOnErrors $true
	
	Log-Verbose "Is-TaskExisting: Returning '$taskFound'."
	Write-Output $taskFound
}


# Enable a pre-configured scheduled task and start running it (uses the COM interface).
# Returns $true if the task already exists, or if it was successfully created; $false otherwise.
function Enable-TaskUsingCOM (
	[Parameter(Mandatory=$true)] [string] $_serverName,
	[Parameter(Mandatory=$true)] [string] $_taskName
)
{
	Log-Verbose "Enable-TaskUsingCOM: Entering: `$_serverName=$_serverName, `$_taskName=$_taskName"
	
	# Create the task scheduler COM object
	$st = New-Object -ComObject "Schedule.Service"
	# connect to remote server
	$st.Connect($_serverName, $null, $null, $null)
	$rootFolder = $st.GetFolder("\")
	$script:taskQueried = $null
	
	$checkCommand = 
	{
		$script:taskQueried = $rootFolder.GetTask($_taskName)
	}
	[bool]$taskFound = TryExecute-ScriptBlock -runCommand $checkCommand -silentOnErrors $true
	
	if ($taskFound)
	{
		# mark the task as enabled
		$script:taskQueried.Enabled = $true
	
		# now, run the task
		$script:taskQueried.Run($null) | out-null
		
		Log-Verbose "Enable-TaskUsingCOM: Task successfully enabled and started."
	}
	else
	{
		Log-Error ($ManageScheduledTask_LocalizedStrings.res_0000 -f $_taskName,$_serverName,"Enable-TaskUsingCOM")
	}
}


# Create a scheduled task using the COM scripting interface
# Returns $true if the task already exists, or if it was successfully created; $false otherwise.
function Create-TaskUsingCOM (
	[Parameter(Mandatory=$true)] [string] $_serverName,
	[Parameter(Mandatory=$true)] [string] $_taskName,
	[Parameter(Mandatory=$true)] [string] $powershellCommand
)
{
	Log-Verbose "Create-TaskUsingCOM: Entering: `$_serverName=$_serverName, `$_taskName=$_taskName, `$powershellCommand=$powershellCommand"
	
	[string]$xmlString = Get-TaskXml $powershellCommand
	
	Log-Verbose "Create-TaskUsingCOM: `$xmlString: `n$xmlString"
	
	# check if the task has already been registered
	[bool]$alreadyRegistered = Is-TaskRegistered $_serverName $_taskName $xmlString
	if ($alreadyRegistered)
	{
		Log-Verbose "Create-TaskUsingCOM: Task '$_taskName' has already been registered on server '$_serverName'."
		return $true
	}
	
	# Create the task scheduler COM object
	$st = New-Object -ComObject "Schedule.Service"
	# connect to remote server
	$st.Connect($_serverName, $null, $null, $null)
	# Get the root folder to create the task in
	$rootFolder = $st.GetFolder("\")
	# Define the new task - MSDN says 0 must be passed in.
	$taskDef = $st.NewTask(0)
	$taskDef.XmlText = $xmlString

	# Finally, register the task:
	# 	flags 		= TASK_CREATE_OR_UPDATE (6)
	# 	logonType 	= TASK_LOGON_SERVICE_ACCOUNT (5)
	$task = $rootFolder.RegisterTaskDefinition($_taskName, $taskDef, 6, $null, $null, 5)	
	
	# Check that the task was successfully registered (wait upto 10 seconds)
	Log-Verbose "Create-TaskUsingCOM: Waiting upto 10 seconds for task to be registered..."
	
	$condition = { return (Is-TaskRegistered $_serverName $_taskName $xmlString) }
	[bool]$isRegistered = WaitForCondition $condition 10
	if ($isRegistered)
	{
		Log-Verbose "Create-TaskUsingCOM: Task '$_taskName' registered successfully on server '$_serverName'."
		return $true
	}
	else
	{
		Log-Error ($ManageScheduledTask_LocalizedStrings.res_0001 -f $_taskName,$_serverName,"Create-TaskUsingCOM")
		return $false
	}
}


# Delete a scheduled task that has already been registered.
function Delete-RegisteredTask(
	[Parameter(Mandatory=$true)] [string] $_serverName,
	[Parameter(Mandatory=$true)] [string] $_taskName)
{
	StopAndDelete-RegisteredTaskInternal $_serverName $_taskName $true
}

# Stop and optionally delete a scheduled task that has already been registered.
function StopAndDelete-RegisteredTaskInternal(
	[Parameter(Mandatory=$true)]  [string] $_serverName,
	[Parameter(Mandatory=$true)]  [string] $_taskName,
	[Parameter(Mandatory=$false)] [bool]   $_delete=$false)
{
	Log-Verbose "StopAndDelete-RegisteredTaskInternal: Entering: `$_serverName=$_serverName, `$_taskName=$_taskName, `$_delete=$_delete"
	
	# Create the task scheduler COM object
	$st = New-Object -ComObject "Schedule.Service"
	# connect to remote server
	$st.Connect($_serverName, $null, $null, $null)
	$rootFolder = $st.GetFolder("\")
	$script:taskQueried = $null
	
	$checkCommand = 
	{
		$script:taskQueried = $rootFolder.GetTask($_taskName)
	}
	[bool]$taskFound = TryExecute-ScriptBlock -runCommand $checkCommand -silentOnErrors $true
	
	if ($taskFound)
	{
		Log-Verbose "StopAndDelete-RegisteredTaskInternal: Found the existing task."
			
		# stop the running instances
		$script:taskQueried.Stop(0)
		
		# mark the task as disabled
		$script:taskQueried.Enabled = $false
				
		if ($_delete)
		{
			# delete the task
			$rootFolder.DeleteTask($_taskName, 0)
			Log-Verbose "StopAndDelete-RegisteredTaskInternal: Task '$_taskName' successfully deleted."
		}
	}
	else
	{
		Log-Verbose "StopAndDelete-RegisteredTaskInternal: The task does not exist. Exiting."
	}
}


# Returns $true if a specified task is already registered and enabled on the given server.
function Is-TaskRegistered( 
	[Parameter(Mandatory=$true)] [string] $_serverName,
	[Parameter(Mandatory=$true)] [string] $_taskName,
	[Parameter(Mandatory=$true)] [string] $_taskXmlString
)
{
	[xml]$taskXml = [xml]$_taskXmlString
	
	# Create the task scheduler COM object
	$st = New-Object -ComObject "Schedule.Service"
	# connect to remote server
	$st.Connect($_serverName, $null, $null, $null)
	
	$checkCommand = 
	{
		$rootFolder = $st.GetFolder("\")
		$searchTask = $rootFolder.GetTask($_taskName)
	
		# Check that the task was successfully registered
		if ($searchTask)
		{
			Log-Verbose "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has state: $($searchTask.State), Enabled=$($searchTask.Enabled)"
			
			# First, check if some critical arguments are the same
			[xml]$oldXml = [xml]$searchTask.Xml
			if ($oldXml.Task.Actions.Exec.Command -ine $taskXml.Task.Actions.Exec.Command)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different Execution Command.`nOld: $($oldXml.Task.Actions.Exec.Command)`nNew: $($taskXml.Task.Actions.Exec.Command)"
			}
			if ($oldXml.Task.Actions.Exec.Arguments -ine $taskXml.Task.Actions.Exec.Arguments)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different Execution Arguments.`nOld: $($oldXml.Task.Actions.Exec.Arguments)`nNew: $($taskXml.Task.Actions.Exec.Arguments)"
			}
			if ($oldXml.Task.Triggers.TimeTrigger.Repetition.Interval -ine $taskXml.Task.Triggers.TimeTrigger.Repetition.Interval)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different Repetion Interval Arguments.`nOld: $($oldXml.Task.Triggers.TimeTrigger.Repetition.Interval)`nNew: $($taskXml.Task.Triggers.TimeTrigger.Repetition.Interval)"
			}
			if ($oldXml.Task.Triggers.TimeTrigger.ExecutionTimeLimit -ine $taskXml.Task.Triggers.TimeTrigger.ExecutionTimeLimit)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different ExecutionTimeLimit Arguments.`nOld: $($oldXml.Task.Triggers.TimeTrigger.ExecutionTimeLimit)`nNew: $($taskXml.Task.TimeTrigger.Triggers.ExecutionTimeLimit)"
			}
			if ($oldXml.Task.Settings.ExecutionTimeLimit -ine $taskXml.Task.Settings.ExecutionTimeLimit)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different ExecutionTimeLimit Settings Arguments.`nOld: $($oldXml.Task.Settings.ExecutionTimeLimit)`nNew: $($taskXml.Task.Settings.ExecutionTimeLimit)"
			}		
			if ($oldXml.Task.Settings.MultipleInstancesPolicy -ine $taskXml.Task.Settings.MultipleInstancesPolicy)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different MultipleInstancesPolicy Arguments.`nOld: $($oldXml.Task.Settings.MultipleInstancesPolicy)`nNew: $($taskXml.Task.Settings.MultipleInstancesPolicy)"
			}
			if ($oldXml.Task.Settings.AllowHardTerminate -ine $taskXml.Task.Settings.AllowHardTerminate)
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has different AllowHardTerminate Arguments.`nOld: $($oldXml.Task.Settings.AllowHardTerminate)`nNew: $($taskXml.Task.Settings.AllowHardTerminate)"
			}
			# The RestartOnFailure settings are being removed
			if ( ( $oldXml.GetElementsByTagName("RestartOnFailure") | Measure-Object ).Count -gt 0 )
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' has RestartOnFailure Arguments."
			}
			
			# Now, check that it is enabled and running
			# State can be Unknown=0,Disabled=1,Queued=2,Ready=3,Running=4
			if ( !$searchTask.Enabled -or ($searchTask.State -lt 2) )
			{
				Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' is not registered properly."
			}
		}
		else
		{
			Log-Error "Is-TaskRegistered: Task '$_taskName' on server '$_serverName' was not found."
		}
	}
	
	[bool]$success = TryExecute-ScriptBlock -runCommand $checkCommand -silentOnErrors $true
	return $success
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

# Common function to wait for condition
function WaitForCondition {
    param ([ScriptBlock] $condition, [int] $seconds)
     
    $endTime = [DateTime]::Now.addseconds($seconds)
    while ([DateTime]::Now -lt $endTime)
    {
	   if ( &$condition )
	   {
	       return $true
	   }
	   Sleep-ForSeconds 2
    }
     
    # Check one last time
    if ( &$condition )
    {
        return $true
    }
    return $false
}

# Is-WriteExchangeSetupLogPresent
#	Tests if Write-ExchangeSetupLog command is available.
# Returns:
#	bool (true if Write-ExchangeSetupLog command is available, false otherwise)
function Is-WriteExchangeSetupLogPresent
{
	$checkCommand = Get-Command Write-ExchangeSetupLog -ErrorAction SilentlyContinue
	return ($checkCommand -ne $null)
}


# Sleep for the specified duration (in seconds)
function Sleep-ForSeconds ( [int]$sleepSecs )
{
	Log-Verbose "Sleeping for $sleepSecs seconds..."
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
	if (!$RunFromSetup)
	{
		$timeStamp = Get-CurrentTimeString
		Write-Verbose "$timeStamp $msg"
	}
	elseif ($script:writeToExchangeSetupLog)
	{
		$line = "[{0}] {1}" -F "ManageScheduledTask.ps1", $msg
		Write-ExchangeSetupLog -Info $line
	}
	else
	{
		$line = "[{0}] [{1}] {2}" -F $(get-date).ToString("HH:mm:ss"), "ManageScheduledTask.ps1", $msg
		add-content -Path "$($script:exLogDir)\ServiceControl.log" -Value $line
	}
}

# Common function for warning logging
function Log-Warning ( [string]$msg )
{
	if (!$RunFromSetup)
	{
		$timeStamp = Get-CurrentTimeString
		Write-Warning "$timeStamp $msg"
	}
	elseif ($script:writeToExchangeSetupLog)
	{
		$line = "[{0}] {1}" -F "ManageScheduledTask.ps1", $msg
		Write-ExchangeSetupLog -Warning $line
	}
	else
	{
		$line = "[{0}] [{1}] [Warning] {2}" -F $(get-date).ToString("HH:mm:ss"), "ManageScheduledTask.ps1", $msg
		add-content -Path "$($script:exLogDir)\ServiceControl.log" -Value $line
	}
}

# Common function for error logging
function Log-Error ( [string]$msg, [switch]$Stop)
{
	if (!$RunFromSetup)
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
	elseif ($script:writeToExchangeSetupLog)
	{
		$line = "[{0}] {1}" -F "ManageScheduledTask.ps1", $msg
		$exception = New-Object -TypeName System.Exception -ArgumentList $line
		Write-ExchangeSetupLog -Error $exception
	}
	else
	{
		$line = "[{0}] [{1}] [Error] {2}" -F $(get-date).ToString("HH:mm:ss"), "ManageScheduledTask.ps1", $msg
		add-content -Path "$($script:exLogDir)\ServiceControl.log" -Value $line
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
	Log-Error ($ManageScheduledTask_LocalizedStrings.res_0002 -f $failedCommand,$failedMessage) -Stop:$Stop
}



###################################################################
###  Entry point for the script itself
###################################################################


# Command Write-ExchangeSetupLog is not available unless when run during setup.
# So we could log stuff into a separate file.
# Must set this flag before everything else
$script:writeToExchangeSetupLog = Is-WriteExchangeSetupLogPresent

if ($RunFromSetup)
{
	# If log file folder doesn't exist, create it
	if (!(Test-Path $script:exLogDir))
	{
		New-Item $script:exLogDir -type directory	
	}
}


$Command = $MyInvocation.MyCommand
Log-Verbose "Starting: $($Command.Path)"
# The command below is useful to see what parameters are defined in this script cmdlet.
# $Command | fl Path, CommandType, Parameters, ParameterSets


if ($PSCmdlet.ParameterSetName  -eq "Install" )
{
	# Get the UNC script path on the remote server
	$remotePath = "\\" + $ServerName + "\" + $PsScriptPath.Replace(":", "$")
	if (Test-Path $remotePath)
	{
		Log-Verbose "Found script: '$remotePath'"
	}
	else
	{
		Log-Error ($ManageScheduledTask_LocalizedStrings.res_0003 -f $remotePath)
		return
	}
	
	if ($DeleteExisting)
	{
		# First delete the existing task
		Delete-RegisteredTask $ServerName $TaskName
	}
		
	# Using '&amp;' instead of '&' because the $psCommand will be passed into an XML block
	$psCommand = "&amp; '" + $PsScriptPath + "' " + $PsScriptArgs
	[bool]$success = Create-TaskUsingCOM $ServerName $TaskName $psCommand
	if ($success)
	{
		Log-Verbose "Task created successfully."
	}
	else
	{
		Log-Error $ManageScheduledTask_LocalizedStrings.res_0004
	}
}
elseif ($PSCmdlet.ParameterSetName  -eq "Uninstall" )
{
	# Delete the existing task 
	Delete-RegisteredTask $ServerName $TaskName
}
elseif ($PSCmdlet.ParameterSetName  -eq "Disable" )
{
	# Disable the existing task but don't delete it
	StopAndDelete-RegisteredTaskInternal $ServerName $TaskName $false
}
elseif ($PSCmdlet.ParameterSetName  -eq "Enable" )
{
	# Enable the existing task
	Enable-TaskUsingCOM $ServerName $TaskName
}
elseif ($PSCmdlet.ParameterSetName  -eq "TestExistence" )
{
	# returns $true or $false
	Is-TaskExisting $ServerName $TaskName
}

# SIG # Begin signature block
# MIIaxwYJKoZIhvcNAQcCoIIauDCCGrQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU4NfHe2n2tOvN1k29+NPxUq2f
# XAWgghWCMIIEwzCCA6ugAwIBAgITMwAAAHGzLoprgqofTgAAAAAAcTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTUwMzIwMTczMjAz
# WhcNMTYwNjIwMTczMjAzWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkI4RUMtMzBBNC03MTQ0MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA6pG9soj9FG8h
# NigDZjM6Zgj7W0ukq6AoNEpDMgjAhuXJPdUlvHs+YofWfe8PdFOj8ZFjiHR/6CTN
# A1DF8coAFnulObAGHDxEfvnrxLKBvBcjuv1lOBmFf8qgKf32OsALL2j04DROfW8X
# wG6Zqvp/YSXRJnDSdH3fYXNczlQqOVEDMwn4UK14x4kIttSFKj/X2B9R6u/8aF61
# wecHaDKNL3JR/gMxR1HF0utyB68glfjaavh3Z+RgmnBMq0XLfgiv5YHUV886zBN1
# nSbNoKJpULw6iJTfsFQ43ok5zYYypZAPfr/tzJQlpkGGYSbH3Td+XA3oF8o3f+gk
# tk60+Bsj6wIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFPj9I4cFlIBWzTOlQcJszAg2
# yLKiMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAC0EtMopC1n8Luqgr0xOaAT4ku0pwmbMa3DJh+i+h/xd9N1P
# pRpveJetawU4UUFynTnkGhvDbXH8cLbTzLaQWAQoP9Ye74OzFBgMlQv3pRETmMaF
# Vl7uM7QMN7WA6vUSaNkue4YIcjsUe9TZ0BZPwC8LHy3K5RvQrumEsI8LXXO4FoFA
# I1gs6mGq/r1/041acPx5zWaWZWO1BRJ24io7K+2CrJrsJ0Gnlw4jFp9ByE5tUxFA
# BMxgmdqY7Cuul/vgffW6iwD0JRd/Ynq7UVfB8PDNnBthc62VjCt2IqircDi0ASh9
# ZkJT3p/0B3xaMA6CA1n2hIa5FSVisAvSz/HblkUwggTsMIID1KADAgECAhMzAAAA
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
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBK8wggSr
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAAAymzVMhI1xOFV
# AAEAAADKMAkGBSsOAwIaBQCggcgwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFExb
# /oOQtUD+/QAssW/6tKtLfuAiMGgGCisGAQQBgjcCAQwxWjBYoDCALgBNAGEAbgBh
# AGcAZQBTAGMAaABlAGQAdQBsAGUAZABUAGEAcwBrAC4AcABzADGhJIAiaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQCG
# oAWf6ct2Iiy201bxlzz3JNUPYnef71mpRrH4XnwHB1EsUBXYxWwRJ9mID5SBxAxh
# I960Rl1CKeEMksU0m8iUXf0O0twzNegs4fMqRRNh/VMXaiSGbMmaTipKFK1caem/
# xkFoSCQnRoExJico86QazKBesBHfgKcptnrnfNimnGkx4VO3chQvuC2Spdv6ADk0
# o6c5Hh4QHNw+0IHToUWBoI/88RNnDTHUAzondCdBvg69oCyRt16PYOin2AtNS8Cq
# sRJn2Oonmop6EVZH4CtXpfvIuPFln0ecNchphXMXnaJIGi1xM5Ru+koCfkCMezuU
# ZA7zF4MVaaBk/o5KfRbQoYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGO
# MHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMT
# GE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAHGzLoprgqofTgAAAAAAcTAJ
# BgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0B
# CQUxDxcNMTUwNDEwMDI1NzQzWjAjBgkqhkiG9w0BCQQxFgQU1Wh+R749Aos6ARKv
# V0/GpjkeAt0wDQYJKoZIhvcNAQEFBQAEggEAnnLhBTl39nqyafUPyC7sTK0JBhYD
# g9UDfzHZzDPJLR+g2IZO5VK3Nsz/umX8num7U8/vGD+tH3dWNeD6BhWHJISb6jYa
# zDN9umrRDIQaXuZwCRdo5Ht3VeUkpCRBUtP/0T2+RomlVczjrp7MlV+DBQFm7QlQ
# r+RIw58a+JJ6CTq3YeatO1HrwExO2Km+phNoXC1GHsSOeUuW3NFpyTmEtUHjO3rM
# UVxFLr+eQEbanjimbwmAUDiUcnbgZzqurzV0M29Oh8rcr7a6jRQ2CKqVJOnqmjoJ
# NCtTTojahXYul9OmV1qVtUyTs40uu/xVAj0pHdydxabM6LrWrDJrFpXiVA==
# SIG # End signature block
