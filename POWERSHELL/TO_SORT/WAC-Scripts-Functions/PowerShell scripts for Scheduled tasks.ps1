function Add-WACSTScheduledTaskAction {
<#

.SYNOPSIS
Adds a new action to existing scheduled task actions.

.DESCRIPTION
Adds a new action to existing scheduled task actions.

.ROLE
Administrators

.PARAMETER taskName
    The name of the task

.PARAMETER taskPath
    The task path.

.PARAMETER actionExecute
    The name of executable to run. By default looks in System32 if Working Directory is not provided

.PARAMETER actionArguments
    The arguments for the executable.

.PARAMETER workingDirectory
    The path to working directory
#>

param (
    [parameter(Mandatory=$true)]
    [string]
    $taskName,
    [parameter(Mandatory=$true)]
    [string]
    $taskPath,
    [parameter(Mandatory=$true)]
    [string]
    $actionExecute,
    [string]
    $actionArguments,
    [string]
    $workingDirectory  
)

Import-Module ScheduledTasks

#
# Prepare action parameter bag
#
$taskActionParams = @{
    Execute = $actionExecute;
} 

if ($actionArguments) {
    $taskActionParams.Argument = $actionArguments;
}
if ($workingDirectory) {
     $taskActionParams.WorkingDirectory = $workingDirectory;
}

######################################################
#### Main script
######################################################

# Create action object
$action = New-ScheduledTaskAction @taskActionParams

$task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath
$actionsArray =  $task.Actions
$actionsArray += $action 
Set-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $actionsArray
}
## [END] Add-WACSTScheduledTaskAction ##
function Add-WACSTScheduledTaskTrigger {
 <#

.SYNOPSIS
Adds a new trigger to existing scheduled task triggers.

.DESCRIPTION
Adds a new trigger to existing scheduled task triggers.

.ROLE
Administrators

.PARAMETER taskName
    The name of the task

.PARAMETER taskDescription
    The description of the task.

.PARAMETER taskPath
    The task path.

.PARAMETER triggerAt
    The date/time to trigger the task.    

.PARAMETER triggerFrequency
    The frequency of the task occurence. Possible values Daily, Weekly, Monthly, Once, AtLogOn, AtStartup

.PARAMETER daysInterval
    The number of days interval to run task.

.PARAMETER weeklyInterval
    The number of weeks interval to run task.

.PARAMETER daysOfWeek
    The days of the week to run the task. Possible values can be an array of Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday

.PARAMETER username
    The username associated with the trigger.

.PARAMETER repetitionInterval
    The repitition interval.

.PARAMETER repetitionDuration
    The repitition duration.

.PARAMETER randomDelay
    The delay before running the trigger.
#>
 param (
    [parameter(Mandatory=$true)]
    [string]
    $taskName,
    [parameter(Mandatory=$true)]
    [string]
    $taskPath,
    [AllowNull()][System.Nullable[DateTime]]
    $triggerAt,
    [parameter(Mandatory=$true)]
    [string]
    $triggerFrequency, 
    [Int32]
    $daysInterval, 
    [Int32]
    $weeksInterval,
    [string[]]
    $daysOfWeek,
    [string]
    $username,
    [string]
    $repetitionInterval,
    [string]
    $repetitionDuration,
    [boolean]
    $stopAtDurationEnd,
    [string]
    $randomDelay,
    [string]
    $executionTimeLimit
)

Import-Module ScheduledTasks

#
# Prepare task trigger parameter bag
#
$taskTriggerParams = @{} 

if ($triggerAt -and $triggerFrequency -in ('Daily','Weekly', 'Once')) {
   $taskTriggerParams.At =  $triggerAt;
}
   
    
# Build optional switches
if ($triggerFrequency -eq 'Daily' )
{
    $taskTriggerParams.Daily = $true;
    if ($daysInterval -ne 0) 
    {
       $taskTriggerParams.DaysInterval = $daysInterval;
    }
}
elseif ($triggerFrequency -eq 'Weekly')
{
    $taskTriggerParams.Weekly = $true;
    if ($weeksInterval -ne 0) 
    {
        $taskTriggerParams.WeeksInterval = $weeksInterval;
    }
    if ($daysOfWeek -and $daysOfWeek.Length -gt 0) 
    {
        $taskTriggerParams.DaysOfWeek = $daysOfWeek;
    }
}
elseif ($triggerFrequency -eq 'Once')
{
    $taskTriggerParams.Once = $true;
}
elseif ($triggerFrequency -eq 'AtLogOn')
{
    $taskTriggerParams.AtLogOn = $true;
}
elseif ($triggerFrequency -eq 'AtStartup')
{
    $taskTriggerParams.AtStartup = $true;
}

if ($username) 
{
   $taskTriggerParams.User = $username;
}


######################################################
#### Main script
######################################################

# Create trigger object
$triggersArray = @()
$triggerNew = New-ScheduledTaskTrigger @taskTriggerParams

$task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath
$triggersArray =  $task.Triggers

Set-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Trigger $triggerNew | out-null

$task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath
$trigger = $task.Triggers[0]


if ($repetitionInterval -and $trigger.Repetition -ne $null) 
{
   $trigger.Repetition.Interval = $repetitionInterval;
}
if ($repetitionDuration -and $trigger.Repetition -ne $null) 
{
   $trigger.Repetition.Duration = $repetitionDuration;
}
if ($stopAtDurationEnd -and $trigger.Repetition -ne $null) 
{
   $trigger.Repetition.StopAtDurationEnd = $stopAtDurationEnd;
}
if($executionTimeLimit) {
 $task.Triggers[0].ExecutionTimeLimit = $executionTimeLimit;
}

if([bool]($task.Triggers[0].PSobject.Properties.name -eq "RandomDelay")) 
{
    $task.Triggers[0].RandomDelay = $randomDelay;
}

if([bool]($task.Triggers[0].PSobject.Properties.name -eq "Delay")) 
{
    $task.Triggers[0].Delay = $randomDelay;
}

$triggersArray += $trigger

Set-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Trigger $triggersArray 
}
## [END] Add-WACSTScheduledTaskTrigger ##
function Disable-WACSTScheduledTask {
<#

.SYNOPSIS
Script to disable a scheduled tasks.

.DESCRIPTION
Script to disable a scheduled tasks.

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [String]
  $taskPath,

  [Parameter(Mandatory = $true)]
  [String]
  $taskName
)
Import-Module ScheduledTasks

Disable-ScheduledTask -TaskPath $taskPath -TaskName $taskName

}
## [END] Disable-WACSTScheduledTask ##
function Enable-WACSTScheduledTask {
<#

.SYNOPSIS
Script to enable a scheduled tasks.

.DESCRIPTION
Script to enable a scheduled tasks.

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [String]
  $taskPath,

  [Parameter(Mandatory = $true)]
  [String]
  $taskName
)

Import-Module ScheduledTasks

Enable-ScheduledTask -TaskPath $taskPath -TaskName $taskName

}
## [END] Enable-WACSTScheduledTask ##
function Get-WACSTEventLogs {
<#

.SYNOPSIS
Script to get event logs and sources.

.DESCRIPTION
Script to get event logs and sources. This is used to allow user selection when creating event based triggers.

.ROLE
Readers

#>

Import-Module Microsoft.PowerShell.Diagnostics -ErrorAction SilentlyContinue

Get-WinEvent -ListLog * -ErrorAction SilentlyContinue

}
## [END] Get-WACSTEventLogs ##
function Get-WACSTScheduledTasks {
<#

.SYNOPSIS
Script to get list of scheduled tasks.

.DESCRIPTION
Script to get list of scheduled tasks.

.ROLE
Readers

#>

param (
  [Parameter(Mandatory = $false)]
  [String]
  $taskPath,

  [Parameter(Mandatory = $false)]
  [String]
  $taskName
)

Import-Module ScheduledTasks

Add-Type -AssemblyName "System.Linq"
Add-Type -AssemblyName "System.Xml"
Add-Type -AssemblyName "System.Xml.Linq"

function ConvertTo-CustomTriggerType ($trigger) {
  $customTriggerType = ''
  if ($trigger.CimClass -and $trigger.CimClass.CimClassName) {
    $cimClassName = $trigger.CimClass.CimClassName
    if ($cimClassName -eq 'MSFT_TaskTrigger') {
        $ns = [System.Xml.Linq.XNamespace]('http://schemas.microsoft.com/windows/2004/02/mit/task')
        $xml = Export-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath
        $d = [System.Xml.Linq.XDocument]::Parse($xml)
        $scheduleByMonth = $d.Descendants($ns + "ScheduleByMonth")
        if ($scheduleByMonth.Count -gt 0) {
          $customTriggerType = 'MSFT_TaskMonthlyTrigger'
        }
        else {
          $scheduleByMonthDOW = $d.Descendants($ns + "ScheduleByMonthDayOfWeek");
          if ($scheduleByMonthDOW.Count -gt 0) {
            $customTriggerType = 'MSFT_TaskMonthlyDOWTrigger'
          }
        }
    }
  }
  return $customTriggerType
}

function New-TaskWrapper
{
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline=$true)]
    $task
  )

  $task | Add-Member -MemberType NoteProperty -Name 'status' -Value $task.state.ToString()
  $info = Get-ScheduledTaskInfo $task

  $triggerCopies = @()
  for ($i=0;$i -lt $task.Triggers.Length;$i++)
  {
    $trigger = $task.Triggers[$i];
    $triggerCopy = $trigger.PSObject.Copy();
    if ($trigger -ne $null) {

        if ($trigger.StartBoundary -eq $null -or$trigger.StartBoundary -eq '')
        {
            $startDate = $null;
        }

        else
        {
            $startDate = [datetime]($trigger.StartBoundary)
        }

        $triggerCopy | Add-Member -MemberType NoteProperty -Name 'TriggerAtDate' -Value $startDate -TypeName System.DateTime

        if ($trigger.EndBoundary -eq $null -or$trigger.EndBoundary -eq '')
        {
            $endDate = $null;
        }

        else
        {
            $endDate = [datetime]($trigger.EndBoundary)
        }

        $triggerCopy | Add-Member -MemberType NoteProperty -Name 'TriggerEndDate' -Value $endDate -TypeName System.DateTime

        $customTriggerType = ConvertTo-CustomTriggerType -trigger $triggerCopy
        if ($customTriggerType) {
          $triggerCopy | Add-Member -MemberType NoteProperty -Name 'CustomParsedTriggerType' -Value $customTriggerType
        }

        $triggerCopies += $triggerCopy
    }

  }

  $task | Add-Member -MemberType NoteProperty -Name 'TriggersEx' -Value $triggerCopies

  New-Object -TypeName PSObject -Property @{

      ScheduledTask = $task
      ScheduledTaskInfo = $info
  }
}

if ($taskPath -and $taskName) {
  try
  {
    $task = Get-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop
    New-TaskWrapper $task
  }
  catch
  {
  }
} else {
    Get-ScheduledTask | ForEach-Object {
      New-TaskWrapper $_
    }
}

}
## [END] Get-WACSTScheduledTasks ##
function New-WACSTBasicTask {
<#

.SYNOPSIS
Creates and registers a new scheduled task.

.DESCRIPTION
Creates and registers a new scheduled task.

.ROLE
Administrators

.PARAMETER taskName
    The name of the task

.PARAMETER taskDescription
    The description of the task.

.PARAMETER taskPath
    The task path.

.PARAMETER taskAuthor
    The task author.

.PARAMETER triggerAt
    The date/time to trigger the task.

.PARAMETER triggerFrequency
    The frequency of the task occurence. Possible values Daily, Weekly, Monthly, Once, AtLogOn, AtStartup

.PARAMETER triggerMonthlyFrequency
    The monthly frequencty of the task occurence. Possible values Monthly (day of month), MonthlyDOW( day of week)

.PARAMETER daysInterval
    The number of days interval to run task.

.PARAMETER weeklyInterval
    The number of weeks interval to run task.

.PARAMETER daysOfWeek
    The days of the week to run the task. Possible values can be an array of Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday

.PARAMETER months
    The months of the year that the task is to run. Possible values January thru February

.PARAMETER daysOfMonth
    The specific days of the month that the task can run. Possible values 1-31 and Last. This applies when the task frequency is Monthly.

.PARAMETER weeksOfMonth
    The specific weeks of the month that the task can run. Possible values 1-4 and Last. This applies when the task frequency is MonthlyDOW.

.PARAMETER actionExecute
    The name of executable to run. By default looks in System32 if Working Directory is not provided

.PARAMETER actionArguments
    The arguments for the executable.

.PARAMETER workingDirectory
    The path to working directory
#>

param (
  [parameter(Mandatory = $true)]
  [string]
  $taskName,
  [string]
  $taskDescription,
  [parameter(Mandatory = $true)]
  [string]
  $taskPath,
  [parameter(Mandatory = $true)]
  [string]
  $taskAuthor,
  [parameter(Mandatory = $true)]
  [string]
  $triggerFrequency,
  [string]
  $triggerMonthlyFrequency,
  [AllowNull()][System.Nullable[DateTime]]
  $triggerAt,
  [Int32]
  $daysInterval,
  [Int32]
  $weeklyInterval,
  [string[]]
  $daysOfWeek,
  [string[]]
  $months = @(),
  [string[]]
  $daysOfMonth = @(),
  [string[]]
  $weeksOfMonth = @(),
  [parameter(Mandatory = $true)]
  [string]
  $actionExecute,
  [string]
  $actionArguments,
  [string]
  $workingDirectory,
  [string]
  $eventLogName,
  [string]
  $eventLogSource,
  [Int32]
  $eventLogId,
  [string]
  $userGroupControl,
  [bool]
  $highestPrivilege
  #### WIP: Password relevant elements below.
  # [string]
  # $password
)

Import-Module ScheduledTasks

##SkipCheck=true##
$Source = @"

namespace SME {

    using System;
    using System.Linq;
    using System.Xml.Linq;

    public class TaskSchedulerXml
    {
        public XNamespace ns = "http://schemas.microsoft.com/windows/2004/02/mit/task";

        public XElement CreateMonthlyTrigger(DateTime startBoundary, bool enabled, string[] months, string[] days)
        {
            var element = new XElement(ns + "CalendarTrigger",
                new XElement(ns + "StartBoundary", startBoundary.ToString("s")),
                new XElement(ns + "Enabled", enabled),
                new XElement(ns + "ScheduleByMonth",
                        new XElement(ns + "DaysOfMonth",
                            from day in days
                            select new XElement(ns + "Day", day)
                        ),
                        new XElement(ns + "Months",
                            from month in months
                            select new XElement(ns + month)
                       )
                    )
                );
            return element;
        }

        public XElement CreateMonthlyDOWTrigger(DateTime startBoundary, bool enabled, string[] months, string[] days, string[] weeks)
        {
            var element = new XElement(ns + "CalendarTrigger",
                new XElement(ns + "StartBoundary", startBoundary.ToString("s")),
                new XElement(ns + "Enabled", enabled),
                new XElement(ns + "ScheduleByMonthDayOfWeek",
                        new XElement(ns + "Weeks",
                            from week in weeks
                            select new XElement(ns + "Week", week)
                        ),
                        new XElement(ns + "DaysOfWeek",
                            from day in days
                            select new XElement(ns + day)
                        ),
                        new XElement(ns + "Months",
                            from month in months
                            select new XElement(ns + month)
                       )
                    )
                );
            return element;
        }

        public XElement CreateEventTrigger(string eventLogName, string eventLogSource, string eventLogId, bool enabled)
        {
            XNamespace ns = "http://schemas.microsoft.com/windows/2004/02/mit/task";

            var queryText = string.Format("*[System[Provider[@Name='{0}'] and EventID={1}]]", eventLogSource, eventLogId);

            var queryElement = new XElement("QueryList",
                            new XElement("Query", new XAttribute("Id", "0"), new XAttribute("Path", eventLogName),
                                new XElement("Select", new XAttribute("Path", eventLogName), queryText
                                )
                            )
                        );

            var element = new XElement(ns + "EventTrigger",
                    new XElement(ns + "Enabled", enabled),
                    new XElement(ns + "Subscription", queryElement.ToString()
                    )
                );

            return element;
        }

        public void UpdateTriggers(XElement newTrigger, XDocument d)
        {
            var triggers = d.Descendants(ns + "Triggers").FirstOrDefault();
            if (triggers != null) {
                triggers.ReplaceAll(newTrigger);
            }
        }
    }
  }

"@
##SkipCheck=false##

Add-Type -AssemblyName "System.Linq"
Add-Type -AssemblyName "System.Xml"
Add-Type -AssemblyName "System.Xml.Linq"
Add-Type -TypeDefinition $Source -Language CSharp  -ReferencedAssemblies ("System.Linq", "System.Xml", "System.Xml.Linq")

enum TriggerFrequency {
  Daily
  Weekly
  Monthly
  MonthlyDOW
  Once
  AtLogOn
  AtStartUp
  AtRegistration
  OnIdle
  OnEvent
  CustomTrigger
}

function New-ScheduledTaskXmlTemplate {
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $taskName,
    [Parameter(Mandatory = $true)]
    [string]
    $taskPath,
    [Parameter(Mandatory = $true)]
    [string]
    $taskDescription,
    [Parameter(Mandatory = $true)]
    [string]
    $taskAuthor,
    [Parameter(Mandatory = $true)]
    $taskActionParameters
  )

  # create a task as template
  $action = New-ScheduledTaskAction @taskActionParameters
  $trigger = New-ScheduledTaskTrigger -Once -At 12AM
  $settingSet = New-ScheduledTaskSettingsSet

  $principalParams = @{ }

  #### WIP: Password relevant elements below.

  # if ($password) {
  #   $principalParams += @{'User' = $userGroupControl
  #                   'Password' = $password
  #                 }
  # } else {
  #   if ($userGroupControl.EndsWith('$')) { # gMSA account specific setting
  #     $principal = Get-Principal -userId $userGroupControl -highestPrivilege $highestPrivilege -logonType "Password"
  #   } else {
  #     $principal = Get-Principal -userId $userGroupControl -highestPrivilege $highestPrivilege -logonType "Interactive"
  #   }
  #   $principalParams += @{'Principal' = $principal}
  # }

  if ($userGroupControl.EndsWith('$')) {
    # gMSA account specific setting
    $principal = Get-Principal -userId $userGroupControl -highestPrivilege $highestPrivilege -logonType "Password"
  }
  else {
    $principal = Get-Principal -userId $userGroupControl -highestPrivilege $highestPrivilege -logonType "Interactive"
  }
  $principalParams += @{'Principal' = $principal }

  Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $action -Description $taskDescription -Trigger $trigger -Settings $settingSet @principalParams
  Set-Author -taskPath $taskPath -taskName $taskName -taskAuthor $taskAuthor

  $xml = Export-ScheduledTask -TaskName $taskName -TaskPath $taskPath
  Unregister-ScheduledTask -Confirm:$false -TaskName  $taskName -TaskPath $taskPath

  return $xml
}

function Set-MonthlyTrigger {
  param (
    [Parameter(Mandatory = $true)]
    $taskXml,
    [Parameter(Mandatory = $true)]
    [DateTime]
    $startBoundary,
    [Parameter(Mandatory = $true)]
    [Boolean]
    $enabled,
    [Parameter(Mandatory = $true)]
    [string]
    $triggerMonthlyFrequency,
    [Parameter(Mandatory = $true)]
    [string[]]
    $months,
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [string[]]
    $daysOfMonth,
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    [string[]]
    $weeksOfMonth
  )

  $obj = New-Object SME.TaskSchedulerXml
  $element = $null
  if ($triggerMonthlyFrequency -eq 'Monthly') {
    $element = $obj.CreateMonthlyTrigger($startBoundary, $enabled, $months, $daysOfMonth)
  }
  elseif ( $triggerMonthlyFrequency -eq 'MonthlyDOW') {
    $element = $obj.CreateMonthlyDOWTrigger($startBoundary, $enabled, $months, $daysOfWeek, $weeksOfMonth)
  }

  $d = [System.Xml.Linq.XDocument]::Parse($taskXml)
  $obj.UpdateTriggers($element, $d)

  return $d.ToString()
}

function Set-EventTrigger {
  param (
    [Parameter(Mandatory = $true)]
    $taskXml,
    [Parameter(Mandatory = $true)]
    [Boolean]
    $enabled,
    [Parameter(Mandatory = $true)]
    [string]
    $eventLogName,
    [Parameter(Mandatory = $true)]
    [string]
    $eventLogSource,
    [Parameter(Mandatory = $true)]
    [string]
    $eventLogId
  )

  $obj = New-Object SME.TaskSchedulerXml
  $element = $obj.CreateEventTrigger($eventLogName, $eventLogSource, $eventLogId, $enabled)

  $d = [System.Xml.Linq.XDocument]::Parse($taskXml)
  $obj.UpdateTriggers($element, $d)

  return $d.ToString()
}

function Set-Author {
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $taskName,
    [Parameter(Mandatory = $true)]
    [string]
    $taskPath,
    [Parameter(Mandatory = $true)]
    [string]
    $taskAuthor
  )

  $task = Get-ScheduledTask -TaskPath $taskPath -TaskNAme $taskName
  $task.Author = $taskAuthor
  $task | Set-ScheduledTask
}

function Get-Principal {
  param (
    [string]
    $userId,
    [bool]
    $highestPrivilege,
    [string]
    $logonType
  )

  $principal = @{ }

  if ($userId) {
    $principal += @{'UserId' = $userId; }
  }

  if ($highestPrivilege) {
    $principal += @{'RunLevel' = 1 }
  }

  if ($logonType) {
    $principal += @{'LogonType' = $logonType }
  }

  return New-ScheduledTaskPrincipal @principal

}

function Set-Properties {
  param (
    [Parameter(Mandatory = $true)]
    $settings,
    [Parameter(Mandatory = $true)]
    $object
  )

  $settings.GetEnumerator() | ForEach-Object { if ($_.value) { $object[$_.key] = $_.value } }
}

function Set-ActionParameters {
  $taskActionParams = @{ }

  $settings = @{
    'Execute'          = $actionExecute
    'Argument'         = $actionArguments
    'WorkingDirectory' = $workingDirectory
  }

  Set-Properties -settings $settings -object $taskActionParams

  return $taskActionParams
}

function Set-TriggerParameters {
  $taskTriggerParams = @{ }

  switch ($triggerFrequency) {
    Daily { $taskTriggerParams.Daily = $true }
    Weekly { $taskTriggerParams.Weekly = $true }
    Monthly { $taskTriggerParams.Monthly = $true; }
    Once { $taskTriggerParams.Once = $true; }
    AtLogOn { $taskTriggerParams.AtLogOn = $true; }
    AtStartup { $taskTriggerParams.AtStartup = $true; }
  }

  $settings = @{
    'At'            = $triggerAt
    'DaysInterval'  = $daysInterval
    'WeeksInterval' = $weeklyInterval
    'DaysOfWeek'    = $daysOfWeek
  }

  Set-Properties -settings $settings -object $taskTriggerParams

  return $taskTriggerParams
}

function Test-UseXmlToCreateScheduledTask {
  return ($triggerFrequency -eq [TriggerFrequency]::Monthly) -Or ($triggerFrequency -eq [TriggerFrequency]::OnEvent)
}

#
# Prepare action parameter bag
#
$taskActionParams = Set-ActionParameters

#
# Prepare task trigger parameter bag
#
$taskTriggerParams = Set-TriggerParameters

######################################################
#### Main script
######################################################

if (-Not (Test-UseXmlToCreateScheduledTask)) {
  # Create action, trigger and default settings
  $action = New-ScheduledTaskAction @taskActionParams
  $trigger = New-ScheduledTaskTrigger @taskTriggerParams
  $settingSet = New-ScheduledTaskSettingsSet

  $principalParams = @{ }

  #### WIP: Password relevant elements below.

  # if ($password) {
  #   $principalParams += @{'User' = $userGroupControl
  #                   'Password' = $password
  #                 }
  # } else {
  #   if ($userGroupControl.EndsWith('$')) { # gMSA account specific setting
  #     $principal = Get-Principal -userId $userGroupControl -highestPrivilege $highestPrivilege -logonType "Password"
  #   } else {
  #     $principal = Get-Principal -userId $userGroupControl -highestPrivilege $highestPrivilege -logonType "Interactive"
  #   }
  #   $principalParams += @{'Principal' = $principal}
  # }

  if ($userGroupControl.EndsWith('$')) {
    # gMSA account specific setting
    $principal = Get-Principal -userId $userGroupControl -highestPrivilege $highestPrivilege -logonType "Password"
  }
  else {
    $principal = Get-Principal -userId $userGroupControl -highestPrivilege $highestPrivilege -logonType "Interactive"
  }
  $principalParams += @{'Principal' = $principal }

  Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $action -Description $taskDescription -Trigger $trigger -Settings $settingSet @principalParams
  Set-Author -taskPath $taskPath -taskName $taskName -taskAuthor $taskAuthor
}
else {

  $xml = New-ScheduledTaskXmlTemplate -taskName $taskName -taskPath $taskPath -taskDescription $taskDescription -taskAuthor $taskAuthor -taskActionParameters $taskActionParams
  $updatedXml = ''

  if ($triggerFrequency -eq [TriggerFrequency]::Monthly) {
    $updatedXml = Set-MonthlyTrigger -taskXml $xml -startBoundary $triggerAt -enabled $true -triggerMonthlyFrequency $triggerMonthlyFrequency -months $months -daysOfMonth $daysOfMonth -weeksOfMonth $weeksOfMonth
  }
  elseif ($triggerFrequency -eq [TriggerFrequency]::OnEvent) {
    $updatedXml = Set-EventTrigger -taskXml $xml -enabled $true -eventLogName $eventLogName -eventLogSource $eventLogSource -eventLogId $eventLogId
  }

  Register-ScheduledTask -Xml $updatedXml -TaskName  $taskName -TaskPath $taskPath
}

}
## [END] New-WACSTBasicTask ##
function Remove-WACSTScheduledTask {
<#

.SYNOPSIS
Script to delete a scheduled tasks.

.DESCRIPTION
Script to delete a scheduled tasks.

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [String]
  $taskPath,

  [Parameter(Mandatory = $true)]
  [String]
  $taskName
)

Import-Module ScheduledTasks

ScheduledTasks\Unregister-ScheduledTask -TaskPath $taskPath -TaskName $taskName -Confirm:$false

}
## [END] Remove-WACSTScheduledTask ##
function Remove-WACSTScheduledTaskAction {
<#

.SYNOPSIS
Removes action from scheduled task actions.

.DESCRIPTION
Removes action from scheduled task actions.

.ROLE
Administrators

.PARAMETER taskName
    The name of the task

.PARAMETER taskPath
    The task path.

.PARAMETER actionExecute
    The name of executable to run. By default looks in System32 if Working Directory is not provided

.PARAMETER actionArguments
    The arguments for the executable.

.PARAMETER workingDirectory
    The path to working directory
#>

param (
    [parameter(Mandatory=$true)]
    [string]
    $taskName,
    [parameter(Mandatory=$true)]
    [string]
    $taskPath,
    [parameter(Mandatory=$true)]
    [string]
    $actionExecute,
    [string]
    $actionArguments,
    [string]
    $workingDirectory
)

Import-Module ScheduledTasks


######################################################
#### Main script
######################################################

$task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath
$actionsArray =  @()

$task.Actions| ForEach-Object {
    $matched = $true;  
  
    if( -not ([string]::IsNullOrEmpty($_.Arguments) -and [string]::IsNullOrEmpty($actionArguments)))
    {
        if ($_.Arguments -ne $actionArguments)
        {
            $matched = $false;
        }
    }

    $workingDirectoryMatched  = $true;
    if( -not ([string]::IsNullOrEmpty($_.WorkingDirectory) -and [string]::IsNullOrEmpty($workingDirectory)))
    {
        if ($_.WorkingDirectory -ne $workingDirectory)
        {
            $matched = $false;
        }
    }

    $executeMatched  = $true;
    if ($_.Execute -ne $actionExecute) 
    {
          $matched = $false;
    }

    if (-not ($matched))
    {
        $actionsArray += $_;
    }
}


Set-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $actionsArray
}
## [END] Remove-WACSTScheduledTaskAction ##
function Set-WACSTScheduledTaskConditions {
<#

.SYNOPSIS
Set/modify scheduled task setting set.

.DESCRIPTION
Set/modify scheduled task setting set.

.ROLE
Administrators

.PARAMETER taskName
    The name of the task

.PARAMETER taskPath
    The task path.

.PARAMETER dontStopOnIdleEnd
    Indicates that Task Scheduler does not terminate the task if the idle condition ends before the task is completed.
    
.PARAMETER idleDurationInMins
    Specifies the amount of time that the computer must be in an idle state before Task Scheduler runs the task.
    
.PARAMETER idleWaitTimeoutInMins
   Specifies the amount of time that Task Scheduler waits for an idle condition to occur before timing out.
    
.PARAMETER restartOnIdle
   Indicates that Task Scheduler restarts the task when the computer cycles into an idle condition more than once.
    
.PARAMETER runOnlyIfIdle
    Indicates that Task Scheduler runs the task only when the computer is idle.
    
.PARAMETER allowStartIfOnBatteries
    Indicates that Task Scheduler starts if the computer is running on battery power.
    
.PARAMETER dontStopIfGoingOnBatteries
    Indicates that the task does not stop if the computer switches to battery power.

.PARAMETER runOnlyIfNetworkAvailable
    Indicates that Task Scheduler runs the task only when a network is available. Task Scheduler uses the NetworkID parameter and NetworkName parameter that you specify in this cmdlet to determine if the network is available.

.PARAMETER networkId
    Specifies the ID of a network profile that Task Scheduler uses to determine if the task can run. You must specify the ID of a network if you specify the RunOnlyIfNetworkAvailable parameter.

.PARAMETER networkName
   Specifies the name of a network profile that Task Scheduler uses to determine if the task can run. The Task Scheduler UI uses this setting for display purposes. Specify a network name if you specify the RunOnlyIfNetworkAvailable parameter.

#>

param (
    [parameter(Mandatory=$true)]
    [string]
    $taskName,
    [parameter(Mandatory=$true)]
    [string]
    $taskPath,
    [Boolean]
    $stopOnIdleEnd,
    [string]
    $idleDuration,
    [string]
    $idleWaitTimeout,
    [Boolean]
    $restartOnIdle,
    [Boolean]
    $runOnlyIfIdle,
    [Boolean]
    $disallowStartIfOnBatteries,
    [Boolean]
    $stopIfGoingOnBatteries,
    [Boolean]
    $wakeToRun
)

Import-Module ScheduledTasks

$task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath;

# Idle related conditions.
$task.settings.RunOnlyIfIdle = $runOnlyIfIdle;

$task.Settings.IdleSettings.IdleDuration = $idleDuration;
$task.Settings.IdleSettings.WaitTimeout = $idleWaitTimeout;

$task.Settings.IdleSettings.RestartOnIdle = $restartOnIdle;
$task.Settings.IdleSettings.StopOnIdleEnd = $stopOnIdleEnd;

# Power related condition.
$task.Settings.DisallowStartIfOnBatteries = $disallowStartIfOnBatteries;

$task.Settings.StopIfGoingOnBatteries = $stopIfGoingOnBatteries;

$task.Settings.WakeToRun = $wakeToRun;

$task | Set-ScheduledTask;
}
## [END] Set-WACSTScheduledTaskConditions ##
function Set-WACSTScheduledTaskGeneralSettings {
<#

.SYNOPSIS
Creates and registers a new scheduled task.

.DESCRIPTION
Creates and registers a new scheduled task.

.ROLE
Administrators

.PARAMETER taskName
    The name of the task

.PARAMETER taskDescription
    The description of the task.

.PARAMETER taskPath
    The task path.

.PARAMETER username
    The username to use to run the task.
#>

param (
    [parameter(Mandatory=$true)]
    [string]
    $taskName,
    [string]
    $taskDescription,
    [parameter(Mandatory=$true)]
    [string]
    $taskPath,
    [string]
    $username
)

Import-Module ScheduledTasks

######################################################
#### Main script
######################################################

$task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath
if($task) {
    
    $task.Description = $taskDescription;
  
    if ($username)
    {
        $task | Set-ScheduledTask -User $username ;
    } 
    else 
    {
        $task | Set-ScheduledTask
    }
}
}
## [END] Set-WACSTScheduledTaskGeneralSettings ##
function Set-WACSTScheduledTaskSecurity {
<#

.SYNOPSIS
Updates the security used to run a scheduled task.

.DESCRIPTION
Set which user or group should run the task. If blank, run with highest privileges.

.ROLE
Administrators

.PARAMETER taskName
    The name of the task.

.PARAMETER taskPath
    The task path.

.PARAMETER username
    The username to use to run the task.

.PARAMETER password
    The password to use to validate the username.

.PARAMETER highestPrivilege
    Indicates whether to run the task with the highest privileges of an account.

.PARAMETER doNotStorePassword
    Indicates whether to store the password of an account if the task runs even when an account is not logged in.

.PARAMETER runAnytime
    Indicates whether to run the task regardless of if the account is logged in.

#>

param (
  [parameter(Mandatory = $true)]
  [string]
  $taskName,
  [string]
  $taskPath,
  [string]
  $username,
  [bool]
  $highestPrivilege
  #### WIP: Password relevant elements below.
  # [string]
  # $password,
  # [bool]
  # $doNotStorePassword
)

Import-Module ScheduledTasks

######################################################
#### Main script
######################################################

$task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath;
$isGMSA = $username.EndsWith('$');

$principal = $task.Principal;
$principal.UserId = $username;

if (($password) -or $isGMSA) {
  $principal.LogonType = 1;
}
else {
  $principal.LogonType = 3;
}

if ($highestPrivilege) {
  $principal.RunLevel = 1;
}
else {
  $principal.RunLevel = 0;
}

# Must re-register the task under the username and password under Password/S4U logons
if (($principal.LogonType -le 2) -and !$isGMSA) {
  $taskParams = @{'Settings' = $task.Settings;
    'Principal'              = $principal;
  }

  if ($task.Actions) {
    $taskParams += @{'Action' = $task.Actions; }
  }

  if ($task.Triggers) {
    $taskParams += @{'Trigger' = $task.Triggers; }
  }

  if ($task.Description) {
    $taskParams += @{'Description' = $task.Description; }
  }

  $newTask = New-ScheduledTask @taskParams;

  # Register some task to validate credentials, then "rename" (i.e. re-register the actual task)
  $randomFileName = [System.IO.Path]::GetRandomFileName();

  Register-ScheduledTask -TaskName $randomFileName -TaskPath $taskPath -InputObject $newTask -User $username -Password $password -ErrorAction Stop;
  Unregister-ScheduledTask -TaskName $randomFileName -taskPath $taskPath -Confirm:$false;

  Unregister-ScheduledTask -TaskName $taskName -taskPath $taskPath -Confirm:$false;
  Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -InputObject $newTask -User $username -Password $password -ErrorAction Stop;

}
else {
  Set-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Principal $principal;
}

}
## [END] Set-WACSTScheduledTaskSecurity ##
function Set-WACSTScheduledTaskSettingsSet {
<#

.SYNOPSIS
Set/modify scheduled task setting set.

.DESCRIPTION
Set/modify scheduled task setting set.

.ROLE
Administrators

.PARAMETER taskName
    The name of the task

.PARAMETER taskPath
    The task path.

.PARAMETER disallowDemandStart
    Indicates that the task cannot be started by using either the Run command or the Context menu.

.PARAMETER startWhenAvailable
    Indicates that Task Scheduler can start the task at any time after its scheduled time has passed.

.PARAMETER executionTimeLimitInMins
   Specifies the amount of time that Task Scheduler is allowed to complete the task.

.PARAMETER restartIntervalInMins
    Specifies the amount of time between Task Scheduler attempts to restart the task.

.PARAMETER restartCount
    Specifies the number of times that Task Scheduler attempts to restart the task.

.PARAMETER deleteExpiredTaskAfterInMins
    Specifies the amount of time that Task Scheduler waits before deleting the task after it expires.

.PARAMETER multipleInstances
   Specifies the policy that defines how Task Scheduler handles multiple instances of the task. Possible Enum values Parallel, Queue, IgnoreNew

.PARAMETER disallowHardTerminate
   Indicates that the task cannot be terminated by using TerminateProcess.
#>

param (
    [parameter(Mandatory=$true)]
    [string]
    $taskName,
    [parameter(Mandatory=$true)]
    [string]
    $taskPath,
    [Boolean]
    $allowDemandStart,
    [Boolean]
    $allowHardTerminate,
    [Boolean]
    $startWhenAvailable, 
    [string]
    $executionTimeLimit, 
    [string]
    $restartInterval, 
    [Int32]
    $restartCount, 
    [string]
    $deleteExpiredTaskAfter,
    [Int32]
    $multipleInstances  #Parallel, Queue, IgnoreNew
    
)

Import-Module ScheduledTasks

#
# Prepare action parameter bag
#

$task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath;

$task.settings.AllowDemandStart =  $allowDemandStart;
$task.settings.AllowHardTerminate = $allowHardTerminate;

$task.settings.StartWhenAvailable = $startWhenAvailable;

if ($executionTimeLimit -eq $null -or $executionTimeLimit -eq '') {
    $task.settings.ExecutionTimeLimit = 'PT0S';
} 
else 
{
    $task.settings.ExecutionTimeLimit = $executionTimeLimit;
} 

if ($restartInterval -eq $null -or $restartInterval -eq '') {
    $task.settings.RestartInterval = $null;
} 
else
{
    $task.settings.RestartInterval = $restartInterval;
} 

if ($restartCount -gt 0) {
    $task.settings.RestartCount = $restartCount;
}
<#if ($deleteExpiredTaskAfter -eq '' -or $deleteExpiredTaskAfter -eq $null) {
    $task.settings.DeleteExpiredTaskAfter = $null;
}
else 
{
    $task.settings.DeleteExpiredTaskAfter = $deleteExpiredTaskAfter;
}#>

if ($multipleInstances) {
    $task.settings.MultipleInstances = $multipleInstances;
}

$task | Set-ScheduledTask ;
}
## [END] Set-WACSTScheduledTaskSettingsSet ##
function Start-WACSTScheduledTask {
<#

.SYNOPSIS
Script to start a scheduled tasks.

.DESCRIPTION
Script to start a scheduled tasks.

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [String]
  $taskPath,

  [Parameter(Mandatory = $true)]
  [String]
  $taskName
)

Import-Module ScheduledTasks

Get-ScheduledTask -TaskPath $taskPath -TaskName $taskName | ScheduledTasks\Start-ScheduledTask

}
## [END] Start-WACSTScheduledTask ##
function Stop-WACSTScheduledTask {
<#

.SYNOPSIS
Script to stop a scheduled tasks.

.DESCRIPTION
Script to stop a scheduled tasks.

.ROLE
Administrators

#>

param (
  [Parameter(Mandatory = $true)]
  [String]
  $taskPath,

  [Parameter(Mandatory = $true)]
  [String]
  $taskName
)

Import-Module ScheduledTasks

Get-ScheduledTask -TaskPath $taskPath -TaskName $taskName | ScheduledTasks\Stop-ScheduledTask

}
## [END] Stop-WACSTScheduledTask ##
function Update-WACSTScheduledTaskAction {
<#

.SYNOPSIS
Updates existing scheduled task action.

.DESCRIPTION
Updates existing scheduled task action.

.ROLE
Administrators

.PARAMETER taskName
    The name of the task

.PARAMETER taskPath
    The task path.

.PARAMETER oldActionExecute
    The name of executable to run. By default looks in System32 if Working Directory is not provided

.PARAMETER newActionExecute
    The name of executable to run. By default looks in System32 if Working Directory is not provided

.PARAMETER oldActionArguments
    The arguments for the executable.

.PARAMETER newActionArguments
    The arguments for the executable.

.PARAMETER oldWorkingDirectory
    The path to working directory

.PARAMETER newWorkingDirectory
    The path to working directory
#>

param (
    [parameter(Mandatory=$true)]
    [string]
    $taskName,
    [parameter(Mandatory=$true)]
    [string]
    $taskPath,
    [parameter(Mandatory=$true)]
    [string]
    $newActionExecute,
    [parameter(Mandatory=$true)]
    [string]
    $oldActionExecute,
    [string]
    $newActionArguments,
    [string]
    $oldActionArguments,
    [string]
    $newWorkingDirectory,
    [string]
    $oldWorkingDirectory
)

Import-Module ScheduledTasks


######################################################
#### Main script
######################################################

$task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath
$actionsArray = $task.Actions

foreach ($action in $actionsArray) {
    $argMatched = $true;
    if( -not ([string]::IsNullOrEmpty($action.Arguments) -and [string]::IsNullOrEmpty($oldActionArguments)))
    {
        if ($action.Arguments -ne $oldActionArguments)
        {
            $argMatched = $false;
        }
    }

    $workingDirectoryMatched  = $true;
    if( -not ([string]::IsNullOrEmpty($action.WorkingDirectory) -and [string]::IsNullOrEmpty($oldWorkingDirectory)))
    {
        if ($action.WorkingDirectory -ne $oldWorkingDirectory)
        {
            $workingDirectoryMatched = $false;
        }
    }

    $executeMatched  = $true;
    if ($action.Execute -ne $oldActionExecute) 
    {
          $executeMatched = $false;
    }

    if ($argMatched -and $executeMatched -and $workingDirectoryMatched)
    {
        $action.Execute = $newActionExecute;
        $action.Arguments = $newActionArguments;
        $action.WorkingDirectory = $newWorkingDirectory;
        break
    }
}


Set-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $actionsArray
}
## [END] Update-WACSTScheduledTaskAction ##
function Update-WACSTScheduledTaskTrigger {
 <#

.SYNOPSIS
Adds a new trigger to existing scheduled task triggers.

.DESCRIPTION
Adds a new trigger to existing scheduled task triggers.

.ROLE
Administrators

.PARAMETER taskName
    The name of the task

.PARAMETER taskPath
    The task path.

.PARAMETER triggerClassName
    The cim class Name for Trigger being edited.

.PARAMETER triggersToCreate
    Collections of triggers to create/edit, should be of same type. The script will preserve any other trigger than cim class specified in triggerClassName. 
    This is done because individual triggers can not be identified by Id. Everytime update to any trigger is made we recreate all triggers that are of the same type supplied by user in triggers to create collection.
#>
 param (
    [parameter(Mandatory=$true)]
    [string]
    $taskName,
    [parameter(Mandatory=$true)]
    [string]
    $taskPath,
    [string]
    $triggerClassName,
    [object[]]
    $triggersToCreate
)

Import-Module ScheduledTasks

######################################################
#### Functions
######################################################


function Create-Trigger 
 {
    Param (
    [object]
    $trigger
    )

    if($trigger) 
    {
        #
        # Prepare task trigger parameter bag
        #
        $taskTriggerParams = @{} 
        # Parameter is not required while creating Logon trigger /startup Trigger
        if ($trigger.triggerAt -and $trigger.triggerFrequency -in ('Daily','Weekly', 'Once')) {
           $taskTriggerParams.At =  $trigger.triggerAt;
        }
   
    
        # Build optional switches
        if ($trigger.triggerFrequency -eq 'Daily')
        {
            $taskTriggerParams.Daily = $true;
            
            if ($trigger.daysInterval -and $trigger.daysInterval -ne 0) 
            {
               $taskTriggerParams.DaysInterval = $trigger.daysInterval;
            }
        }
        elseif ($trigger.triggerFrequency -eq 'Weekly')
        {
            $taskTriggerParams.Weekly = $true;
            if ($trigger.weeksInterval -and $trigger.weeksInterval -ne 0) 
            {
               $taskTriggerParams.WeeksInterval = $trigger.weeksInterval;
            }
            if ($trigger.daysOfWeek) 
            {
               $taskTriggerParams.DaysOfWeek = $trigger.daysOfWeek;
            }
        }
        elseif ($trigger.triggerFrequency -eq 'Once')
        {
            $taskTriggerParams.Once = $true;
        }
        elseif ($trigger.triggerFrequency -eq 'AtLogOn')
        {
            $taskTriggerParams.AtLogOn = $true;
        }
        elseif ($trigger.triggerFrequency -eq 'AtStartup')
        {
            $taskTriggerParams.AtStartup = $true;
        }
        
        if ($trigger.username) 
        {
           $taskTriggerParams.User = $trigger.username;
        }


        # Create trigger object
        $triggerNew = New-ScheduledTaskTrigger @taskTriggerParams

        $task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath
       
        Set-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Trigger $triggerNew | out-null

        $task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath
     

        if ($trigger.repetitionInterval -and $task.Triggers[0].Repetition -ne $null) 
        {
           $task.Triggers[0].Repetition.Interval = $trigger.repetitionInterval;
        }
        if ($trigger.repetitionDuration -and $task.Triggers[0].Repetition -ne $null) 
        {
           $task.Triggers[0].Repetition.Duration = $trigger.repetitionDuration;
        }
        if ($trigger.stopAtDurationEnd -and $task.Triggers[0].Repetition -ne $null) 
        {
           $task.Triggers[0].Repetition.StopAtDurationEnd = $trigger.stopAtDurationEnd;
        }
        if($trigger.executionTimeLimit) 
        {
            $task.Triggers[0].ExecutionTimeLimit = $trigger.executionTimeLimit;
        }
        if($trigger.randomDelay -ne '')
        {
            if([bool]($task.Triggers[0].PSobject.Properties.name -eq "RandomDelay")) 
            {
                $task.Triggers[0].RandomDelay = $trigger.randomDelay;
            }

            if([bool]($task.Triggers[0].PSobject.Properties.name -eq "Delay")) 
            {
                $task.Triggers[0].Delay = $trigger.randomDelay;
            }
        }

        if($trigger.enabled -ne $null) 
        {
            $task.Triggers[0].Enabled = $trigger.enabled;
        }

        if($trigger.endBoundary -and $trigger.endBoundary -ne '') 
        {
            $date = [datetime]($trigger.endBoundary);
            $task.Triggers[0].EndBoundary = $date.ToString("yyyy-MM-ddTHH:mm:sszzz"); #convert date to specific string.
        }

        # Activation date is also stored in StartBoundary for Logon/Startup triggers. Setting it in appropriate context
        if($trigger.triggerAt -ne '' -and $trigger.triggerAt -ne $null -and $trigger.triggerFrequency -in ('AtLogOn','AtStartup')) 
        {
            $date = [datetime]($trigger.triggerAt);
            $task.Triggers[0].StartBoundary = $date.ToString("yyyy-MM-ddTHH:mm:sszzz"); #convert date to specific string.
        }


        return  $task.Triggers[0];
       } # end if
 }

######################################################
#### Main script
######################################################

$task = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath
$triggers = $task.Triggers;
$allTriggers = @()
try {

    foreach ($t in $triggers)
    {
        # Preserve all the existing triggers which are of different type then the modified trigger type.
        if ($t.CimClass.CimClassName -ne $triggerClassName) 
        {
            $allTriggers += $t;
        } 
    }

     # Once all other triggers are preserved, recreate the ones passed on by the UI
     foreach ($t in $triggersToCreate)
     {
        $newTrigger = Create-Trigger -trigger $t
        $allTriggers += $newTrigger;
     }

    Set-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Trigger $allTriggers
} 
catch 
{
     Set-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Trigger $triggers
     throw $_.Exception
}

}
## [END] Update-WACSTScheduledTaskTrigger ##
function Get-WACSTCimWin32LogicalDisk {
<#

.SYNOPSIS
Gets Win32_LogicalDisk object.

.DESCRIPTION
Gets Win32_LogicalDisk object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_LogicalDisk

}
## [END] Get-WACSTCimWin32LogicalDisk ##
function Get-WACSTCimWin32NetworkAdapter {
<#

.SYNOPSIS
Gets Win32_NetworkAdapter object.

.DESCRIPTION
Gets Win32_NetworkAdapter object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_NetworkAdapter

}
## [END] Get-WACSTCimWin32NetworkAdapter ##
function Get-WACSTCimWin32PhysicalMemory {
<#

.SYNOPSIS
Gets Win32_PhysicalMemory object.

.DESCRIPTION
Gets Win32_PhysicalMemory object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_PhysicalMemory

}
## [END] Get-WACSTCimWin32PhysicalMemory ##
function Get-WACSTCimWin32Processor {
<#

.SYNOPSIS
Gets Win32_Processor object.

.DESCRIPTION
Gets Win32_Processor object.

.ROLE
Readers

#>
##SkipCheck=true##


import-module CimCmdlets

Get-CimInstance -Namespace root/cimv2 -ClassName Win32_Processor

}
## [END] Get-WACSTCimWin32Processor ##
function Get-WACSTClusterInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a cluster.

.DESCRIPTION
Retrieves the inventory data for a cluster.

.ROLE
Readers

#>

Import-Module CimCmdlets -ErrorAction SilentlyContinue

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
Import-Module FailoverClusters -ErrorAction SilentlyContinue

Import-Module Storage -ErrorAction SilentlyContinue
<#

.SYNOPSIS
Get the name of this computer.

.DESCRIPTION
Get the best available name for this computer.  The FQDN is preferred, but when not avaialble
the NetBIOS name will be used instead.

#>

function getComputerName() {
    $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object Name, DNSHostName

    if ($computerSystem) {
        $computerName = $computerSystem.DNSHostName

        if ($null -eq $computerName) {
            $computerName = $computerSystem.Name
        }

        return $computerName
    }

    return $null
}

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell cmdlets installed on this server?

#>

function getIsClusterCmdletAvailable() {
    $cmdlet = Get-Command "Get-Cluster" -ErrorAction SilentlyContinue

    return !!$cmdlet
}

<#

.SYNOPSIS
Get the MSCluster Cluster CIM instance from this server.

.DESCRIPTION
Get the MSCluster Cluster CIM instance from this server.

#>
function getClusterCimInstance() {
    $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue

    if ($namespace) {
        return Get-CimInstance -Namespace root/mscluster MSCluster_Cluster -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object fqdn, S2DEnabled
    }

    return $null
}


<#

.SYNOPSIS
Determines if the current cluster supports Failover Clusters Time Series Database.

.DESCRIPTION
Use the existance of the path value of cmdlet Get-StorageHealthSetting to determine if TSDB
is supported or not.

#>
function getClusterPerformanceHistoryPath() {
    $storageSubsystem = Get-StorageSubSystem clus* -ErrorAction SilentlyContinue
    $storageHealthSettings = Get-StorageHealthSetting -InputObject $storageSubsystem -Name "System.PerformanceHistory.Path" -ErrorAction SilentlyContinue

    return $null -ne $storageHealthSettings
}

<#

.SYNOPSIS
Get some basic information about the cluster from the cluster.

.DESCRIPTION
Get the needed cluster properties from the cluster.

#>
function getClusterInfo() {
    $returnValues = @{}

    $returnValues.Fqdn = $null
    $returnValues.isS2DEnabled = $false
    $returnValues.isTsdbEnabled = $false

    $cluster = getClusterCimInstance
    if ($cluster) {
        $returnValues.Fqdn = $cluster.fqdn
        $isS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -eq 1)
        $returnValues.isS2DEnabled = $isS2dEnabled

        if ($isS2DEnabled) {
            $returnValues.isTsdbEnabled = getClusterPerformanceHistoryPath
        } else {
            $returnValues.isTsdbEnabled = $false
        }
    }

    return $returnValues
}

<#

.SYNOPSIS
Are the cluster PowerShell Health cmdlets installed on this server?

.DESCRIPTION
Are the cluster PowerShell Health cmdlets installed on this server?

s#>
function getisClusterHealthCmdletAvailable() {
    $cmdlet = Get-Command -Name "Get-HealthFault" -ErrorAction SilentlyContinue

    return !!$cmdlet
}
<#

.SYNOPSIS
Are the Britannica (sddc management resources) available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) available on the cluster?

#>
function getIsBritannicaEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual machine available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual machine available on the cluster?

#>
function getIsBritannicaVirtualMachineEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualMachine -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Are the Britannica (sddc management resources) virtual switch available on the cluster?

.DESCRIPTION
Are the Britannica (sddc management resources) virtual switch available on the cluster?

#>
function getIsBritannicaVirtualSwitchEnabled() {
    return $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_VirtualSwitch -ErrorAction SilentlyContinue)
}

###########################################################################
# main()
###########################################################################

$clusterInfo = getClusterInfo

$result = New-Object PSObject

$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $clusterInfo.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2DEnabled' -Value $clusterInfo.isS2DEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsTsdbEnabled' -Value $clusterInfo.isTsdbEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterHealthCmdletAvailable' -Value (getIsClusterHealthCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value (getIsBritannicaEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualMachineEnabled' -Value (getIsBritannicaVirtualMachineEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaVirtualSwitchEnabled' -Value (getIsBritannicaVirtualSwitchEnabled)
$result | Add-Member -MemberType NoteProperty -Name 'IsClusterCmdletAvailable' -Value (getIsClusterCmdletAvailable)
$result | Add-Member -MemberType NoteProperty -Name 'CurrentClusterNode' -Value (getComputerName)

$result

}
## [END] Get-WACSTClusterInventory ##
function Get-WACSTClusterNodes {
<#

.SYNOPSIS
Retrieves the inventory data for cluster nodes in a particular cluster.

.DESCRIPTION
Retrieves the inventory data for cluster nodes in a particular cluster.

.ROLE
Readers

#>

import-module CimCmdlets

# JEA code requires to pre-import the module (this is slow on failover cluster environment.)
import-module FailoverClusters -ErrorAction SilentlyContinue

###############################################################################
# Constants
###############################################################################

Set-Variable -Name LogName -Option Constant -Value "Microsoft-ServerManagementExperience" -ErrorAction SilentlyContinue
Set-Variable -Name LogSource -Option Constant -Value "SMEScripts" -ErrorAction SilentlyContinue
Set-Variable -Name ScriptName -Option Constant -Value $MyInvocation.ScriptName -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Are the cluster PowerShell cmdlets installed?

.DESCRIPTION
Use the Get-Command cmdlet to quickly test if the cluster PowerShell cmdlets
are installed on this server.

#>

function getClusterPowerShellSupport() {
    $cmdletInfo = Get-Command 'Get-ClusterNode' -ErrorAction SilentlyContinue

    return $cmdletInfo -and $cmdletInfo.Name -eq "Get-ClusterNode"
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster CIM provider.

.DESCRIPTION
When the cluster PowerShell cmdlets are not available fallback to using
the cluster CIM provider to get the needed information.

#>

function getClusterNodeCimInstances() {
    # Change the WMI property NodeDrainStatus to DrainStatus to match the PS cmdlet output.
    return Get-CimInstance -Namespace root/mscluster MSCluster_Node -ErrorAction SilentlyContinue | `
        Microsoft.PowerShell.Utility\Select-Object @{Name="DrainStatus"; Expression={$_.NodeDrainStatus}}, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Get the cluster nodes using the cluster PowerShell cmdlets.

.DESCRIPTION
When the cluster PowerShell cmdlets are available use this preferred function.

#>

function getClusterNodePsInstances() {
    return Get-ClusterNode -ErrorAction SilentlyContinue | Microsoft.PowerShell.Utility\Select-Object DrainStatus, DynamicWeight, Name, NodeWeight, FaultDomain, State
}

<#

.SYNOPSIS
Use DNS services to get the FQDN of the cluster NetBIOS name.

.DESCRIPTION
Use DNS services to get the FQDN of the cluster NetBIOS name.

.Notes
It is encouraged that the caller add their approprate -ErrorAction when
calling this function.

#>

function getClusterNodeFqdn([string]$clusterNodeName) {
    return ([System.Net.Dns]::GetHostEntry($clusterNodeName)).HostName
}

<#

.SYNOPSIS
Writes message to event log as warning.

.DESCRIPTION
Writes message to event log as warning.

#>

function writeToEventLog([string]$message) {
    Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue
    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Warning `
        -Message $message  -ErrorAction SilentlyContinue
}

<#

.SYNOPSIS
Get the cluster nodes.

.DESCRIPTION
When the cluster PowerShell cmdlets are available get the information about the cluster nodes
using PowerShell.  When the cmdlets are not available use the Cluster CIM provider.

#>

function getClusterNodes() {
    $isClusterCmdletAvailable = getClusterPowerShellSupport

    if ($isClusterCmdletAvailable) {
        $clusterNodes = getClusterNodePsInstances
    } else {
        $clusterNodes = getClusterNodeCimInstances
    }

    $clusterNodeMap = @{}

    foreach ($clusterNode in $clusterNodes) {
        $clusterNodeName = $clusterNode.Name.ToLower()
        try 
        {
            $clusterNodeFqdn = getClusterNodeFqdn $clusterNodeName -ErrorAction SilentlyContinue
        }
        catch 
        {
            $clusterNodeFqdn = $clusterNodeName
            writeToEventLog "[$ScriptName]: The fqdn for node '$clusterNodeName' could not be obtained. Defaulting to machine name '$clusterNodeName'"
        }

        $clusterNodeResult = New-Object PSObject

        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FullyQualifiedDomainName' -Value $clusterNodeFqdn
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'Name' -Value $clusterNodeName
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DynamicWeight' -Value $clusterNode.DynamicWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'NodeWeight' -Value $clusterNode.NodeWeight
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'FaultDomain' -Value $clusterNode.FaultDomain
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'State' -Value $clusterNode.State
        $clusterNodeResult | Add-Member -MemberType NoteProperty -Name 'DrainStatus' -Value $clusterNode.DrainStatus

        $clusterNodeMap.Add($clusterNodeName, $clusterNodeResult)
    }

    return $clusterNodeMap
}

###########################################################################
# main()
###########################################################################

getClusterNodes

}
## [END] Get-WACSTClusterNodes ##
function Get-WACSTDecryptedDataFromNode {
<#

.SYNOPSIS
Gets data after decrypting it on a node.

.DESCRIPTION
Decrypts data on node using a cached RSAProvider used during encryption within 3 minutes of encryption and returns the decrypted data.
This script should be imported or copied directly to other scripts, do not send the returned data as an argument to other scripts.

.PARAMETER encryptedData
Encrypted data to be decrypted (String).

.ROLE
Readers

#>
param (
  [Parameter(Mandatory = $true)]
  [String]
  $encryptedData
)

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function DecryptDataWithJWKOnNode {
  if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue) {
    $rsaProvider = (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    $decryptedBytes = $rsaProvider.Decrypt([Convert]::FromBase64String($encryptedData), [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)
    return [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
  }
  # If you copy this script directly to another, you can get rid of the throw statement and add custom error handling logic such as "Write-Error"
  throw [System.InvalidOperationException] "Password decryption failed. RSACryptoServiceProvider Instance not found"
}

}
## [END] Get-WACSTDecryptedDataFromNode ##
function Get-WACSTEncryptionJWKOnNode {
<#

.SYNOPSIS
Gets encrytion JSON web key from node.

.DESCRIPTION
Gets encrytion JSON web key from node.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

New-Variable -Name rsaProviderInstanceName -Value "RSA" -Option Constant

function Get-RSAProvider
{
    if(Get-Variable -Scope Global -Name $rsaProviderInstanceName -EA SilentlyContinue)
    {
        return (Get-Variable -Scope Global -Name $rsaProviderInstanceName).Value
    }

    $Global:RSA = New-Object System.Security.Cryptography.RSACryptoServiceProvider -ArgumentList 4096
    return $RSA
}

function Get-JsonWebKey
{
    $rsaProvider = Get-RSAProvider
    $parameters = $rsaProvider.ExportParameters($false)
    return [PSCustomObject]@{
        kty = 'RSA'
        alg = 'RSA-OAEP'
        e = [Convert]::ToBase64String($parameters.Exponent)
        n = [Convert]::ToBase64String($parameters.Modulus).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }
}

$jwk = Get-JsonWebKey
ConvertTo-Json $jwk -Compress

}
## [END] Get-WACSTEncryptionJWKOnNode ##
function Get-WACSTServerInventory {
<#

.SYNOPSIS
Retrieves the inventory data for a server.

.DESCRIPTION
Retrieves the inventory data for a server.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0

Import-Module CimCmdlets

Import-Module Storage -ErrorAction SilentlyContinue

<#

.SYNOPSIS
Converts an arbitrary version string into just 'Major.Minor'

.DESCRIPTION
To make OS version comparisons we only want to compare the major and
minor version.  Build number and/os CSD are not interesting.

#>

function convertOsVersion([string]$osVersion) {
  [Ref]$parsedVersion = $null
  if (![Version]::TryParse($osVersion, $parsedVersion)) {
    return $null
  }

  $version = [Version]$parsedVersion.Value
  return New-Object Version -ArgumentList $version.Major, $version.Minor
}

<#

.SYNOPSIS
Determines if CredSSP is enabled for the current server or client.

.DESCRIPTION
Check the registry value for the CredSSP enabled state.

#>

function isCredSSPEnabled() {
  Set-Variable credSSPServicePath -Option Constant -Value "WSMan:\localhost\Service\Auth\CredSSP"
  Set-Variable credSSPClientPath -Option Constant -Value "WSMan:\localhost\Client\Auth\CredSSP"

  $credSSPServerEnabled = $false;
  $credSSPClientEnabled = $false;

  $credSSPServerService = Get-Item $credSSPServicePath -ErrorAction SilentlyContinue
  if ($credSSPServerService) {
    $credSSPServerEnabled = [System.Convert]::ToBoolean($credSSPServerService.Value)
  }

  $credSSPClientService = Get-Item $credSSPClientPath -ErrorAction SilentlyContinue
  if ($credSSPClientService) {
    $credSSPClientEnabled = [System.Convert]::ToBoolean($credSSPClientService.Value)
  }

  return ($credSSPServerEnabled -or $credSSPClientEnabled)
}

<#

.SYNOPSIS
Determines if the Hyper-V role is installed for the current server or client.

.DESCRIPTION
The Hyper-V role is installed when the VMMS service is available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>

function isHyperVRoleInstalled() {
  $vmmsService = Get-Service -Name "VMMS" -ErrorAction SilentlyContinue

  return $vmmsService -and $vmmsService.Name -eq "VMMS"
}

<#

.SYNOPSIS
Determines if the Hyper-V PowerShell support module is installed for the current server or client.

.DESCRIPTION
The Hyper-V PowerShell support module is installed when the modules cmdlets are available.  This is much
faster then checking Get-WindowsFeature and works on Windows Client SKUs.

#>
function isHyperVPowerShellSupportInstalled() {
  # quicker way to find the module existence. it doesn't load the module.
  return !!(Get-Module -ListAvailable Hyper-V -ErrorAction SilentlyContinue)
}

<#

.SYNOPSIS
Determines if Windows Management Framework (WMF) 5.0, or higher, is installed for the current server or client.

.DESCRIPTION
Windows Admin Center requires WMF 5 so check the registey for WMF version on Windows versions that are less than
Windows Server 2016.

#>
function isWMF5Installed([string] $operatingSystemVersion) {
  Set-Variable Server2016 -Option Constant -Value (New-Object Version '10.0')   # And Windows 10 client SKUs
  Set-Variable Server2012 -Option Constant -Value (New-Object Version '6.2')

  $version = convertOsVersion $operatingSystemVersion
  if (-not $version) {
    # Since the OS version string is not properly formatted we cannot know the true installed state.
    return $false
  }

  if ($version -ge $Server2016) {
    # It's okay to assume that 2016 and up comes with WMF 5 or higher installed
    return $true
  }
  else {
    if ($version -ge $Server2012) {
      # Windows 2012/2012R2 are supported as long as WMF 5 or higher is installed
      $registryKey = 'HKLM:\SOFTWARE\Microsoft\PowerShell\3\PowerShellEngine'
      $registryKeyValue = Get-ItemProperty -Path $registryKey -Name PowerShellVersion -ErrorAction SilentlyContinue

      if ($registryKeyValue -and ($registryKeyValue.PowerShellVersion.Length -ne 0)) {
        $installedWmfVersion = [Version]$registryKeyValue.PowerShellVersion

        if ($installedWmfVersion -ge [Version]'5.0') {
          return $true
        }
      }
    }
  }

  return $false
}

<#

.SYNOPSIS
Determines if the current usser is a system administrator of the current server or client.

.DESCRIPTION
Determines if the current usser is a system administrator of the current server or client.

#>
function isUserAnAdministrator() {
  return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

<#

.SYNOPSIS
Get some basic information about the Failover Cluster that is running on this server.

.DESCRIPTION
Create a basic inventory of the Failover Cluster that may be running in this server.

#>
function getClusterInformation() {
  $returnValues = @{ }

  $returnValues.IsS2dEnabled = $false
  $returnValues.IsCluster = $false
  $returnValues.ClusterFqdn = $null
  $returnValues.IsBritannicaEnabled = $false

  $namespace = Get-CimInstance -Namespace root/MSCluster -ClassName __NAMESPACE -ErrorAction SilentlyContinue
  if ($namespace) {
    $cluster = Get-CimInstance -Namespace root/MSCluster -ClassName MSCluster_Cluster -ErrorAction SilentlyContinue
    if ($cluster) {
      $returnValues.IsCluster = $true
      $returnValues.ClusterFqdn = $cluster.Fqdn
      $returnValues.IsS2dEnabled = !!(Get-Member -InputObject $cluster -Name "S2DEnabled") -and ($cluster.S2DEnabled -gt 0)
      $returnValues.IsBritannicaEnabled = $null -ne (Get-CimInstance -Namespace root/sddc/management -ClassName SDDC_Cluster -ErrorAction SilentlyContinue)
    }
  }

  return $returnValues
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the passed in computer name.

#>
function getComputerFqdnAndAddress($computerName) {
  $hostEntry = [System.Net.Dns]::GetHostEntry($computerName)
  $addressList = @()
  foreach ($item in $hostEntry.AddressList) {
    $address = New-Object PSObject
    $address | Add-Member -MemberType NoteProperty -Name 'IpAddress' -Value $item.ToString()
    $address | Add-Member -MemberType NoteProperty -Name 'AddressFamily' -Value $item.AddressFamily.ToString()
    $addressList += $address
  }

  $result = New-Object PSObject
  $result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $hostEntry.HostName
  $result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $addressList
  return $result
}

<#

.SYNOPSIS
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

.DESCRIPTION
Get the Fully Qaulified Domain (DNS domain) Name (FQDN) of the current server or client.

#>
function getHostFqdnAndAddress($computerSystem) {
  $computerName = $computerSystem.DNSHostName
  if (!$computerName) {
    $computerName = $computerSystem.Name
  }

  return getComputerFqdnAndAddress $computerName
}

<#

.SYNOPSIS
Are the needed management CIM interfaces available on the current server or client.

.DESCRIPTION
Check for the presence of the required server management CIM interfaces.

#>
function getManagementToolsSupportInformation() {
  $returnValues = @{ }

  $returnValues.ManagementToolsAvailable = $false
  $returnValues.ServerManagerAvailable = $false

  $namespaces = Get-CimInstance -Namespace root/microsoft/windows -ClassName __NAMESPACE -ErrorAction SilentlyContinue

  if ($namespaces) {
    $returnValues.ManagementToolsAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ManagementTools" })
    $returnValues.ServerManagerAvailable = !!($namespaces | Where-Object { $_.Name -ieq "ServerManager" })
  }

  return $returnValues
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>
function isRemoteAppEnabled() {
  Set-Variable key -Option Constant -Value "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Terminal Server\\TSAppAllowList"

  $registryKeyValue = Get-ItemProperty -Path $key -Name fDisabledAllowList -ErrorAction SilentlyContinue

  if (-not $registryKeyValue) {
    return $false
  }
  return $registryKeyValue.fDisabledAllowList -eq 1
}

<#

.SYNOPSIS
Check the remote app enabled or not.

.DESCRIPTION
Check the remote app enabled or not.

#>

<#
c
.SYNOPSIS
Get the Win32_OperatingSystem information as well as current version information from the registry

.DESCRIPTION
Get the Win32_OperatingSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller. Included in the results are current version
information from the registry

#>
function getOperatingSystemInfo() {
  $operatingSystemInfo = Get-CimInstance Win32_OperatingSystem | Microsoft.PowerShell.Utility\Select-Object csName, Caption, OperatingSystemSKU, Version, ProductType, OSType, LastBootUpTime, SerialNumber
  $currentVersion = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" | Microsoft.PowerShell.Utility\Select-Object CurrentBuild, UBR, DisplayVersion

  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name CurrentBuild -Value $currentVersion.CurrentBuild
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name UpdateBuildRevision -Value $currentVersion.UBR
  $operatingSystemInfo | Add-Member -MemberType NoteProperty -Name DisplayVersion -Value $currentVersion.DisplayVersion

  return $operatingSystemInfo
}

<#

.SYNOPSIS
Get the Win32_ComputerSystem information

.DESCRIPTION
Get the Win32_ComputerSystem instance and filter the results to just the required properties.
This filtering will make the response payload much smaller.

#>
function getComputerSystemInfo() {
  return Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue | `
    Microsoft.PowerShell.Utility\Select-Object TotalPhysicalMemory, DomainRole, Manufacturer, Model, NumberOfLogicalProcessors, Domain, Workgroup, DNSHostName, Name, PartOfDomain, SystemFamily, SystemSKUNumber
}

<#

.SYNOPSIS
Get SMBIOS locally from the passed in machineName


.DESCRIPTION
Get SMBIOS locally from the passed in machine name

#>
function getSmbiosData($computerSystem) {
  <#
    Array of chassis types.
    The following list of ChassisTypes is copied from the latest DMTF SMBIOS specification.
    REF: https://www.dmtf.org/sites/default/files/standards/documents/DSP0134_3.1.1.pdf
  #>
  $ChassisTypes =
  @{
    1  = 'Other'
    2  = 'Unknown'
    3  = 'Desktop'
    4  = 'Low Profile Desktop'
    5  = 'Pizza Box'
    6  = 'Mini Tower'
    7  = 'Tower'
    8  = 'Portable'
    9  = 'Laptop'
    10 = 'Notebook'
    11 = 'Hand Held'
    12 = 'Docking Station'
    13 = 'All in One'
    14 = 'Sub Notebook'
    15 = 'Space-Saving'
    16 = 'Lunch Box'
    17 = 'Main System Chassis'
    18 = 'Expansion Chassis'
    19 = 'SubChassis'
    20 = 'Bus Expansion Chassis'
    21 = 'Peripheral Chassis'
    22 = 'Storage Chassis'
    23 = 'Rack Mount Chassis'
    24 = 'Sealed-Case PC'
    25 = 'Multi-system chassis'
    26 = 'Compact PCI'
    27 = 'Advanced TCA'
    28 = 'Blade'
    29 = 'Blade Enclosure'
    30 = 'Tablet'
    31 = 'Convertible'
    32 = 'Detachable'
    33 = 'IoT Gateway'
    34 = 'Embedded PC'
    35 = 'Mini PC'
    36 = 'Stick PC'
  }

  $list = New-Object System.Collections.ArrayList
  $win32_Bios = Get-CimInstance -class Win32_Bios
  $obj = New-Object -Type PSObject | Microsoft.PowerShell.Utility\Select-Object SerialNumber, Manufacturer, UUID, BaseBoardProduct, ChassisTypes, Chassis, SystemFamily, SystemSKUNumber, SMBIOSAssetTag
  $obj.SerialNumber = $win32_Bios.SerialNumber
  $obj.Manufacturer = $win32_Bios.Manufacturer
  $computerSystemProduct = Get-CimInstance Win32_ComputerSystemProduct
  if ($null -ne $computerSystemProduct) {
    $obj.UUID = $computerSystemProduct.UUID
  }
  $baseboard = Get-CimInstance Win32_BaseBoard
  if ($null -ne $baseboard) {
    $obj.BaseBoardProduct = $baseboard.Product
  }
  $systemEnclosure = Get-CimInstance Win32_SystemEnclosure
  if ($null -ne $systemEnclosure) {
    $obj.SMBIOSAssetTag = $systemEnclosure.SMBIOSAssetTag
  }
  $obj.ChassisTypes = Get-CimInstance Win32_SystemEnclosure | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty ChassisTypes
  $obj.Chassis = New-Object -TypeName 'System.Collections.ArrayList'
  $obj.ChassisTypes | ForEach-Object -Process {
    $obj.Chassis.Add($ChassisTypes[[int]$_])
  }
  $obj.SystemFamily = $computerSystem.SystemFamily
  $obj.SystemSKUNumber = $computerSystem.SystemSKUNumber
  $list.Add($obj) | Out-Null

  return $list

}

<#

.SYNOPSIS
Get the azure arc status information

.DESCRIPTION
Get the azure arc status information

#>
function getAzureArcStatus() {

  $LogName = "Microsoft-ServerManagementExperience"
  $LogSource = "SMEScript"
  $ScriptName = "Get-ServerInventory.ps1 - getAzureArcStatus()"

  Microsoft.PowerShell.Management\New-EventLog -LogName $LogName -Source $LogSource -ErrorAction SilentlyContinue

  Get-Service -Name himds -ErrorVariable Err -ErrorAction SilentlyContinue | Out-Null

  if (!!$Err) {

    $Err = "The Azure arc agent is not installed. Details: $Err"

    Microsoft.PowerShell.Management\Write-EventLog -LogName $LogName -Source $LogSource -EventId 0 -Category 0 -EntryType Information `
    -Message "[$ScriptName]: $Err"  -ErrorAction SilentlyContinue

    $status = "NotInstalled"
  }
  else {
    $status = (azcmagent show --json | ConvertFrom-Json -ErrorAction Stop).status
  }

  return $status
}

<#

.SYNOPSIS
Gets an EnforcementMode that describes the system lockdown policy on this computer.

.DESCRIPTION
By checking the system lockdown policy, we can infer if PowerShell is in ConstrainedLanguage mode as a result of an enforced WDAC policy.
Note: $ExecutionContext.SessionState.LanguageMode should not be used within a trusted (by the WDAC policy) script context for this purpose because
the language mode returned would potentially not reflect the system-wide lockdown policy/language mode outside of the execution context.

#>
function getSystemLockdownPolicy() {
  return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy().ToString()
}

<#

.SYNOPSIS
Determines if the operating system is HCI.

.DESCRIPTION
Using the operating system 'Caption' (which corresponds to the 'ProductName' registry key at HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion) to determine if a server OS is HCI.

#>
function isServerOsHCI([string] $operatingSystemCaption) {
  return $operatingSystemCaption -eq "Microsoft Azure Stack HCI"
}

###########################################################################
# main()
###########################################################################

$operatingSystem = getOperatingSystemInfo
$computerSystem = getComputerSystemInfo
$isAdministrator = isUserAnAdministrator
$fqdnAndAddress = getHostFqdnAndAddress $computerSystem
$hostname = [Environment]::MachineName
$netbios = $env:ComputerName
$managementToolsInformation = getManagementToolsSupportInformation
$isWmfInstalled = isWMF5Installed $operatingSystem.Version
$clusterInformation = getClusterInformation -ErrorAction SilentlyContinue
$isHyperVPowershellInstalled = isHyperVPowerShellSupportInstalled
$isHyperVRoleInstalled = isHyperVRoleInstalled
$isCredSSPEnabled = isCredSSPEnabled
$isRemoteAppEnabled = isRemoteAppEnabled
$smbiosData = getSmbiosData $computerSystem
$azureArcStatus = getAzureArcStatus
$systemLockdownPolicy = getSystemLockdownPolicy
$isHciServer = isServerOsHCI $operatingSystem.Caption

$result = New-Object PSObject
$result | Add-Member -MemberType NoteProperty -Name 'IsAdministrator' -Value $isAdministrator
$result | Add-Member -MemberType NoteProperty -Name 'OperatingSystem' -Value $operatingSystem
$result | Add-Member -MemberType NoteProperty -Name 'ComputerSystem' -Value $computerSystem
$result | Add-Member -MemberType NoteProperty -Name 'Fqdn' -Value $fqdnAndAddress.Fqdn
$result | Add-Member -MemberType NoteProperty -Name 'AddressList' -Value $fqdnAndAddress.AddressList
$result | Add-Member -MemberType NoteProperty -Name 'Hostname' -Value $hostname
$result | Add-Member -MemberType NoteProperty -Name 'NetBios' -Value $netbios
$result | Add-Member -MemberType NoteProperty -Name 'IsManagementToolsAvailable' -Value $managementToolsInformation.ManagementToolsAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsServerManagerAvailable' -Value $managementToolsInformation.ServerManagerAvailable
$result | Add-Member -MemberType NoteProperty -Name 'IsWmfInstalled' -Value $isWmfInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCluster' -Value $clusterInformation.IsCluster
$result | Add-Member -MemberType NoteProperty -Name 'ClusterFqdn' -Value $clusterInformation.ClusterFqdn
$result | Add-Member -MemberType NoteProperty -Name 'IsS2dEnabled' -Value $clusterInformation.IsS2dEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsBritannicaEnabled' -Value $clusterInformation.IsBritannicaEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVRoleInstalled' -Value $isHyperVRoleInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsHyperVPowershellInstalled' -Value $isHyperVPowershellInstalled
$result | Add-Member -MemberType NoteProperty -Name 'IsCredSSPEnabled' -Value $isCredSSPEnabled
$result | Add-Member -MemberType NoteProperty -Name 'IsRemoteAppEnabled' -Value $isRemoteAppEnabled
$result | Add-Member -MemberType NoteProperty -Name 'SmbiosData' -Value $smbiosData
$result | Add-Member -MemberType NoteProperty -Name 'AzureArcStatus' -Value $azureArcStatus
$result | Add-Member -MemberType NoteProperty -Name 'SystemLockdownPolicy' -Value $systemLockdownPolicy
$result | Add-Member -MemberType NoteProperty -Name 'IsHciServer' -Value $isHciServer

$result

}
## [END] Get-WACSTServerInventory ##

# SIG # Begin signature block
# MIIoKgYJKoZIhvcNAQcCoIIoGzCCKBcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCFgOy2ShekQv7H
# EJexc89vbfIV2jEcSZzpshxrpKOv96CCDXYwggX0MIID3KADAgECAhMzAAADTrU8
# esGEb+srAAAAAANOMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI5WhcNMjQwMzE0MTg0MzI5WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDdCKiNI6IBFWuvJUmf6WdOJqZmIwYs5G7AJD5UbcL6tsC+EBPDbr36pFGo1bsU
# p53nRyFYnncoMg8FK0d8jLlw0lgexDDr7gicf2zOBFWqfv/nSLwzJFNP5W03DF/1
# 1oZ12rSFqGlm+O46cRjTDFBpMRCZZGddZlRBjivby0eI1VgTD1TvAdfBYQe82fhm
# WQkYR/lWmAK+vW/1+bO7jHaxXTNCxLIBW07F8PBjUcwFxxyfbe2mHB4h1L4U0Ofa
# +HX/aREQ7SqYZz59sXM2ySOfvYyIjnqSO80NGBaz5DvzIG88J0+BNhOu2jl6Dfcq
# jYQs1H/PMSQIK6E7lXDXSpXzAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUnMc7Zn/ukKBsBiWkwdNfsN5pdwAw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMDUxNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAD21v9pHoLdBSNlFAjmk
# mx4XxOZAPsVxxXbDyQv1+kGDe9XpgBnT1lXnx7JDpFMKBwAyIwdInmvhK9pGBa31
# TyeL3p7R2s0L8SABPPRJHAEk4NHpBXxHjm4TKjezAbSqqbgsy10Y7KApy+9UrKa2
# kGmsuASsk95PVm5vem7OmTs42vm0BJUU+JPQLg8Y/sdj3TtSfLYYZAaJwTAIgi7d
# hzn5hatLo7Dhz+4T+MrFd+6LUa2U3zr97QwzDthx+RP9/RZnur4inzSQsG5DCVIM
# pA1l2NWEA3KAca0tI2l6hQNYsaKL1kefdfHCrPxEry8onJjyGGv9YKoLv6AOO7Oh
# JEmbQlz/xksYG2N/JSOJ+QqYpGTEuYFYVWain7He6jgb41JbpOGKDdE/b+V2q/gX
# UgFe2gdwTpCDsvh8SMRoq1/BNXcr7iTAU38Vgr83iVtPYmFhZOVM0ULp/kKTVoir
# IpP2KCxT4OekOctt8grYnhJ16QMjmMv5o53hjNFXOxigkQWYzUO+6w50g0FAeFa8
# 5ugCCB6lXEk21FFB1FdIHpjSQf+LP/W2OV/HfhC3uTPgKbRtXo83TZYEudooyZ/A
# Vu08sibZ3MkGOJORLERNwKm2G7oqdOv4Qj8Z0JrGgMzj46NFKAxkLSpE5oHQYP1H
# tPx1lPfD7iNSbJsP6LiUHXH1MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAANOtTx6wYRv6ysAAAAAA04wDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIJPI6J8DA/usUdiVuntDeyEr
# 33QARE83BjqxpV0nBD20MEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEAycXwy3CZ3xjloqW4SnmrtpkLRxbYjH/6X7A5/rhdwMdvXZs957XTnqoW
# htYXmC5xZk2Hb1LYPq/0kyljnNQqIHM1XFJy+MczzBIORs+8yFbJ91NVwQD0NPQu
# BKQZFdya89oITyzYxZMog8Ar/AgHE4rkgCOZg0b89ogcsrxhbIS7Rc8aaGxiuH7T
# jhrEO9n0OlL8Fhp6QTnoU0TVgtLo1VbxVjQXC24um+V8ylL8uq9SDwss72HR62wo
# 6chgkJA8xhBqtasprklOgSANTLuQnupNm38jm7PdAeAaZq7yP+mCRy+gWRzQosnv
# Iw6BuCAi94f0PINbzoqMHSDdaoGXJKGCF5QwgheQBgorBgEEAYI3AwMBMYIXgDCC
# F3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFSBgsq
# hkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCCw9vNvcjhzC/QNX0Z6JR+GHkpLwDWpUYo2CfV9iB7jXgIGZVbH85ML
# GBMyMDIzMTIwNjIzNDc1MC43NTRaMASAAgH0oIHRpIHOMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046ODYwMy0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wg
# ghHqMIIHIDCCBQigAwIBAgITMwAAAdebDR5XLoxRjgABAAAB1zANBgkqhkiG9w0B
# AQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMzA1MjUxOTEy
# MzdaFw0yNDAyMDExOTEyMzdaMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046ODYwMy0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDErGCkN2X/UvuNCcfl0yVBNo+LIIyzG7A10X5kVgGn
# p9s8mf4aZsukZu5rvLs7NqaNExcwnPuHIWdp6kswja1Yw9SxTX+E0leq+WBucIRK
# WdcMumIDBgLE0Eb/3/BY95ZtT1XsnnatBFZhr0uLkDiT9HgrRb122sm7/YkyMigF
# kT0JuoiSPXoLL7waUE9teI9QOkojqjRlcIC4YVNY+2UIBM5QorKNaOdz/so+TIF6
# mzxX5ny2U/o/iMFVTfvwm4T8g/Yqxwye+lOma9KK98v6vwe/ii72TMTVWwKXFdXO
# ysP9GiocXt38cuP9c8aE1eH3q4FdGTgKOd0rG+xhCgsRF8GqLT7k58VpQnJ8u+yj
# RW6Lomt5Rcropgf9EH8e4foDUoUyU5Q7iPgwOJxYhoKxRjGZlthDmp5ex+6U6zv9
# 5rd973668pGpCku0IB43L/BTzMcDAV4/xu6RfcVFwarN/yJq5qfZyMspH5gcaTCV
# AouXkQTc8LwtfxtgIz53qMSVR9c9gkSnxM5c1tHgiMX3D2GBnQan95ty+CdTYAAh
# jgBTcyj9P7OGEMhr3lyaZxjr3gps6Zmo47VOTI8tsSYHhHtD8BpBog39L5e4/lDJ
# g/Oq4rGsFKSxMXuIRZ1E08dmX67XM7qmvm27O804ChEmb+COR8Wb46MFEEz62ju+
# xQIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFK6nwLv9WQL3NIxEJyPuJMZ6MI2NMB8G
# A1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY3Jvc29mdCUy
# MFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBsBggrBgEFBQcBAQRgMF4w
# XAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvY2Vy
# dHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3J0MAwG
# A1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQD
# AgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBSBd3UJ+IsvdMCX+K7xqHa5UBtVC1CaXZv
# HRd+stW0lXA/dTNneCW0TFrBoJY59b9fnbTouPReaku2l3X5bmhsao6DCRVuqcmh
# VPAZySXGeoVfj52cLGiyZLEw6TQzu6D++vjNOGmSibO0KE9Gdv8hQERx5RG0KgrT
# mk8ckeC1VUqueUQHKVCESqTDUDD8dXTLWCmm6HqmQX6/+gKDSXggwpc75hi2AbKS
# o4tulMwTfXJdGdwrsiHjkz8nzIW/Z3PnMgGFU76KuzYFV0XyH9DTS/DPO86RLtQj
# A5ZlVGymTPfTnw7kxoiLJN/yluMHIkHSzpaJvCiqX+Dn1QGREEnNIZeRekvLourq
# PREIOTm1bJRJ065c9YX7bJ0naPixzm5y8Y2B+YIIEAi4jUraOh3oE7a4JvIW3Eg3
# oNqP7qhpd7xMLxq2WnM+U9bqWTeT4VCopAhXu2uGQexdLq7bWdcYwyEFDhS4Z9N0
# uw3h6bjB7S4MX96pfYSEV0MKFGOKbmfCUS7WemkuFqZy0oNHPPx+cfdNYeSF6bhO
# PHdsro1EVd3zWIkdD1G5kEDPnEQtFartM8H+bv5zUhAUJs8qLzuFAdBZQLueD9XZ
# eynjQKwEeAz63xATICh8tOUM2zMgSEhVL8Hm45SB6foes4BTC0Y8SZWov3Iahtvw
# yHFbUqs1YjCCB3EwggVZoAMCAQICEzMAAAAVxedrngKbSZkAAAAAABUwDQYJKoZI
# hvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# MjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAy
# MDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg4r25Phdg
# M/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjoYH1qUoNEt6aORmsHFPPF
# dvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTamDu6
# GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v3byNpOORj7I5LFGc6XBp
# Dco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL64NF50Zu
# yjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4dPf0gz3N9QZpGdc3E
# XzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0
# lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8QmguEOqEUUbi0b1q
# GFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ
# +QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435UsSFF5PA
# PBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIurQIDAQABo4IB3TCCAdkw
# EgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTEmr6CkTxG
# NSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMFwGA1UdIARV
# MFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5Lmh0bTATBgNVHSUEDDAK
# BggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMC
# AYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvX
# zpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYI
# KwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDANBgkqhkiG
# 9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv6lwUtj5OR2R4sQaTlz0x
# M7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th542DYunKmC
# VgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1bSNU5HhTdSRXud2f8449
# xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVCs/wM
# nosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU6ZGyqVvfSaN0DLzskYDS
# PeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDFNLB62FD+CljdQDzHVG2d
# Y3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltEAY5aGZFrDZ+kKNxn
# GSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdUCbFpAUR+fKFhbHP+Crvs
# QWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmNcP7ntdAoGokL
# jzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTmdHRbatGePu1+oDEzfbzL
# 6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6FwZvKhggNN
# MIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEn
# MCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjg2MDMtMDVFMC1EOTQ3MSUwIwYDVQQD
# ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQAx
# W9uizG3hEY89uL2uu+X+mG/rdaCBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
# IFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6Rr7BTAiGA8yMDIzMTIwNjEzNTIw
# NVoYDzIwMjMxMjA3MTM1MjA1WjB0MDoGCisGAQQBhFkKBAExLDAqMAoCBQDpGvsF
# AgEAMAcCAQACAhmuMAcCAQACAhQUMAoCBQDpHEyFAgEAMDYGCisGAQQBhFkKBAIx
# KDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZI
# hvcNAQELBQADggEBABBj7ThDNMdxaHkX0TozBJ4noxWOIa+kVWLc+iIn4xY4LqQO
# r0pgMDMDPdEBem0JzJVucdRW53HlEXDlKBkpCisKh3Ol+tg0jxUTneEjF2CWKAZm
# LLDSrxAODA2DwKn3yZOg0B2tVaffGWyMsTU/P3FVZhBZHzQokfx2mGxCZlYJT+pd
# 0zNuNmfc/oCYo+2HWaPtSAv1JUDRN2i/zcFapyXSQWgYuRPcXdYEvFcSAPzHOYFB
# 6scdTa+xaXMKaJVee4PwVXoL98UezQK2uU8hGQd81mZ0QfaIcFL9KY/ZbwOnkfVh
# yc9VpC1qijmBWSo1cLKPGpr1Qx8lRVxzimUkNHExggQNMIIECQIBATCBkzB8MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
# b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAdebDR5XLoxRjgABAAAB1zAN
# BglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8G
# CSqGSIb3DQEJBDEiBCCGm6fy+ISdISGYGMs4h8ybc/8ynKW3mhBLSpsj2N+rUzCB
# +gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIJzePl5LXn1PiqNjx8YN7TN1ZI0d
# 1ZX/2zRdnI97rJo7MIGYMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
# MTACEzMAAAHXmw0eVy6MUY4AAQAAAdcwIgQgV17fu48zPARGXcyXuc6lkXPnvIcx
# xRY1Q90gFJm4BNEwDQYJKoZIhvcNAQELBQAEggIAh4KyFd68x2uFoSkPG87C7d9S
# WBboIGKo2gdMonkxabAV2WMeLniymteReStsgNusxw4pR64m82Qx0gkwxArG/F8t
# VxqW1A8u2xFQnCFXRtwwE0YMi4bmzUuKHk12i80Fmru5W3Q3BFmaWZtjkT4gpw3w
# s8IRIKyVUwpguCDchO6ik1HUZiP0spMMJIfoCHlTNJc5WBYtl4qe5XVrd5nowzWY
# ES0b3CJsutLvoN16Y1mH5XhwjZ91Hai7jKn9mYxnpR7O8Z/IHi7CGasf6aO4BpOF
# D3ocpCKMz7aZNPP9URkySk1gJwQg4xgIUF92iOh7MiMaN+hzcluHdfO77TMWmmJ6
# 3Ek/pGRsCu0/AclOUdKMQ3lil/Pdcwl0J6fBIpcpRb+gN6Dy5YSu3GVvfK55mqGl
# rrmOzkSuY/J7WHEDGzjIjo8wraeWtLPgVq4EDDAk5i8BHQ4mKHtKxA2ysHalsFV9
# lk0OPzlLB9oevOZDY5TFzQ2xx8RDO1h8T7RIUoy6t7WC6Is1kRMayhlYPmLq+uax
# cqx+L7YLGjsdcOUb6V0+uQQ2rM28typK7OakKdgvRLNUzzP9lA9Ra1l370pxJ94z
# EX2XTVzZQSKg4rz4hbmkSKuuV8YzIZ27xaRLOqA1MWy58aiUgA5kLMCV7HAERt8k
# AXWIHnJX150hrvbeypI=
# SIG # End signature block
