<#
.EXTERNALHELP CollectOverMetrics-help.xml
#>

# Copyright (c) Microsoft Corporation. All rights reserved.

param (
    [Parameter(ParameterSetName="Default",Mandatory=$true)] [object] $DatabaseAvailabilityGroup,
    [Parameter(ParameterSetName="Default")] [DateTime] $EndTime,
    [Parameter(ParameterSetName="Default")] [DateTime] $StartTime,
    [Parameter(ParameterSetName="Default")] [Switch] $GenerateHtmlReport = $false,
    [Parameter(ParameterSetName="Default")] [Switch] $RawOutput = $false,
    [Parameter(ParameterSetName="Default")] [Switch] $IncludeExtendedEvents = $false,
    [Parameter(ParameterSetName="Summarise",Mandatory=$true)] [String[]] $SummariseCsvFiles,
    [object[]] $Database,
	[switch] $ShowHtmlReport,
    [string] $ReportPath = (Get-Location).Path,
    [Switch] $MergeCsvFiles,
    [ScriptBlock] $ReportFilter = { return $true }
    )
    
Import-LocalizedData -BindingVariable CollectOverMetrics_LocalizedStrings -FileName CollectOverMetrics.strings.psd1
Set-StrictMode -Version 2.0

###
### The ActiveManager raises ops channel events for each of the actions that it takes,
### and for each "top level" event, it also raises events for the specific stages that
### the operation is passing through.
###
### For each action (top-level or subaction), there is a specific event for the start
### of the action and two events that could be raised at the end (success or failure).
### The final event also includes the time that was taken by that action.
###
### This table lays out the details for each of the kinds of action that the ActiveManager
### reports: the init, success, and fail event IDs; whether it's a top-level or sub-
### action; and the output field to write the duration into.
###
### We then pivot this definition around to make a table indexed by the event ID, which
### we can then use to lookup the incoming events from the machines.  Each event ID then
### gets a record of the action (move, acll, etc), the phase (init, fail, success),
### whether it was a top-level action, and what field to write its duration to.
###
$ActionEventMarkers = @{
    Mount =    @{ Init = 300; Failed = 301; Success = 302; IsTopLevelAction = $true; OutputField = "DurationOutage"; }
    Dismount = @{ Init = 303; Failed = 304; Success = 305; IsTopLevelAction = $true; OutputField = "DurationOutage"; }
    Move =     @{ Init = 306; Failed = 307; Success = 308; IsTopLevelAction = $true; OutputField = "DurationOutage"; }
    Remount =  @{ Init = 309; Failed = 310; Success = 311; IsTopLevelAction = $true; OutputField = "DurationOutage"; }
    Acll =     @{ Init = 312; Failed = 313; Success = 314; IsTopLevelAction = $false;  OutputField = "DurationAcll"; }
    Bcs =      @{ Init = 334; Success = 335; IsTopLevelAction = $false;  OutputField = "DurationBcs"; }
    DirectMount =   @{ Init = 315; Failed = 316; Success = 318; IsTopLevelAction = $false; OutputField = "DurationMount"; }
    StoreDismount = @{ Init = 319; Failed = 320; Success = 321; IsTopLevelAction = $false; OutputField = "DurationDismount"; }
    }
    
$ActionEventLookupTable = @{}

foreach ($Action in $ActionEventMarkers.Keys) {

    $ActionData = $ActionEventMarkers[$Action]
    $IsTopLevelAction = $ActionData["IsTopLevelAction"]
    $OutputField = $ActionData["OutputField"]

    foreach ($Phase in @("Init","Failed","Success", "Perf")) {
        
        if ($ActionData.ContainsKey($Phase)) {    
            $ActionEventLookupTable[ $ActionData[$Phase] ] = @{
                Action = $Action;
                Phase = $Phase;
                IsTopLevelAction = $IsTopLevelAction;
                OutputField = $OutputField;
            }
        }
    }
}

### (For building the XPath queries) Find contiguous series of event Ids
### in the event ID lists.  Given a list of event Ids
###
###     100,101,102,105,110,111,112
###
### this method returns the list of query elements:
###
###     (EventId >=100 and EventId <=102)
###     EventId=105
###     (EventId >=110 and EventId <=112)
###
### which we then join together to form the query
###
function FindRanges (
    [Parameter(Mandatory=$true)]
    [int[]] $EventIds
    ) {
    
   $start = $null
   $last = $null
    foreach ($next in ($EventIds | Sort -Descending:$false)) {
        if ($start -eq $null) {
            $start = $next
            $last = $next
        } else {
            if ( ($next - $last) -eq 1 ) {
                $last = $next
            } else {
                if ($last -eq $start) {
                    Write-Output "EventID=$start"
                } else {
                    write-output "(EventID >= $start and EventID <= $last)"
                }
                $start = $next
                $last = $next
            }
        }
    }
    
    if ($last -eq $start) {
        Write-Output "EventID=$start"
    } else {
        write-output "(EventID >= $start and EventID <= $last)"
    }
}

### Build a simple XPath query that will filter for some set of
### event IDs within a certain time frame
function BuildXPathQueryString (
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime,
    [String] $Provider,
    [int[]] $EventIds
    ) {
    
    $strStartTime = $StartTime.ToUniversalTime().ToString("o")
    $strEndTime   = $EndTime.ToUniversalTime().ToString("o")
    
    $QueryString = "*[ System[ "
    $QueryString += "TimeCreated[@SystemTime > '$strStartTime' and @SystemTime < '$strEndTime']"
    
    if ($EventIds) {
        $EventIdQuery = [String]::Join(" or ", (FindRanges -EventIds $EventIds))
        $QueryString += " and ($EventIdQuery)"
    }
    
    if ($Provider) {
        $QueryString += " and Provider[@Name = '$Provider']"
    }
    
    $QueryString += " ] ]"
    
    return $QueryString
}

### These objects collect the information for a single ActiveManager
### operation.  Each of these objects will be a single row in the CSV
function CreateEmptyActionEntry
{
    $entry = New-Object object
    
    Add-Member -InputObject $entry -MemberType NoteProperty -Name DatabaseName        -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name TimeRecoveryStarted -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name TimeRecoveryEnded   -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name ActionInitiator     -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name ActionCategory      -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name ActionReason        -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name Result              -Value "Unknown"
    Add-Member -InputObject $entry -MemberType NoteProperty -Name DurationOutage      -Value 0.0
    Add-Member -InputObject $entry -MemberType NoteProperty -Name DurationDismount    -Value 0.0
    Add-Member -InputObject $entry -MemberType NoteProperty -Name DurationBcs         -Value 0.0
    Add-Member -InputObject $entry -MemberType NoteProperty -Name DurationAcll        -Value 0.0
    Add-Member -InputObject $entry -MemberType NoteProperty -Name DurationMount       -Value 0.0
    Add-Member -InputObject $entry -MemberType NoteProperty -Name DurationOther       -Value 0.0
    Add-Member -InputObject $entry -MemberType NoteProperty -Name ActiveOnStart       -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name ActiveOnFinish      -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name PAMServer           -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name NumberOfAttempts    -Value ([int] 0)
    Add-Member -InputObject $entry -MemberType NoteProperty -Name ReplayedLogs        -Value ([int] 0)
    Add-Member -InputObject $entry -MemberType NoteProperty -Name LostLogs            -Value ([int] 0)
    Add-Member -InputObject $entry -MemberType NoteProperty -Name LostBytes           -Value ([int] 0)
    Add-Member -InputObject $entry -MemberType NoteProperty -Name DatabaseGuid        -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name UniqueOperationId   -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name ErrorMessage        -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name StoreMountLids      -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name StoreMountProgress  -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name EseMountTiming      -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name EseDismountTiming   -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name BcsTimingDetails    -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name AcllTimingDetails   -Value $null
    Add-Member -InputObject $entry -MemberType NoteProperty -Name AcllCopiedLogs      -Value ([int] 0)
    Add-Member -InputObject $entry -MemberType NoteProperty -Name AcllFinalReplayQueue -Value ([int] 0)
    Add-Member -InputObject $entry -MemberType NoteProperty -Name MountRpcTimingDetails -Value $null

    return $entry
}

### 
### The flow for getting the data from ActiveManager's database operation events is:
### 
### 1 For each server in the DAG, read all the events with IDs defined in $ActionEventLookupTable
### 2 Group them by the unique operation IDs - each group now represents an operation
###
### These first two steps are in GetPamOperationalChannelStatsForServer
###
### 3 For each group, create an object that will hold the stats for it.
###   Each object will become a row in the table.
### 4 From each event in the group, populate the relevant fields in the object
### 
### Processing a group into a single object is in GenerateStatisticsFromGroup.
### Then we can continue by looking for the ACLL loss events that match each
### operation, matching by names & times in RTM, by operation ID in SP1. This is
### function MergeAcllLossReportsFromServer
###
### Finally, we merge in the store mount LID events (MergeStoreMountTimingEventsFromServers).
### Matching these into operations is trickier because the events are from the Application
### log which has only second-level resolution on its timestamps, rather than the tick-level
### resolution that the other channel's events have.  We round the times down to the nearest
### second for the comparisons.  This could mean that two events for the same database on the
### same server will each hit the same second interval, and we will just append these matches
### together.
###
function GenerateStatisticsFromGroup (
    [Parameter(Mandatory=$true, ValueFromPipeline=$true)] [Microsoft.PowerShell.Commands.GroupInfo] $EventGroup
    )  {
    
    Begin {
        $ActionMap = @{
            AdminMove         = @{ ActionCategory = "Move";  ActionInitiator = "Admin"};
            AdminMount        = @{ ActionCategory = "Mount"; ActionInitiator = "Admin"};
            AdminDismount     = @{ ActionCategory = "Dismount"; ActionInitiator = "Admin"};
            StartupAutoMount  = @{ ActionCategory = "Mount"; ActionInitiator = "Automatic"};
            StoreRestartMount = @{ ActionCategory = "Mount"; ActionInitiator = "Automatic"};
            AutomaticFailover = @{ ActionCategory = "Move";  ActionInitiator = "Automatic"};
            SwitchOver        = @{ ActionCategory = "Move";  ActionInitiator = "Automatic"};
        }
        
        $BcsTimingDetailFieldsMap = @{
            HasDatabaseBeenMountedElapsedTime = "HDBMET";
            GetDatabaseCopiesElapsedTime = "GDCET";
            DetermineServersToContactElapsedTime = "DSTCET";
            GetCopyStatusRpcElapsedTime = "GCSRET"
        }
    }
    
    Process {
    
        $OutputObject = CreateEmptyActionEntry
        $OutputObject.UniqueOperationId = $EventGroup.Name
    
        $EventGroup.Group |
            Sort -property TimeCreated,RecordId |
            where { $ActionEventLookupTable.ContainsKey( $_.Id ) } |
            Foreach {
            
                $event = $_
                $eventXML = ([xml] $event.ToXml()).Event.UserData.EventXML
                
                $ActionData = $ActionEventLookupTable[$event.Id]
                # Read the records in the action's data into some local variables.
                ($Action, $Phase, $IsTopLevelAction, $OutputField) = $ActionData["Action", "Phase", "IsTopLevelAction", "OutputField"]
            
                switch ($Phase) {
                    Init {
                        
                        if ($IsTopLevelAction) {
                            $OutputObject.DatabaseName    = $eventXML.DatabaseName
                            $OutputObject.DatabaseGuid    = $eventXML.DatabaseGuid
                            $OutputObject.TimeRecoveryStarted = $event.TimeCreated
                            $OutputObject.ActionInitiator = $eventXML.ActionInitiator
                            $OutputObject.ActionCategory  = $eventXML.ActionCategory
                            $OutputObject.ActionReason    = $eventXML.ActionReason
                            $OutputObject.ActiveOnStart   = $eventXML.ActiveServer
                            $OutputObject.PAMServer       = $event.MachineName
                            
                            # R3 servers don't output the same elements in these fields, so we have to do some re-mapping
                            if ($OutputObject.ActionCategory -and $ActionMap.ContainsKey($OutputObject.ActionCategory)) {
                                $category = $OutputObject.ActionCategory
                                $OutputObject.ActionCategory  = $ActionMap[$category]["ActionCategory"]
                                $OutputObject.ActionInitiator = $ActionMap[$category]["ActionInitiator"]
                            }
							
							if ( ( $eventXML.GetElementsByTagName("MoveComment") | Measure-Object ).Count -gt 0 )
							{
								if ( $eventXML.MoveComment.Trim() -match "Moved as part of database redistribution \(RedistributeActiveDatabases\.ps1\)\." )
								{
									# The DB was moved by the rebalancer script (RedistributeActiveDatabases.ps1).
									# NOTE: The previous value was "Cmdlet".
									$OutputObject.ActionReason = "Rebalance"
								}
							}
                        }
                        
                    }
                    
                    default {
                    
                        $OutputObject.$OutputField += ([TimeSpan] $eventXML.ElapsedTime).TotalSeconds
                    
                        if ($IsTopLevelAction) {
                            $OutputObject.TimeRecoveryEnded = $event.TimeCreated
                            $OutputObject.ActiveOnFinish = $eventXML.ActiveServer
                            $OutputObject.Result = $Phase
                            
                            # Remove any newlines in the error message, they screw up reading the CSV
                            if ((Get-Member -InputObject $eventXML -Name ErrorMessage) -and $eventXml.ErrorMessage) {
                                $OutputObject.ErrorMessage = $eventXML.ErrorMessage.Replace("`n", " \n ")
                            }
                        }
                        else {
                            $OutputObject.NumberOfAttempts = $eventXML.AttemptNumber
                        }
                        
                        # Add the extra timing data from ACLL and BCS.1
                        switch ($Phase) {
                            Acll {
                            }
                            Bcs {
                                $RemappedTimingFieldStrings = 
                                    foreach ($key in $BcsTimingDetailFieldsMap.Keys) {
                                        "$($BcsTimingDetailFieldsMap[$key])=$($eventXML.$key)"
                                    } 
                                
                                $OutputObject.BcsTimingDetails =
                                    [String]::Join("; ", $RemappedTimingFieldStrings)
                            }
                        }
                    }
                }
            }
        
        # Update the DurationOther field with the difference between the total time and
        # the sum of the sub-action's durations.
        $OutputObject.DurationOther =
            $OutputObject.DurationOutage -
            ($OutputObject.DurationDismount + $OutputObject.DurationAcll + $OutputObject.DurationMount)
        
        return $OutputObject
    }
}


function GetPamOperationalChannelStatsForServer(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [String] $Server, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {
    
    Process {
        write-host ($CollectOverMetrics_LocalizedStrings.res_0000 -f $Server)
    
        $XPathFilter = BuildXPathQueryString -StartTime $StartTime -EndTime $EndTime -EventIds ($ActionEventLookupTable.Keys)

        $wEvents = Get-WinEvent -ComputerName $Server -LogName "Microsoft-Exchange-HighAvailability/Operational" -FilterXPath $XPathFilter -ErrorAction SilentlyContinue
    
        return $wEvents |
            Group -Property {$_.Properties[0].Value} |
            GenerateStatisticsFromGroup
    }
}


function GetPamOperationalChannelStatsForServers(
    [Parameter(Mandatory=$true)] [String[]] $Servers, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {

    $PamStats = @()

    foreach ($Server in $Servers) {
        $PamStats += @(GetPamOperationalChannelStatsForServer -Server $Server -StartTime $StartTime -EndTime $EndTime)
    }
    
    write-host ($CollectOverMetrics_LocalizedStrings.res_0001 -f $PamStats.Count)

    return $PamStats
}


function MergeAcllPerfEventsFromServer(
    [Parameter(Mandatory=$true)] [Object[]] $AmOperationsData,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [String] $Server, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {
    
    Begin {
    
        $AcllTimingDetailsFieldsMap = @{
            AcllQueuedOpStartElapsedTime="AQOpStET";
            AcquireSuspendLockElapsedTime="ASusLkET";
            CopyLogsOverallElapsedTime="CLOET";
            CopyStatus="CS";
            IsNewCopierInspectorCreated="INCIC";
            IsAcllFoundDeadConnection="IAFDC";
            IsAcllCouldNotControlLogCopier="IACNCLC";
        }
        
    }
    
    Process {

        $XPathFilter = BuildXPathQueryString -StartTime $StartTime -EndTime $EndTime -EventIds 336
        @(Get-WinEvent -ComputerName $Server -LogName "Microsoft-Exchange-HighAvailability/Operational" -FilterXPath $XPathFilter -ErrorAction SilentlyContinue) |
            foreach {
                $AcllPerfEvent = $_
                $eventXml = ([xml] $AcllPerfEvent.ToXml()).Event.UserData.EventXML
                
                :FIND_MATCHING_OP foreach ($op in $AmOperationsData) { 
                
                    if ($op.UniqueOperationId -eq $eventXML.UniqueId) {
                    
                        $op.AcllCopiedLogs = $eventXML.NumberOfLogsCopied
                        $op.AcllFinalReplayQueue = $eventXML.ReplayQueueLengthAcllEnd
                        
                        $RemappedAcllFields =
                            foreach ($key in $AcllTimingDetailsFieldsMap.Keys) {
                                "$($AcllTimingDetailsFieldsMap[$key])=$($eventXML.$key)"
                            }
                            
                        if ($op.AcllTimingDetails) {
                            $op.AcllTimingDetails += ";; "
                        }
                        $op.AcllTimingDetails +=
                            [String]::Join("; ", $RemappedAcllFields)
                    
                        break FIND_MATCHING_OP
                        
                    }
                }
            }
    }
}


function MergeMountRpcPerfEventsFromServer(
    [Parameter(Mandatory=$true)] [Object[]] $AmOperationsData,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [String] $Server, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {
    
    Begin {
        $MountRpcTimingDetailsFieldsMap = @{
            PreMountQueuedOpStartElapsedTime = "PMQOStET";
            PreMountQueuedOpExecutionElapsedTime = "PMQOExET";
            StoreMountElapsedTime = "SMET"
        }   
    }
    
    Process {

        $XPathFilter = BuildXPathQueryString -StartTime $StartTime -EndTime $EndTime -EventIds 337
        @(Get-WinEvent -ComputerName $Server -LogName "Microsoft-Exchange-HighAvailability/Operational" -FilterXPath $XPathFilter -ErrorAction SilentlyContinue) |
            foreach {
                $RpcPerfEvent = $_
                $eventXml = ([xml] $RpcPerfEvent.ToXml()).Event.UserData.EventXML
                
                :FIND_MATCHING_OP foreach ($op in $AmOperationsData) { 
                
                    if (($op.DatabaseGuid -eq $eventXml.DatabaseGuid) -and 
                        ($op.ActiveOnFinish -eq $RpcPerfEvent.MachineName) -and
                        ($op.TimeRecoveryStarted -le $RpcPerfEvent.TimeCreated) -and
                        ($op.TimeRecoveryEnded -ge $RpcPerfEvent.TimeCreated)) {
                        
                        $RemappedMountRpcFields =
                            foreach ($key in $MountRpcTimingDetailsFieldsMap.Keys) {
                                "$($MountRpcTimingDetailsFieldsMap[$key])=$($eventXML.$key)"
                            }
                            
                        if ($op.MountRpcTimingDetails) {
                            $op.MountRpcTimingDetails += ";; "
                        }
                        $op.MountRpcTimingDetails +=
                            [String]::Join("; ", $RemappedMountRpcFields)
                                                
                        break FIND_MATCHING_OP
                        
                    }
                }
            }
    }
}



function ParseLossReportEvent(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [System.Diagnostics.Eventing.Reader.EventLogRecord] $LossReport
    ) {
    Begin {
        $AcllLossId = 324
        $GranularLossId = 333
    }
    
    Process {
        if ($LossReport) {
            $EventXML = ([xml] $LossReport.ToXml()).Event.UserData.EventXML
    
            $OutputObject = New-Object -TypeName "Object"
            Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name SourceMachine -Value $LossReport.MachineName
            Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name Time -Value $LossReport.TimeCreated
            Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name DatabaseGuid -Value $EventXML.DatabaseGuid
            Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name LossType -Value ""
            
            switch ($LossReport.Id) {
                
                $AcllLossId {
                    $OutputObject.LossType = "ACLL"
                    Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name NumberOfLostLogs -Value $EventXML.NumberOfLostLogs
                }
                
                $GranularLossId {
                    $OutputObject.LossType = "Granular"
                    Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name LossInBytes -Value $EventXML.LossInBytes                
                }
            }
            
            return $OutputObject
        }
    }
}

function MergeLossReportsFromServer(
    [Parameter(Mandatory=$true)] [Object[]] $AmOperationsData,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [String] $Server, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {
    
    Process {

        write-host ($CollectOverMetrics_LocalizedStrings.res_0002 -f $Server)
        $XPathFilter = BuildXPathQueryString -StartTime $StartTime -EndTime $EndTime -EventIds 324,333
        @(Get-WinEvent -ComputerName $Server -LogName "Microsoft-Exchange-HighAvailability/Operational" -FilterXPath $XPathFilter -ErrorAction SilentlyContinue) |
            ParseLossReportEvent |
            foreach {
                $LossEvent = $_
                
                :FIND_MATCHING_OP foreach ($op in $AmOperationsData) { 
                
                    if (($op.DatabaseGuid -eq $LossEvent.DatabaseGuid) -and 
                        ($op.ActiveOnFinish -eq $LossEvent.SourceMachine) -and
                        ($op.TimeRecoveryStarted -le $LossEvent.Time) -and
                        ($op.TimeRecoveryEnded -ge $LossEvent.Time)) {
                        
                        if ($LossEvent.LossType -eq "ACLL") {
                            $op.LostLogs = $LossEvent.NumberOfLostLogs
                        } else {
                            $op.LostBytes = $LossEvent.LossInBytes
                        }
                        break FIND_MATCHING_OP
                        
                    }
                }
            }

    }
}

function MergeLossReportsFromServers(
    [Parameter(Mandatory=$true)] [Object[]] $AmOperationsData,
    [Parameter(Mandatory=$true)] [String[]] $Servers, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {
    
    foreach ($Server in $Servers) { 
        MergeLossReportsFromServer -Server $Server -AmOperationsData $AmOperationsData -StartTime $StartTime -EndTime $EndTime
        MergeAcllPerfEventsFromServer -Server $Server -AmOperationsData $AmOperationsData -StartTime $StartTime -EndTime $EndTime
        MergeMountRpcPerfEventsFromServer -Server $Server -AmOperationsData $AmOperationsData -StartTime $StartTime -EndTime $EndTime
    }
}

    
### Parse the array of bytes into a series of LID/Tick pairs.  Each item in the pair is a 32-bit integer,
### and we have to be careful with how we treat the order of the bytes when forming the values.
###
$LidDescriptionMap = @{
    38307 = "Start mount";
    36259 = "Post EcInitJet";
    44451 = "Post EcAttachDBs";
    52643 = "Post MDB_BASE::EcConfig";
    46499 = "Before creating system mailbox/folders";
 
    50595 = "Delete Unicode indices";
    34211 = "Detach/reattach DB while fixing corrupt indices";
 
    58787 = "Post Upgraders";
    42403 = "Detach/reattach during upgrade process";
 
    62883 = "Complete"
}

function ParseStoreLidByteArray (
    [Parameter(Mandatory=$true)] [System.Byte[]] $bytes
    ) {
    $LidTickStrings = @()
    $ProgressString = ""
    $PreviousTick = $null
    
    if ($bytes.Count -eq 0) { return "" }
    if (($bytes.Count % 8) -ne 0) { throw ($CollectOverMetrics_LocalizedStrings.res_0003 -f $bytes.count) }

    $LidTickCount = [int]($bytes.COunt / 8)
   
    foreach ($currentLidTick in @(0 .. $LidTickCount)) {
                    
        $Offset = $currentLidTick * 8
        $LidEnd = $Offset
        $LidStart = $Offset + 3
        $TickEnd = $Offset + 4
        $TickStart = $Offset + 7

        # Reverse the bytes in each word (stupid endian-ness) and then reconstruct the word out of those bytes.
        $Lid  = $bytes[$LidStart .. $LidEnd]   | foreach -Begin { $Tot = 0 } -Process { $Tot *= 256; $Tot += $_ } -End { return $Tot } 
        $Tick = $bytes[$TickStart .. $TickEnd] | foreach -Begin { $Tot = 0 } -Process { $Tot *= 256; $Tot += $_ } -End { return $Tot }                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            

        $LidName = if ($LidDescriptionMap.ContainsKey($Lid)) { $LidDescriptionMap[$Lid] } else { "($Lid)" }

        $LidTickStrings += "$Lid::$Tick"

        if (!$PreviousTick) {
        
            $PreviousTick = $Tick
            $ProgressString = "$LidName"
            
        } elseif ($Lid -ne 0) {
            
            $ProgressString += "--newline--$( ($Tick - $PreviousTick)/1000 )s--newline--$LidName"
            $PreviousTick = $Tick
        }

    }
    
    return ([String]::Join("; ", $LidTickStrings), $ProgressString)
    
}

function ParseStoreMountEvent (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [System.Diagnostics.Eventing.Reader.EventLogRecord] $MountEvent
    ) {
    
    Process {
        if ($MountEvent) {
           $data = new-object -TypeName Object
           Add-Member -InputObject $data -MemberType NoteProperty -Name "DatabaseName" -Value $MountEvent.Properties[0].Value
           Add-Member -InputObject $data -MemberType NoteProperty -Name "SourceMachine" -Value $MountEvent.MachineName
           Add-Member -InputObject $data -MemberType NoteProperty -Name "TimeCreated" -Value $MountEvent.TimeCreated
           Add-Member -InputObject $data -MemberType NoteProperty -Name "LidProgression" -Value ""
           Add-Member -InputObject $data -MemberType NoteProperty -Name "LidTickPairs" -Value ""
           if ( $MountEvent.Properties.Count -gt 6 ) {
               # The event has got the Lid/Tick data, parse it
               ($data.LidTickPairs, $data.LidProgression) = ParseStoreLidByteArray $MountEvent.Properties[6].Value
           }
           return $data
        }
    }
}

function RoundTimeDownToSecond (
    [Parameter(Mandatory=$true)] [DateTime] $Time
    ) {
    
    $ExtraTicks = $Time.Ticks % 10000000
    
    return $Time.AddTicks( -1 * $ExtraTicks )
}

function FindOperationMatchingStoreEvent (
    [Parameter(Mandatory=$true)] [HashTable] $AmOperationsLookupTable,
    [Parameter(ValueFromPipeline=$true, Mandatory=$true)] [object] $StoreEvent
    ) {
    
    if (! $AmOperationsLookupTable.ContainsKey($StoreEvent.DatabaseName) ) {
        return $null
    }
    
    $OperationsForDatabase = $AmOperationsLookupTable[$StoreEvent.DatabaseName]
    
    if (! $OperationsForDatabase["FINISH"].ContainsKey($StoreEvent.SourceMachine) ) {
        return $null
    }
    
    $OperationsForCopy = $OperationsForDatabase["FINISH"][$StoreEvent.SourceMachine]

    return $OperationsForCopy | 
        where { ((RoundTimeDownToSecond $_.TimeRecoveryStarted) -le $StoreEvent.TimeCreated) -and
                ((RoundTimeDownToSecond $_.TimeRecoveryEnded) -ge $StoreEvent.TimeCreated) } |
        select -first 1

}    
    

function MergeStoreMountTimingEventsFromServer(
    [Parameter(Mandatory=$true)] [HashTable] $AmOperationsLookupTable,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [String] $Server, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {
        
    Process {
        Write-Host ($CollectOverMetrics_LocalizedStrings.res_0004 -f $Server)
        
        $XPathFilter = BuildXPathQueryString -StartTime $StartTime -EndTime $EndTime -Provider "MsExchangeIS Mailbox Store" -EventIds 9523
        
        Get-WinEvent -ComputerName $Server -LogName "Application" -FilterXPath $XPathFilter -ErrorAction SilentlyContinue |
            ParseStoreMountEvent |
            foreach {
                $StoreMountEvent = $_
                $op = FindOperationMatchingStoreEvent -StoreEvent $StoreMountEvent -AmOperationsLookupTable $AmOperationsLookupTable
                
                if ($op) {
                    $op.StoreMountLids += $StoreMountEvent.LidTickPairs + ";; "
                    $op.StoreMountProgress += $StoreMountEvent.LidProgression + ";; "
                }
            }
    }
} # MergeStoreMountTimingEventsFromServer

function MergeStoreMountTimingEventsFromServers(
    [Parameter(Mandatory=$true)] [HashTable] $AmOperationsLookupTable,
    [Parameter(Mandatory=$true)] [String[]] $Servers, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {
    
    $Servers | MergeStoreMountTimingEventsFromServer -AmOperationsLookupTable $AmOperationsLookupTable -StartTime $StartTime -EndTime $EndTime

}


function ParseEseEvent (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [System.Diagnostics.Eventing.Reader.EventLogRecord] $EseEvent
    ) {
    
    Process {
        if ($EseEvent) {
            
            $data = new-object -TypeName Object
           
            # Strangely, the database name property has an extra ':' at the end
            Add-Member -InputObject $data -MemberType NoteProperty -Name "DatabaseName" -Value $EseEvent.Properties[2].Value.Trim(": ")
            Add-Member -InputObject $data -MemberType NoteProperty -Name "SourceMachine" -Value $EseEvent.MachineName
            Add-Member -InputObject $data -MemberType NoteProperty -Name "TimeCreated" -Value $EseEvent.TimeCreated
            Add-Member -InputObject $data -MemberType NoteProperty -Name "Type" -Value ""
           
            switch ($EseEvent.Id) {
                (103) {
                    $data.Type = "Dismount"
                    Add-Member -InputObject $data -MemberType NoteProperty -Name "EseTiming" -Value $EseEvent.Properties[4].Value
                }
                (105) {
                    $data.Type = "Mount"
                    Add-Member -InputObject $data -MemberType NoteProperty -Name "EseTiming" -Value $EseEvent.Properties[5].Value
                }
                (301) {
                    $data.Type = "ReplayedLog"
                }
            }

            return $data
        }
    }
    
}

function FindOperationMatchingEseEvent (
    [Parameter(Mandatory=$true)] [Hashtable] $AmOperationsLookupTable,
    [Parameter(ValueFromPipeline=$true, Mandatory=$true)] [object] $EseEvent
    ) {

    if (! $AmOperationsLookupTable.ContainsKey($EseEvent.DatabaseName)) {
        return $null
    }

    $OperationsForDatabase = $AmOperationsLookupTable[$EseEvent.DatabaseName]

    if ($EseEvent.Type -ne "Dismount") {
        $LookupTableGroup = "FINISH"
    } else {
        $LookupTableGroup = "START"
    }

    if (! $OperationsForDatabase[$LookupTableGroup].ContainsKey($EseEvent.SourceMachine)) {
        return $null
    }
    
    $OperationsForCopy = $OperationsForDatabase[$LookupTableGroup][$EseEvent.SourceMachine]

    return $OperationsForCopy | 
        where { ((RoundTimeDownToSecond $_.TimeRecoveryStarted) -le $EseEvent.TimeCreated) -and
                ((RoundTimeDownToSecond $_.TimeRecoveryEnded) -ge $EseEvent.TimeCreated) } |
        select -first 1
}

function MergeEseTimingEventsFromServer(
    [Parameter(Mandatory=$true)] [HashTable] $AmOperationsLookupTable,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [String] $Server, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {
        
    Process {
        Write-Host ($CollectOverMetrics_LocalizedStrings.res_0017 -f $Server)
        
        $XPathFilter = BuildXPathQueryString -StartTime $StartTime -EndTime $EndTime -Provider ESE -EventIds 103,105,301
        
        Get-WinEvent -ComputerName $Server -LogName "Application" -FilterXPath $XPathFilter -ErrorAction SilentlyContinue |
            ParseEseEvent |
            foreach {
                $EseEvent = $_
                $op = FindOperationMatchingEseEvent -AmOperationsLookupTable $AmOperationsLookupTable -EseEvent $EseEvent
                if ($op) { 
                    switch ($EseEvent.Type) {
                        ("Mount") {
                            $op.EseMountTiming += $EseEvent.EseTiming + ";; "
                        }
                        ("Dismount") {
                            $op.EseDismountTiming += $EseEvent.EseTiming + ";; "
                        }
                        ("ReplayedLog") {
                           $op.ReplayedLogs++;
                        }
                    }
                }
            }
    }
} # MergeStoreMountTimingEventsFromServer



function MergeEseTimingEventsFromServers(
    [Parameter(Mandatory=$true)] [HashTable] $AmOperationsLookupTable,
    [Parameter(Mandatory=$true)] [String[]] $Servers, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {
    
    $Servers | MergeEseTimingEventsFromServer -AmOperationsLookupTable $AmOperationsLookupTable -StartTime $StartTime -EndTime $EndTime

}


function GetAmOperationsDataForServers(
    [Parameter(Mandatory=$true)] [String[]] $Servers, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {

    $AmOperationsData = GetPamOperationalChannelStatsForServers -Servers $Servers -StartTime $StartTime -EndTime $EndTime
    if ($AmOperationsData) {
        MergeLossReportsFromServers -Servers $Servers -AmOperationsData $AmOperationsData -StartTime $StartTime -EndTime $EndTime
        if ($IncludeExtendedEvents) {
            $AmOperationsLookupTable = @{}
            $AmOperationsData |
                where {$_.DatabaseName -ne $null} |
                Group-Object DatabaseName |
                foreach {
                    $AmOperationsLookupTable[$_.Name] = @{ 
                        START  = ($_.Group | Group-Object -AsHashTable ActiveOnStart);
                        FINISH = ($_.Group | Group-Object -AsHashTable ActiveOnFinish)
                    }
                }
            
            MergeStoreMountTimingEventsFromServers -Servers $Servers -AmOperationsLookupTable $AmOperationsLookupTable -StartTime $StartTime -EndTime $EndTime
            MergeEseTimingEventsFromServers -Servers $Servers -AmOperationsLookupTable $AmOperationsLookupTable -StartTime $StartTime -EndTime $EndTime
        }
    }
    
    return $AmOperationsData  
}

##################################################################
### Report writing functions
##################################################################


### The CSV file that stores the raw operation stats will mean that everything's a
### string, this is the list of fields that are actually numeric.  We'll convert them
### to doubles when we read them in.

$NumericCounters = @("DurationOutage","DurationDismount","DurationBcs",
                     "DurationAcll","DurationMount","DurationOther",
                     "AcllCopiedLogs", "ReplayedLogs", "LostLogs", "LostBytes","AcllFinalReplayQueue" )
$CounterDataElements = @("Average","Maximum","90th Percentile","Quartiles")

### These fields are the 'identity' elements for each group in each dataset, they form the
### leftmost columns
$IdentityFields = @("Action Type", "Action Trigger", "Action Reason")

### These fields form the first summary data elements
$OverallSummaryFields = @("Total", "Failures", "Under 30s", "Over 30s", "Lossy Mount")

###
### Importing the data from the raw CSV files.  We go to special lengths to handle the
### The error messages from ActiveManager operations are generally in a format like this:
###
###     An Active Manager operation failed. Error: The database action failed. Error: An 
###     error occurred while trying to select a database copy for possible activation. 
###     Error: The database 'AA-EXM14-01 MBX Store 108' was not mounted because errors 
###     occurred either while validating database copies for possible activation, or while
###     attempting to activate another copy. Detailed error(s): 
###    
###     sinex14mbxc414: 
###                   An Active Manager operation failed with a transient error. Please
###                   retry the operation. Error: MapiExceptionNetworkError: Unable to 
###                   make admin interface connection to server. (hr=0x80040115, ec=-2147221227)
###     Diagnostic context:
###         Lid: 12696   dwParam: 0x6D9      Msg: EEInfo: Generation Time: 2009-10-03 07:05:00:411
###         ... <etc> ...
###   
###     sinex14mbxc402: 
###                   Database copy 'AA-EXM14-01 MBX Store 108' is in a failed state on server
###                   'SINEX14MBXC402.southpacific.corp.microsoft.com'. Reason: The Microsoft
###                   Exchange Replication service failed to talk to the local Information Store
###                   service. This often means that the Information Store service is not running.
###                   Error:MapiExceptionNetworkError: Unable to make admin interface connection
###                   to server. (hr=0x80040115, ec=-2147221227)
###     Diagnostic context:
###         Lid: 12696   dwParam: 0x6D9      Msg: EEInfo: Generation Time: 2009-10-03 07:03:41:96
###         ... <etc> ...
###   
### So it starts with a "preamble" that giving the most general information about the
### failure; this is the section that leads up to "Detailed error(s):".  The preamble
### message itself is made by concatenating the sequence of errors from the lower levels
### together (each level ended by the "Error: " string).  Beyond there we get details that
### may include records for each copy that was rejected during a failover.
###
### We can take the error message and split it into these two sections (preamble vs detail),
### and then further split the preamble section into the sequence of error messages.
###
### Before doing that, we eliminate as much of the text that has names of databases or servers,
### so we can treat them all as being equivalent to each other.
###
function ImportCsvData(
    [Parameter(Mandatory=$true)] [String[]] $CsvFiles,
    [Parameter(Mandatory=$true)] [ScriptBlock] $ReportFilter
    ) {
    $FoundOperationIds = @{}
    
    $CsvFiles |
        foreach { write-host ($CollectOverMetrics_LocalizedStrings.res_0005 -f $_); $_ } |
        Import-Csv -ErrorAction:SilentlyContinue |
        where $ReportFilter |
        where {(! [String]::IsNullOrEmpty($_.ActiveOnStart)) -and (! [String]::IsNullOrEmpty($_.ActiveOnFinish))} |
        foreach {
            if ( ! $FoundOperationIds.ContainsKey($_.UniqueOperationId) ) {
                $FoundOperationIds[$_.UniqueOperationId] = 1
                write-output $_
            }
        } |
        foreach { 
            # Convert strings that have specific database or server names with the generic "<blah>" so
            # we can treat them as all being equivalent (to batch them into groups).  Remember that all
            # replace statements are working with reg-ex's and so the '$' symbols in the
            # substitution text (ie the right-most argument) are important: use either single-quoted
            # strings or escape the '$' character there.
            $CanonicalMessage = ( $_.ErrorMessage -replace "(\s)'[^']+'", '$1<blah>' `
                                                  -replace "length of \d+ logs", "length of <blah> logs" `
                                                  -replace " C:(\w|\d|\\)+", " <blah>" `
                                                  -replace "(<blah> for|attempting to start a replication instance for|failed for database) [^.]+\.", '$1 <blah>.' `
                                                  -replace "\[Server: [^]]+\]", "[Server: <blah>]" `
                                                  -replace "on server \S+. ", "on server <blah>. " `
                                                  -replace "(Incremental seeding of database) .+\\.+ (encountered an error)", '$1 <blah> $2' `
                                                  -replace "(The automatic database operation on database) .+ (was attempt)", '$1 <blah> $2' `
                                                  -replace "\(database=[^)]*\)", '(database=<blah>)')
            ($Preamble, $Details) = [regex]::Split($CanonicalMessage, " Detailed error\(s\): ")
            # The MapiExceptionNetworkError errors have a set of details like LIDs following that exception
            # we'll just drop that portion for these summary reports
            $Preamble = $Preamble -replace "(MapiExceptionNetworkError):.*", '$1: <details>' 
            Add-Member -InputObject $_ -MemberType NoteProperty -Force -Name Preamble -Value ([regex]::Split($Preamble, " Error: "))
            Add-Member -InputObject $_ -MemberType NoteProperty -Force -Name Details -Value ([regex]::Split($Details, " \\n "))

            # Convert these fields to doubles
            foreach ($field in $numericCounters) {
                if ($_ | get-member $field) {
                    $_.$field = [double] $_.$field
                }
            }

            # Convert the times from strings to DateTimes
            $_.TimeRecoveryStarted = if ($_.TimeRecoveryStarted) { [DateTime]$_.TimeRecoveryStarted } else { $null }
            $_.TimeRecoveryEnded = if ($_.TimeRecoveryEnded) { [DateTime]$_.TimeRecoveryEnded } else { $null }
        
            Write-Output $_
    }
}

### Perform our analysis on a failover classes.  For each of the counters that we're
### interested in, we calculate averages, maxima, and some of the spread of results
### (to give a better picture of where our results fall - is the average bad because
### of generally bad results, or were there some serious outliers).
###

# Functions that we'll use when we're processing the data
        
function FindElement(
    [Object[]] $Data,
    [double] $fraction
    ) {
    $ElementIndex = [int] ($fraction * $Data.Count)
    if ($ElementIndex -ge $Data.Count) {
        return 0
    }
    return $Data[$ElementIndex] 
}
        
function CreateDataObject(
    [Parameter(Mandatory=$true)] [String] $Property
    ) {
    $OutputObject = New-Object object
    Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name "Counter" -Value $Property
    Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name "Average" -Value "0"
    Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name "Maximum" -Value "0"
    Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name "90th Percentile" -Value "0"
    Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name "Quartiles" -Value "0 : 0 : 0"
    return $OutputObject
}

function DetermineDataPoints (
    [Parameter(Mandatory=$true)] [String[]] $Properties,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [object] $Data
    ) {
    
    Begin {
        $PropertyValues = @{}
        Foreach ($Property in $Properties) {
            $PropertyValues[$Property] = @()
        }
    }
    
    ### Gather the values into separate groups for each of the properties that we're capturing
    Process {
        Foreach ($Property in $Properties) {
            if ($Data | get-member $Property) {
                $PropertyValues[$Property] += $Data.$Property
            }
        }
    }
    
    ### For each of the groups of property values that we gathered, measure the statistics that
    ### we're interested in.  We can only do that in the 'End' block because it is only then that
    ### we've captured all the information that we need.
    End {
    
        Foreach ($Property in $Properties) {
        
            $OutputObject = CreateDataObject -Property $Property
        
            if ($PropertyValues[$Property]) {
                
                $PropertyValues[$Property] = $PropertyValues[$Property] | Sort -Descending
     
                # We sorted in descending order, so we count 10% of the way in to find
                # the 90th percentile.  Likewise, we count 25%, 50%, and 75% to find the
                # quartiles.
                $NinetythPercentile = FindElement -Data $PropertyValues[$Property] -Fraction 0.1
                $Quartiles = @(0.25,0.5,0.75) | % { FindElement -Data $PropertyValues[$Property] -Fraction $_ }
                $QuartileString = [String]::Join(" : ", @($Quartiles | %{$_.ToString("0.###")}) )
            
                $MaxAndAverage = $PropertyValues[$Property] | measure -Average -Maximum
            
                $OutputObject.Average = $MaxAndAverage.Average.ToString("0.###")
                $OutputObject.Maximum = $MaxAndAverage.Maximum.ToString("0.###")
                $OutputObject."90th Percentile" = $NinetythPercentile.ToString("0.###")
                $OutputObject.Quartiles = $QuartileString
            }
            
            Write-Output $OutputObject
        }
    }
}

function CalculateStatsForGroup (
    [Parameter(Mandatory=$true)] [string] $ActionType,
    [Parameter(Mandatory=$true)] [string] $ActionTrigger,
    [Parameter(Mandatory=$true)] [string] $ActionReason,
    [Parameter(Mandatory=$true)] [Object[]] $DataGroup    
    ) {
        
    $Successes = @($DataGroup | where {$_.Result -eq "Success"})
    $FailureCount = $DataGroup | where {$_.Result -eq "Failed" } | measure
           
    $Result = New-Object object
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Action Type" -Value $ActionType
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Action Trigger" -Value $ActionTrigger
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Action Reason" -Value $ActionReason
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Total" -Value $DataGroup.Count
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Failures" -Value $FailureCount.Count
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Under 30s" -Value ($Successes | where {$_.DurationOutage -le 30} | measure).Count
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Over 30s" -Value ($Successes | where {$_.DurationOutage -gt 30} | measure).Count
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Lossy mount" -Value ($Successes | where {$_.LostLogs -ne 0} | measure).Count
            
    $DataBreakdown = $Successes | DetermineDataPoints -Properties $NumericCounters
                        
    Foreach ($DataPoint in $DataBreakdown) {
        Foreach ($Element in $CounterDataElements) {
            Add-Member -InputObject $Result -MemberType NoteProperty -Name "$($Datapoint.Counter) $Element" -Value $Datapoint.$Element
        }
    }
    
    return $result     
}

### Given the raw data, build each into groups for each category, calculate our summary stats and  
Function GroupAndProcessData(
    [Parameter(Mandatory=$true)] [Object[]] $Data
    ) {
    
    write-host "Processing totals"
    write-output ( CalculateStatsForGroup -DataGroup $Data -ActionType "Total" -ActionTrigger "Total" -ActionReason "Total" )
    
    $Data |
        Group -Property ActionCategory,ActionInitiator,ActionReason |
        Sort -Property Name |
        Foreach {
            write-host "Processing '$($_.Name)'"
            
            ($ActionType,$ActionTrigger,$ActionReason) = $_.Name.Split(",",4) | Foreach { $_.Trim() }
            write-output ( CalculateStatsForGroup -DataGroup $_.Group -ActionType $ActionType -ActionTrigger $ActionTrigger -ActionReason $ActionReason )
        }
}

function FormatPerfSummaryTable (
    [Parameter(Mandatory=$true)] [Object[]] $Data
    ) {
    
    write-host $CollectOverMetrics_LocalizedStrings.res_0007
    
    $colGroupsHtml = "<Colgroup Span = $($IdentityFields.Count)></Colgroup>" +
    "<Colgroup Span = $($OverallSummaryFields.Count)></Colgroup>" +
    [String]::Join("", ($NumericCounters | foreach {"<Colgroup Span = $($CounterDataElements.count)></Colgroup>"} ))

    $NumericCounterColumns = foreach ($counter in $NumericCounters) { foreach ($element in $CounterDataElements) { "$counter $element" } }
    $AllColumns = $IdentityFields + $OverallSummaryFields + $NumericCounterColumns
    $tableHeadingsHtml = "<TR>" +
                         [string]::Join("", ( $AllColumns | % { "<th>$_</th>" } ) ) +
                         "</TR>"
    
    $tableRows = GroupAndProcessData -Data $Data |
        Foreach -Begin { $row = 0 } -Process {
            $datum = $_
            $rowClass = if ( ($row % 2) -eq 0 ) { "Light" } else { "Dark" }
            $row += 1
            
            return "<TR class =`"$rowClass`">" +
                [String]::Join("", ( $AllColumns | % { "<td>$($datum.$_)</td>" } ) ) +
                "</TR>"
        }
        
    $tableRowsHtml = [String]::Join("", $tableRows)
        
    return "<table>" +
           $colGroupsHtml +
           $tableHeadingsHtml +
           $tableRows +
           "</table>"
}

###
### Produce tables showing the common top-level error messages, and the most common reasons
### that copies were rejected during failovers.
###
function FormatCommonErrorsTables (
    [Parameter(Mandatory=$true)] [Object[]] $Data
    ) {
    
    function IsCiFailover(
        [Parameter(Mandatory=$true)] [object]$element
        ) {
        return $element.ActionInitiator -eq "Admin" -and
               $element.ActionCategory  -eq "Move"  -and
               $element.ActionReason    -eq "FailureItem"
    }
    
    write-host $CollectOverMetrics_LocalizedStrings.res_0008
    
    $Failures = $Data | where { $_.Result -eq "Failed" }
    if (! $Failures) {
        write-host $CollectOverMetrics_LocalizedStrings.res_0009
        return "<H3><P ALIGN=`"CENTER`">No failed operations found in the data</P></H3>"
    }
    
    $CommonFailureTitle = "<H3><P ALIGN=`"CENTER`">Common top-level failures for operations</P></H3>"
    $CommonFailureTable = 
        $Failures |
        Group {$_.Preamble[-1]} -NoElement |
        select -Property Count,Name |
        Sort Count -Descending |
        ConvertTo-Html -Fragment
        
    $IndividualNodeFailureMessages = $Failures |
        where {
            $_.Preamble[2] -and
            $_.Preamble[2] -eq "An error occurred while trying to select a database copy for possible activation."
        } |
        foreach {
            $SourceEvent = $_
            $StartingServer = $SourceEvent.ActiveOnStart.Split(".") | select -first 1
            $_.Details |
                foreach { $_.Trim() } |
                foreach -Begin { $ServerName = $null } -Process {
                    if ($ServerName) {
                
                        $CopyType = if ($ServerName -eq $StartingServer) { "Source" } else { "Target" }
                
                        $output = New-Object Object
                        Add-Member -InputObject $output -MemberType NoteProperty -Name Message -Value $_
                        Add-Member -InputObject $output -MemberType NoteProperty -Name Server -Value $ServerName
                        Add-Member -InputObject $output -MemberType NoteProperty -Name CopyType -Value $CopyType
                        Add-Member -InputObject $output -MemberType NoteProperty -Name Database -Value $SourceEvent.DatabaseName
                        Add-Member -InputObject $output -MemberType NoteProperty -Name Time -Value $SourceEvent.TimeRecoveryStarted
                        Add-Member -InputObject $output -MemberType NoteProperty -Name OperationId -Value $SourceEvent.UniqueOperationId
                        Add-Member -InputObject $output -MemberType NoteProperty -Name SourceEvent -value $SourceEvent
                
                        Write-Output $output
                
                        $ServerName = $null
                
                    } elseif ($_ -match "^(\w+):$") {
                        $ServerName = $matches[1]
                    }
                }
        } 
 
    # The Group operator joins the fields that make up the group names with a ", "
    # When we're processing these names, we use this Reg-Ex for splitting them.
    $GroupNameSplitterRegex = [regex]", "
    
    $CommonNodeFailureTitle = "<H3><P ALIGN=`"CENTER`">Common reasons specific copies were rejected during failovers</P></H3>"
    $CommonNodeFailureMessages = $IndividualNodeFailureMessages |
        group CopyType,Message -NoElement |
        select -Property Count,Name |
        sort count -Descending |
        foreach {
            ($SourceOrTarget, $Message) = $GroupNameSplitterRegex.Split($_.Name, 2)
            $o = new-object Object
            $o | Add-Member -MemberType NoteProperty -Name Count -Value $_.Count 
            $o | Add-Member -MemberType NoteProperty -Name "Source / Target" -Value $SourceOrTarget
            $o | Add-Member -MemberType NoteProperty -Name "Error Message" -Value $Message
            return $o
        } |
        ConvertTo-Html -Fragment
        
    return $CommonFailureTitle + $CommonFailureTable + $CommonNodeFailureTitle + $CommonNodeFailureMessages
}


function FormatSlowOperationsTable (
    [Parameter(Mandatory=$true)] [Object[]] $Data
    ) {
    $ColumnsForSlowOperations = @(
        "ActionCategory",
        "ActionInitiator",
        "ActionReason",
        "Result",
        "DurationOutage",
        "DurationDismount",
        "DurationBcs",
        "DurationAcll",
        "DurationMount",
        "DurationOther",
        "NumberOfAttempts",
        "AcllCopiedLogs",
        "LostLogs",
        "ReplayedLogs",
        "LostBytes",
        "AcllFinalReplayQueue",
        "StoreMountProgress",
        "DatabaseName",
        "TimeRecoveryStarted",
        "TimeRecoveryEnded",
        "ActiveOnStart",
        "ActiveOnFinish",
        "PAMServer",
        "BcsTimingDetails",
        "AcllTimingDetails",
        "MountRpcTimingDetails",
        "StoreMountLids",
        "EseDismountTiming",
        "EseMountTiming"
    )
    
    $SlowOperations = $Data |
        where  { $_.DurationOutage -gt 30 } | 
        sort -Descending ActionCategory,ActionInitiator,DurationOutage |
        select $ColumnsForSlowOperations |
        ConvertTo-Html -Fragment |
        foreach { $_ -replace "--newline--","<br/>" }
        
    return $SlowOperations
}

###
### Format the results into a summary report.  This method contains definitions
### of the constant parts of the HTML page (style definitions, table explaining
### the different fields), and then calls FormatPerfSummaryTable and
### FormatCommonErrorsTables to produce the tables that contain the data.
###
function WriteHtmlReport(
    [Parameter(Mandatory=$true)] [Object[]] $Data,
    [Parameter(Mandatory=$true)] [String] $ReportPath,
    [Parameter(Mandatory=$true)] [String] $ReportTitle
    ) {
        
    $titleHtml = "<title>Failover summary report</title>"
    $styleHtml = @"
    	<style type="text/css">
		COLGROUP {
			border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; padding: 7px;
		}
		TABLE {
			border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse; padding: 7px; 
		}
		TBODY {
			border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;
		}
		TH {
			border-width: 1px; border-style: solid; border-color: black; background-color:#666699; color: white; padding:7px
		}
		TD {
			border-width: 1px; border-style: solid; border-color: black; background-color: #99AADD; padding: 7px
		}
		TR.Light TD {
			background-color: #BBCCDD; color: black; padding: 7px
		}
		TR.Dark TD {
			background-color: #AABBDD; color: black; padding: 7px
		}
	</style>
"@
    
	$headingHtml = '<H3><P ALIGN="CENTER">' + $ReportTitle + '</P></H3>'
        
    $summaryTableHtml = FormatPerfSummaryTable -Data $Data
    
    # Gather the operations that took more than 30s
    $slowOperationsTitle =  '<H3><P ALIGN="CENTER">Operations taking more than 30 seconds</P></H3><P ALIGN="CENTER">To see all the operations, load the CSV files in Excel and choose "Format as table".</P>'
    $slowOperationsHtml = FormatSlowOperationsTable -Data $Data
    
    $CommonErrorReport = FormatCommonErrorsTables -Data $Data
    
    ConvertTo-Html -Head ($titleHtml + $styleHtml) -Body ($headingHtml + $summaryTableHtml) -PostContent ($CommonErrorReport + $slowOperationsTitle + $slowOperationsHtml) |
        Out-File $ReportPath
    
}


function LoadExchangeSnapin
{
    if (! (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:SilentlyContinue) )
    {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
    }
}



###################################################################
###  Entry point for the script itself
###################################################################

$ReportTimeStamp = (Get-Date).ToString('yyyy_MM_dd_HH_mm_ss')


if ($ReportPath -and !(Test-Path -PathType Container $ReportPath)) {
    mkdir $ReportPath
}


if ($PsCmdlet.ParameterSetName -eq "Default") {

    if ( (!$StartTime) -and (!$EndTime) ) {
        $EndTime = [DateTime]::Now
        $StartTime = $EndTime.AddDays(-1)
    } elseif (!$StartTime) {
        $StartTime = $EndTime.AddDays(-1)
    } elseif (!$EndTime) {
        $EndTime = $StartTime.AddDays(1)
    }

    if ($EndTime -le $StartTime) {
        throw ($CollectOverMetrics_LocalizedStrings.res_0010 -f $StartTime,$EndTime)
    }
    
    if ($RawOutput -and ($GenerateHtmlReport -or $MergeCsvFiles)) {
        throw $CollectOverMetrics_LocalizedStrings.res_0011
    }

    LoadExchangeSnapin

    $DagsForProcessing = Get-DatabaseAvailabilityGroup $DatabaseAvailabilityGroup -ErrorAction SilentlyContinue
    if (! $DagsForProcessing) {
        throw ($CollectOverMetrics_LocalizedStrings.res_0012 -f $DatabaseAvailabilityGroup)
    }

    $DagReportFiles = @()

    Foreach ($Dag in $DagsForProcessing) {
    
        if ($Dag.Servers) {
            # For LogMiner, we just return the raw data
            if ($RawOutput) {
        
                # For the RawOutput format, we filter out any times that it returns nulls,
                # and flatten the lists of returned results into a single big list of values.
                GetAmOperationsDataForServers -Servers $Dag.Servers -StartTime $StartTime -EndTime $EndTime |
                     Where { $_ } |
                     Foreach { $_ | Write-Output }
            
            } else {
        
                $ReportFileName = "$ReportPath\FailoverReport.$Dag.$ReportTimeStamp.csv"
                GetAmOperationsDataForServers -Servers $Dag.Servers -StartTime $StartTime -EndTime $EndTime  |
                    Export-Csv -NoTypeInformation $ReportFileName -ErrorAction SilentlyContinue
                $DagReportFiles += $ReportFileName
            
            }
        }
    }
    Write-Host
    Write-Host ($CollectOverMetrics_LocalizedStrings.res_0013 + [String]::Join("`n  ", $DagReportFiles))
    Write-Host
}


if (($PsCmdlet.ParameterSetName -eq "Summarise") -or $GenerateHtmlReport -or $ShowHtmlReport -or $MergeCsvFiles) {

    $AmOperationsData = @()
    $HtmlReportTitle = "Database *over summary statistics"
    
    if ($Database) {
        $InitialReportFilter = $ReportFilter
        $MatchingDatabaseGuids = Get-MailboxDatabase $database -ErrorAction:SilentlyContinue | 
            Foreach { $_.Guid }
        $ReportFilter = { 
                (& $InitialReportFilter) -and 
                ($MatchingDatabaseGuids.Contains($_.DatabaseGuid))
            }.GetNewClosure()
    }
    
    if ($PsCmdlet.ParameterSetName -eq "Summarise") {        
        $AmOperationsData = ImportCsvData -CsvFiles $SummariseCsvFiles -ReportFilter $ReportFilter
        $HtmlReportTitle +="<br/>Data compiled from:<br/>" + [String]::Join("<br/>", $SummariseCsvFiles)
    } else {
        $AmOperationsData = ImportCsvData -CsvFiles $DagReportFiles -ReportFilter $ReportFilter
    	$HtmlReportTitle += "<br/>From $($StartTime.ToString("F")) ... To ... $($EndTime.ToString("F"))"
    }


    if ($MergeCsvFiles) {
        $MergedCsvFileName = "$ReportPath\MergedFailoverReports.$ReportTimeStamp.csv"
        $AmOperationsData | Export-Csv -NoTypeInformation $MergedCsvFileName
        Write-Host
        Write-Host ($CollectOverMetrics_LocalizedStrings.res_0014 -f $MergedCsvFileName)
    }

    if (($PsCmdlet.ParameterSetName -eq "Summarise") -or $GenerateHtmlReport -or $ShowHtmlReport) {
        Write-Host
        Write-Host $CollectOverMetrics_LocalizedStrings.res_0015
        $HtmlSummaryReportName = "$ReportPath\FailoverSummary.$ReportTimeStamp.htm"
        WriteHtmlReport -Data $AmOperationsData -ReportPath $HtmlSummaryReportName -ReportTitle $HtmlReportTitle
        Write-Host ($CollectOverMetrics_LocalizedStrings.res_0016 -f $HtmlSummaryReportName)
        
        if ($ShowHtmlReport) {
			$ie = new-object -comobject "InternetExplorer.Application"
			$ie.visible = $true
			$ie.navigate($HtmlSummaryReportName)    
		}
    }
}

# SIG # Begin signature block
# MIIacgYJKoZIhvcNAQcCoIIaYzCCGl8CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUwjplIo+x9G35+Otjr56tkNT+
# d4KgghUvMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# Ok6T/ZsGj62K+Gh+O7p5QpM5SfXQXuxwjUJ1xYJVkBu1VWEwggTDMIIDq6ADAgEC
# AhMzAAAAKzkySMGyyUjzAAAAAAArMA0GCSqGSIb3DQEBBQUAMHcxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQTAeFw0xMjA5MDQyMTEyMzRaFw0xMzEyMDQyMTEyMzRaMIGz
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMQ0wCwYDVQQLEwRN
# T1BSMScwJQYDVQQLEx5uQ2lwaGVyIERTRSBFU046QzBGNC0zMDg2LURFRjgxJTAj
# BgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCmtjAOA2WuUFqGa4WfSKEeycDuXkkHheBwlny+
# uV9iXwYm04s5uxgipS6SrdhLiDoar5uDrsheOYzCMnsWeO03ODrxYvtoggJo7Ou7
# QIqx/qEsNmJgcDlgYg77xhg4b7CS1kANgKYNeIs2a4aKJhcY/7DrTbq7KRPmXEiO
# cEY2Jv40Nas04ffa2FzqmX0xt00fV+t81pUNZgweDjIXPizVgKHO6/eYkQLcwV/9
# OID4OX9dZMo3XDtRW12FX84eHPs0vl/lKFVwVJy47HwAVUZbKJgoVkzh8boJGZaB
# SCowyPczIGznacOz1MNOzzAeN9SYUtSpI0WyrlxBSU+0YmiTAgMBAAGjggEJMIIB
# BTAdBgNVHQ4EFgQUpRgzUz+VYKFDFu+Oxq/SK7qeWNAwHwYDVR0jBBgwFoAUIzT4
# 2VJGcArtQPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1w
# UENBLmNybDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# dDATBgNVHSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAfsywe+Uv
# vudWtc9z26pS0RY5xrTN+tf+HmW150jzm0aIBWZqJoZe/odY3MZjjjiA9AhGfCtz
# sQ6/QarLx6qUpDfwZDnhxdX5zgfOq+Ql8Gmu1Ebi/mYyPNeXxTIh+u4aJaBeDEIs
# ETM6goP97R2zvs6RpJElcbmrcrCer+TPAGKJcKm4SlCM7i8iZKWo5k1rlSwceeyn
# ozHakGCQpG7+kwINPywkDcZqJoFRg0oQu3VjRKppCMYD6+LPC+1WOuzvcqcKDPQA
# 0yK4ryJys+fEnAsooIDK4+HXOWYw50YXGOf6gvpZC3q8qA3+HP8Di2OyTRICI08t
# s4WEO+KhR+jPFTCCBbwwggOkoAMCAQICCmEzJhoAAAAAADEwDQYJKoZIhvcNAQEF
# BQAwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jv
# c29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9y
# aXR5MB4XDTEwMDgzMTIyMTkzMloXDTIwMDgzMTIyMjkzMloweTELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWljcm9zb2Z0IENv
# ZGUgU2lnbmluZyBQQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCy
# cllcGTBkvx2aYCAgQpl2U2w+G9ZvzMvx6mv+lxYQ4N86dIMaty+gMuz/3sJCTiPV
# cgDbNVcKicquIEn08GisTUuNpb15S3GbRwfa/SXfnXWIz6pzRH/XgdvzvfI2pMlc
# RdyvrT3gKGiXGqelcnNW8ReU5P01lHKg1nZfHndFg4U4FtBzWwW6Z1KNpbJpL9oZ
# C/6SdCnidi9U3RQwWfjSjWL9y8lfRjFQuScT5EAwz3IpECgixzdOPaAyPZDNoTgG
# hVxOVoIoKgUyt0vXT2Pn0i1i8UU956wIAPZGoZ7RW4wmU+h6qkryRs83PDietHdc
# pReejcsRj1Y8wawJXwPTAgMBAAGjggFeMIIBWjAPBgNVHRMBAf8EBTADAQH/MB0G
# A1UdDgQWBBTLEejK0rQWWAHJNy4zFha5TJoKHzALBgNVHQ8EBAMCAYYwEgYJKwYB
# BAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQU/dExTtMmipXhmGA7qDFvpjy8
# 2C0wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwHwYDVR0jBBgwFoAUDqyCYEBW
# J5flJRP8KuEKU5VZ5KQwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5taWNy
# b3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQuY3Js
# MFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwDQYJKoZIhvcN
# AQEFBQADggIBAFk5Pn8mRq/rb0CxMrVq6w4vbqhJ9+tfde1MOy3XQ60L/svpLTGj
# I8x8UJiAIV2sPS9MuqKoVpzjcLu4tPh5tUly9z7qQX/K4QwXaculnCAt+gtQxFbN
# LeNK0rxw56gNogOlVuC4iktX8pVCnPHz7+7jhh80PLhWmvBTI4UqpIIck+KUBx3y
# 4k74jKHK6BOlkU7IG9KPcpUqcW2bGvgc8FPWZ8wi/1wdzaKMvSeyeWNWRKJRzfnp
# o1hW3ZsCRUQvX/TartSCMm78pJUT5Otp56miLL7IKxAOZY6Z2/Wi+hImCWU4lPF6
# H0q70eFW6NB4lhhcyTUWX92THUmOLb6tNEQc7hAVGgBd3TVbIc6YxwnuhQ6MT20O
# E049fClInHLR82zKwexwo1eSV32UjaAbSANa98+jZwp0pTbtLS8XyOZyNxL0b7E8
# Z4L5UrKNMxZlHg6K3RDeZPRvzkbU0xfpecQEtNP7LN8fip6sCvsTJ0Ct5PnhqX9G
# uwdgR2VgQE6wQuxO7bN2edgKNAltHIAxH+IOVN3lofvlRxCtZJj/UBYufL8FIXri
# lUEnacOTj5XJjdibIa4NXJzwoq6GaIMMai27dmsAHZat8hZ79haDJLmIz2qoRzEv
# mtzjcT3XAH5iR9HOiMm4GPoOco3Boz2vAkBq/2mbluIQqBC0N1AI1sM9MIIGBzCC
# A++gAwIBAgIKYRZoNAAAAAAAHDANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZImiZPy
# LGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQDEyRN
# aWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMDcwNDAzMTI1
# MzA5WhcNMjEwNDAzMTMwMzA5WjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCfoWyx39tIkip8ay4Z4b3i48WZ
# USNQrc7dGE4kD+7Rp9FMrXQwIBHrB9VUlRVJlBtCkq6YXDAm2gBr6Hu97IkHD/cO
# BJjwicwfyzMkh53y9GccLPx754gd6udOo6HBI1PKjfpFzwnQXq/QsEIEovmmbJNn
# 1yjcRlOwhtDlKEYuJ6yGT1VSDOQDLPtqkJAwbofzWTCd+n7Wl7PoIZd++NIT8wi3
# U21StEWQn0gASkdmEScpZqiX5NMGgUqi+YSnEUcUCYKfhO1VeP4Bmh1QCIUAEDBG
# 7bfeI0a7xC1Un68eeEExd8yb3zuDk6FhArUdDbH895uyAc4iS1T/+QXDwiALAgMB
# AAGjggGrMIIBpzAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBQjNPjZUkZwCu1A
# +3b7syuwwzWzDzALBgNVHQ8EBAMCAYYwEAYJKwYBBAGCNxUBBAMCAQAwgZgGA1Ud
# IwSBkDCBjYAUDqyCYEBWJ5flJRP8KuEKU5VZ5KShY6RhMF8xEzARBgoJkiaJk/Is
# ZAEZFgNjb20xGTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1p
# Y3Jvc29mdCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eYIQea0WoUqgpa1Mc1j0
# BxMuZTBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUH
# AQEESDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
# L2NlcnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDATBgNVHSUEDDAKBggrBgEFBQcD
# CDANBgkqhkiG9w0BAQUFAAOCAgEAEJeKw1wDRDbd6bStd9vOeVFNAbEudHFbbQwT
# q86+e4+4LtQSooxtYrhXAstOIBNQmd16QOJXu69YmhzhHQGGrLt48ovQ7DsB7uK+
# jwoFyI1I4vBTFd1Pq5Lk541q1YDB5pTyBi+FA+mRKiQicPv2/OR4mS4N9wficLwY
# Tp2OawpylbihOZxnLcVRDupiXD8WmIsgP+IHGjL5zDFKdjE9K3ILyOpwPf+FChPf
# wgphjvDXuBfrTot/xTUrXqO/67x9C0J71FNyIe4wyrt4ZVxbARcKFA7S2hSY9Ty5
# ZlizLS/n+YWGzFFW6J1wlGysOUzU9nm/qhh6YinvopspNAZ3GmLJPR5tH4LwC8cs
# u89Ds+X57H2146SodDW4TsVxIxImdgs8UoxxWkZDFLyzs7BNZ8ifQv+AeSGAnhUw
# ZuhCEl4ayJ4iIdBD6Svpu/RIzCzU2DKATCYqSCRfWupW76bemZ3KOm+9gSd0BhHu
# diG/m4LBJ1S2sWo9iaF2YbRuoROmv6pH8BJv/YoybLL+31HIjCPJZr2dHYcSZAI9
# La9Zj7jkIeW1sMpjtHhUBdRBLlCslLCleKuzoJZ1GtmShxN1Ii8yqAhuoFuMJb+g
# 74TKIdbrHk/Jmu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUW
# a8kTo/0xggStMIIEqQIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQIT
# MwAAAJ0ejSeuuPPYOAABAAAAnTAJBgUrDgMCGgUAoIHGMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBRmoOod6m3zyCsfw/doi6lqT4kWADBmBgorBgEEAYI3AgEMMVgw
# VqAugCwAQwBvAGwAbABlAGMAdABPAHYAZQByAE0AZQB0AHIAaQBjAHMALgBwAHMA
# MaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3
# DQEBAQUABIIBAGnigkMY+sf2Ymz5OunCrLqRws6YUTHwCZO9GkcREJvqdJ307X9o
# 1rmy2L8ksRVY+iToESZQBAoHqY0rHBkJFcZBf1zoAbHAIGVNSc2m41gH3PK+ldzE
# oGLc3G7ZvCq8XiOFJlFnmFg4mK3IwgAa6wci69Ytm5UG15CcVH20+gyRmTZyE10O
# SizuLHiz40MVEwidxrmaSjuEtSM3s9VRvOwrcEesFJNQZSwabRhFxvDJo4V+rQ/O
# cdNpBAyWhUYfMXzaT1b61KWJlFqgnl/ALchrGpJOwfC/+squkUW4op8zYIMIcRAU
# jipT132felKx8PLz4L9wHVzdtcHRAEZFh+6hggIoMIICJAYJKoZIhvcNAQkGMYIC
# FTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAAKzkySMGy
# yUjzAAAAAAArMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0xMzAxMDgwODQ2NDZaMCMGCSqGSIb3DQEJBDEWBBTc
# z1wzFrGvZxmztjXgFHX9VvdkXjANBgkqhkiG9w0BAQUFAASCAQBjmeW0XtqVTAv1
# xd73AVsmVODHnT9G20Z4QIEvrPHjWGp4pqynSsfm8kW5HXYJtFDbVYbQ5+b7MWGK
# ynVW6YKTtSHrOE7nZfwdkaxEvZY0VFagtTryId0iUHRE9tyg8LYf1fhzwlc0L8/L
# STl1dxO+mQo795MR2FmVQXDvtCSCmUfrWod23tWPHGy5aB5ZfLyV9CF58ZURmtkv
# TJx3uK7NgOdSL0JitlYVDtnqUhw1Y5UYcZ6oI7/9FoHH0PUTY6U7TThcfRyqGik0
# I4vF8eJ2XFKRbo7w6UtbSDHdFcVGL1CSo39quaZSJlCUgzGYWHpcxj2EpReEI9Ib
# LpBz21F+
# SIG # End signature block
