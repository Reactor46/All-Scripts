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
# MIIaxQYJKoZIhvcNAQcCoIIatjCCGrICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUwjplIo+x9G35+Otjr56tkNT+
# d4KgghWCMIIEwzCCA6ugAwIBAgITMwAAAHD0GL8jIfxQnQAAAAAAcDANBgkqhkiG
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
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBK0wggSp
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAAAymzVMhI1xOFV
# AAEAAADKMAkGBSsOAwIaBQCggcYwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFGag
# 6h3qbfPIKx/D92iLqWpPiRYAMGYGCisGAQQBgjcCAQwxWDBWoC6ALABDAG8AbABs
# AGUAYwB0AE8AdgBlAHIATQBlAHQAcgBpAGMAcwAuAHAAcwAxoSSAImh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAAGGN
# qlMGX4Ee4Tw2p+H99z4th7nn5WYTXlXceovhUTcwQtMw0iFHgQyZMrnOp4m3JmzE
# IQjGLubbAHw/kKxzqGt2Q9ztFDfq/d3fFALl6uhOP34nKTIk2KtkgasNSatPEr7v
# +acuYZ/v6gtVZwH15Bjp9Q3hS+RqaP11S06vmy7eWpRVodfRaTsDG+xndsqXRFdf
# xmVpuac2Cc0cp0T1VeFN/VMiB5bec23QiBdxUIYw+npaw4gMQJKFvuREyRNmo4AD
# 3vAu+n5fwLlhDEd/NUxKVU5oDQRmoi/hxabFVjikInAPfrUs2D+askBt+25wy1AD
# bbmd/cxN3n6pWrnzQKGCAigwggIkBgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhN
# aWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAABw9Bi/IyH8UJ0AAAAAAHAwCQYF
# Kw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkF
# MQ8XDTE1MDQxMDAyNTczMVowIwYJKoZIhvcNAQkEMRYEFFHjd9FqGN1kJPuPWv7Q
# 5uNSX+xTMA0GCSqGSIb3DQEBBQUABIIBAHCIxo+o+1DcaJHBBJ9LAPEOtwLavXiT
# IcPuYdeVJ7tXVQ8ciQP2lFlSScRTE//vlH1BQ0mjNpgIobxf2yWGRaUBUb6Y67cw
# TXWB1fThP7pR91fFIYpAuNDyy94jJ5Wm4zVAzBLyVFMl9snocnQkRy8p8DdtvITI
# 2HxsGSHXxNRjtYl5PEAxFP8gBooN7sPq9INijzDZszPG2mNNqTjBNBYY6LPRnMYF
# WGI0leO5TkOyrqZIGP+Yf0R1UZfnLWT71ySdJlOiPVzXgTU7FMzMGcg+aMGVN8oP
# WWNx4bJXZBl/ajNrumXV4d4dRwCPhzud//9d/FWl13sHjl1b2OpJZcU=
# SIG # End signature block
