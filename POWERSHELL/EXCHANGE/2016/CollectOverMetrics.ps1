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
    [String] $Database,
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

function GetFqdn([string] $serverName)
{
	return (Get-ExchangeServer $serverName).Fqdn
}

$GeneratePamStatisticsBlock = {
    param
    (
        [String]$XPathFilter,
        [HashTable]$ActionEventLookupTable
    )

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

    $wEvents = Get-WinEvent -LogName "Microsoft-Exchange-HighAvailability/Operational" -FilterXPath $XPathFilter -ErrorAction SilentlyContinue
    $groups = $wEvents | Group -Property {$_.Properties[0].Value}
    $statistics = $groups | %{
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
        Add-Member -InputObject $entry -MemberType NoteProperty -Name MoveComment          -Value $null
        Add-Member -InputObject $entry -MemberType NoteProperty -Name SubReason            -Value $null
        $OutputObject = $entry
        $OutputObject.UniqueOperationId = $EventGroup.Name
        $_.Group | 
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
								$OutputObject.MoveComment = $eventXML.MoveComment;                      
								if ( $eventXML.MoveComment.Trim() -match "Moved as part of database redistribution \(RedistributeActiveDatabases\.ps1\)\." )
								{
									# The DB was moved by the rebalancer script (RedistributeActiveDatabases.ps1).
									# NOTE: The previous value was "Cmdlet".
									$OutputObject.ActionReason = "Rebalance"
								}
							}
                            if ( ( $eventXML.GetElementsByTagName("SubReason") | Measure-Object ).Count -gt 0 )
                            {
                                $OutputObject.SubReason = $eventXML.SubReason;  
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
        $OutputObject
    } # Foreach Group

    $statistics
}

function GetPamOperationalChannelStatsForServers(
    [Parameter(Mandatory=$true)] [String[]] $Servers, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {

    $PamStats = @()
    $Jobs = @()

    $XPathFilter = BuildXPathQueryString -StartTime $StartTime -EndTime $EndTime -EventIds ($ActionEventLookupTable.Keys)

    foreach ($Server in $Servers) {
        $fqdn = GetFqdn $Server
        $Jobs += Invoke-Command -AsJob -ComputerName $fqdn -ScriptBlock $GeneratePamStatisticsBlock -ArgumentList $XPathFilter, $ActionEventLookupTable
    }

    foreach($job in $Jobs) {
        Write-Host ($CollectOverMetrics_LocalizedStrings.res_0000 -f $($job.Location))
        $statistics = Receive-Job $job -Wait
        $PamStats += $statistics
    }
    
    write-host ($CollectOverMetrics_LocalizedStrings.res_0001 -f $PamStats.Count)

    return $PamStats
}

$ParseAcclPerfEventsBlock = {
    param
    (
        [String]$XPathFilter
    )

    $AcllTimingDetailsFieldsMap = @{
            AcllQueuedOpStartElapsedTime="AQOpStET";
            AcquireSuspendLockElapsedTime="ASusLkET";
            CopyLogsOverallElapsedTime="CLOET";
            CopyStatus="CS";
            IsNewCopierInspectorCreated="INCIC";
            IsAcllFoundDeadConnection="IAFDC";
            IsAcllCouldNotControlLogCopier="IACNCLC";
        }

    $wEvents = Get-WinEvent -LogName "Microsoft-Exchange-HighAvailability/Operational" -FilterXPath $XPathFilter -ErrorAction SilentlyContinue
    foreach($AcllPerfEvent in $wEvents) {
        $eventXML = ([xml] $AcllPerfEvent.ToXml()).Event.UserData.EventXML
    
        $OutputObject = New-Object -TypeName "Object"
        Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name UniqueOperationId -Value $eventXML.UniqueId
        Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name NumberOfLogsCopied -Value $eventXML.NumberOfLogsCopied
        Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name ReplayQueueLengthAcllEnd -Value $eventXML.ReplayQueueLengthAcllEnd

        $RemappedAcllFields =
            foreach ($key in $AcllTimingDetailsFieldsMap.Keys) {
                "$($AcllTimingDetailsFieldsMap[$key])=$($eventXML.$key)"
            }
        $RemappedAcllFieldsJoined = [String]::Join("; ", $RemappedAcllFields)
        Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name RemappedAcllFields -Value $RemappedAcllFieldsJoined

        $OutputObject
    }
}


function MergeAcllPerfEventsFromParsedData(
    [Parameter(Mandatory=$true)] [Object[]] $AmOperationsData,
    [Parameter(Mandatory=$true)] [Object[]] $AcllPerfData
    ) {
    
    Process {
        $AcllPerfData |
            foreach {
                $AcllPerfEvent = $_
                
                :FIND_MATCHING_OP foreach ($op in $AmOperationsData) { 
                
                    if ($op.UniqueOperationId -eq $AcllPerfEvent.UniqueOperationId) {
                    
                        $op.AcllCopiedLogs = $AcllPerfEvent.NumberOfLogsCopied
                        $op.AcllFinalReplayQueue = $AcllPerfEvent.ReplayQueueLengthAcllEnd
                        
                            
                        if ($op.AcllTimingDetails) {
                            $op.AcllTimingDetails += ";; "
                        }
                        $op.AcllTimingDetails += $AcllPerfEvent.RemappedAcllFields
                    
                        break FIND_MATCHING_OP
                        
                    }
                }
            }
    }
}

$ParseMountRpcPerfEventsBlock = {
        param
    (
        [String]$XPathFilter
    )

    $MountRpcTimingDetailsFieldsMap = @{
        PreMountQueuedOpStartElapsedTime = "PMQOStET";
        PreMountQueuedOpExecutionElapsedTime = "PMQOExET";
        StoreMountElapsedTime = "SMET"
    }

    $wEvents = Get-WinEvent -LogName "Microsoft-Exchange-HighAvailability/Operational" -FilterXPath $XPathFilter -ErrorAction SilentlyContinue
    foreach($RpcPerfEvent in $wEvents) {
        $eventXml = ([xml] $RpcPerfEvent.ToXml()).Event.UserData.EventXML
        $OutputObject = New-Object -TypeName "Object"
        Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name DatabaseGuid -Value $eventXML.DatabaseGuid
        Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name MachineName -Value $RpcPerfEvent.MachineName
        Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name TimeCreated -Value $RpcPerfEvent.TimeCreated

        $RemappedMountRpcFields =
            foreach ($key in $MountRpcTimingDetailsFieldsMap.Keys) {
                "$($MountRpcTimingDetailsFieldsMap[$key])=$($eventXML.$key)"
            }

        $RemappedMountRpcFieldsJoined = [String]::Join("; ", $RemappedMountRpcFields)
        Add-Member -InputObject $OutputObject -MemberType NoteProperty -Name RemappedMountRpcFields -Value $RemappedMountRpcFieldsJoined

        $OutputObject
    }
}

function MergeMountRpcPerfEventsFromParsedData(
    [Parameter(Mandatory=$true)] [Object[]] $AmOperationsData,
    [Parameter(Mandatory=$true)] [Object[]] $RpcPerfData
    ) {
    

    
    Process {
        $RpcPerfData |
            foreach {
                $RpcPerfEvent = $_
                
                :FIND_MATCHING_OP foreach ($op in $AmOperationsData) { 
                
                    if (($op.DatabaseGuid -eq $RpcPerfEvent.DatabaseGuid) -and 
                        ($op.ActiveOnFinish -eq $RpcPerfEvent.MachineName) -and
                        ($op.TimeRecoveryStarted -le $RpcPerfEvent.TimeCreated) -and
                        ($op.TimeRecoveryEnded -ge $RpcPerfEvent.TimeCreated)) {

                        if ($op.MountRpcTimingDetails) {
                            $op.MountRpcTimingDetails += ";; "
                        }
                        $op.MountRpcTimingDetails += $RpcPerfEvent.RemappedMountRpcFields
              
                        break FIND_MATCHING_OP
                        
                    }
                }
            }
    }
}


$ParseLossReportEventsBlock = {
    param
    (
        [String]$XPathFilter
    )

    $AcllLossId = 324
    $GranularLossId = 333
    $wEvents = Get-WinEvent -LogName "Microsoft-Exchange-HighAvailability/Operational" -FilterXPath $XPathFilter -ErrorAction SilentlyContinue
    foreach($LossReport in $wEvents) { 
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
            
            $OutputObject
        }
    }
}

function MergeLossReportsFromParsedData(
    [Parameter(Mandatory=$true)] [Object[]] $AmOperationsData,
    [Parameter(Mandatory=$true)] [Object[]] $LossReports
    ) {
    
    Process {
        $LossReports |
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
    
    $Jobs = @()
    $FQDNs = @()
    $XPathFilter = BuildXPathQueryString -StartTime $StartTime -EndTime $EndTime -EventIds 324,333

    foreach($Server in $Servers) {
        $fqdn = GetFqdn $Server
        $FQDNs += $fqdn
        $Jobs += Invoke-Command -AsJob -ComputerName $fqdn -ScriptBlock $ParseLossReportEventsBlock -ArgumentList $XPathFilter
    }

    foreach($job in $Jobs) {
        Write-Host ($CollectOverMetrics_LocalizedStrings.res_0002 -f $($job.Location))
        $LossReports = Receive-Job $job -Wait
        if($LossReports) { MergeLossReportsFromParsedData -LossReports $LossReports -AmOperationsData $AmOperationsData }
    }

    $Jobs = @()
    $XPathFilter = BuildXPathQueryString -StartTime $StartTime -EndTime $EndTime -EventIds 336

    foreach($fqdn in $FQDNs) {
        $Jobs += Invoke-Command -AsJob -ComputerName $fqdn -ScriptBlock $ParseAcclPerfEventsBlock -ArgumentList $XPathFilter
    }

    foreach($job in $Jobs) {
        $AcllPerfData = Receive-Job $job -Wait
        if($AcllPerfData) { MergeAcllPerfEventsFromParsedData -AcllPerfData $AcllPerfData -AmOperationsData $AmOperationsData }
    }

    $Jobs = @()
    $XPathFilter = BuildXPathQueryString -StartTime $StartTime -EndTime $EndTime -EventIds 337

    foreach ($fqdn in $FQDNs) { 
        $Jobs += Invoke-Command -AsJob -ComputerName $fqdn -ScriptBlock $ParseMountRpcPerfEventsBlock -ArgumentList $XPathFilter
    }

    foreach($job in $Jobs) {
        $RpcPerfData = Receive-Job $job -Wait
        if($RpcPerfData){ MergeMountRpcPerfEventsFromParsedData -RpcPerfData $RpcPerfData -AmOperationsData $AmOperationsData }
    }
}

    
### Parse the array of bytes into a series of LID/Tick pairs.  Each item in the pair is a 32-bit integer,
### and we have to be careful with how we treat the order of the bytes when forming the values.
###
$ParseStoreMountEventBlock = {
    param
    (
        [String]$XPathFilter
    )

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

    $wEvents = Get-WinEvent -LogName "Application" -FilterXPath $XPathFilter -ErrorAction SilentlyContinue 
    foreach($MountEvent in $wEvents) {
        if ($MountEvent) {
           $data = new-object -TypeName Object
           Add-Member -InputObject $data -MemberType NoteProperty -Name "DatabaseName" -Value $MountEvent.Properties[0].Value
           Add-Member -InputObject $data -MemberType NoteProperty -Name "SourceMachine" -Value $MountEvent.MachineName
           Add-Member -InputObject $data -MemberType NoteProperty -Name "TimeCreated" -Value $MountEvent.TimeCreated
           Add-Member -InputObject $data -MemberType NoteProperty -Name "LidProgression" -Value ""
           Add-Member -InputObject $data -MemberType NoteProperty -Name "LidTickPairs" -Value ""
           if ( $MountEvent.Properties.Count -gt 6 ) {
               # The event has got the Lid/Tick data, parse it
               #($data.LidTickPairs, $data.LidProgression) = ParseStoreLidByteArray $MountEvent.Properties[6].Value
               [System.Byte[]]$bytes = $MountEvent.Properties[6].Value
               if($bytes.Count -ne 0) {
                    if (($bytes.Count % 8) -ne 0) { throw ($CollectOverMetrics_LocalizedStrings.res_0003 -f $bytes.count) }
                    $LidTickStrings = @()
                    $ProgressString = ""
                    $PreviousTick = $null
                    $LidTickCount = [int]($bytes.Count / 8)
   
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

                    $data.LidTickPairs = [String]::Join("; ", $LidTickStrings)
                    $data.LidProgression = $ProgressString
               }
           }
           $data
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

$ParseEseEventBlock = {
    param
    (
        [String]$XPathFilter
    )

    $wEvents = Get-WinEvent -LogName "Application" -FilterXPath $XPathFilter -ErrorAction SilentlyContinue 
    foreach($EseEvent in $wEvents) {
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

            $data
        }
    }
    
}

function FindOperationMatchingEseEvent (
    [Parameter(Mandatory=$true)] [Hashtable] $AmOperationsLookupTable,
    [Parameter(Mandatory=$true)] [object] $EseEvent
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

function MergeEseTimingEventsFromParsedData(
    [Parameter(Mandatory=$true)] [HashTable] $AmOperationsLookupTable,
    [Parameter(Mandatory=$true)] [Object[]] $EseData
    ) {
        
    Process {
        $EseData |
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
}

function MergeStoreMountTimingEventsFromParsedData(
    [Parameter(Mandatory=$true)] [HashTable] $AmOperationsLookupTable,
    [Parameter(Mandatory=$true)] [Object[]] $MountData
    ) {
        
    Process {
        $MountData |
        foreach {
            $StoreMountEvent = $_
            $op = FindOperationMatchingStoreEvent -StoreEvent $StoreMountEvent -AmOperationsLookupTable $AmOperationsLookupTable
                
            if ($op) {
                $op.StoreMountLids += $StoreMountEvent.LidTickPairs + ";; "
                $op.StoreMountProgress += $StoreMountEvent.LidProgression + ";; "
            }
        }
    }
}

function MergeStoreMountAndEseTimingEventsFromServers(
    [Parameter(Mandatory=$true)] [HashTable] $AmOperationsLookupTable,
    [Parameter(Mandatory=$true)] [String[]] $Servers, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {
    
    $Jobs = @()
    $FQDNs = @()
    $XPathFilter = BuildXPathQueryString -StartTime $StartTime -EndTime $EndTime -Provider "MsExchangeIS Mailbox Store" -EventIds 9523

    foreach($Server in $Servers) {
        $fqdn = GetFqdn $Server
        $FQDNs += $fqdn
        $Jobs += Invoke-Command -AsJob -ComputerName $fqdn -ScriptBlock $ParseStoreMountEventBlock -ArgumentList $XPathFilter
    }
    
    foreach($job in $Jobs) {
        Write-Host ($CollectOverMetrics_LocalizedStrings.res_0004 -f $($job.Location))
        $MountData = Receive-Job $job -Wait
        if($MountData) { MergeStoreMountTimingEventsFromParsedData -MountData $MountData -AmOperationsLookupTable $AmOperationsLookupTable }
    }

    $Jobs = @()
    $XPathFilter = BuildXPathQueryString -StartTime $StartTime -EndTime $EndTime -Provider ESE -EventIds 103,105,301

    foreach($fqdn in $FQDNs) {
        $Jobs += Invoke-Command -AsJob -ComputerName $fqdn -ScriptBlock $ParseEseEventBlock -ArgumentList $XPathFilter
    }

    foreach($job in $Jobs) {
        Write-Host ($CollectOverMetrics_LocalizedStrings.res_0017 -f $($job.Location))
        $EseData = Receive-Job $job -Wait
        if($EseData) { MergeEseTimingEventsFromParsedData -EseData $EseData -AmOperationsLookupTable $AmOperationsLookupTable }
    }
}

### 
### The flow for getting the data from ActiveManager's database operation events is:
### 
### 1 For each server in the DAG, read all the events with IDs defined in $ActionEventLookupTable
### 2 Group them by the unique operation IDs - each group now represents an operation
### 3 For each group, create an object that will hold the stats for it.
###   Each object will become a row in the table.
### 4 From each event in the group, populate the relevant fields in the object
### 
### Processing a group into a single object is in GeneratePamStatisticsBlock.
### Then we can continue by looking for the ACLL loss events that match each
### operation, matching by names & times in RTM, by operation ID in SP1.
###
### Finally, we merge in the store mount LID events (MergeStoreMountAndEseTimingEventsFromServers).
### Matching these into operations is trickier because the events are from the Application
### log which has only second-level resolution on its timestamps, rather than the tick-level
### resolution that the other channel's events have.  We round the times down to the nearest
### second for the comparisons.  This could mean that two events for the same database on the
### same server will each hit the same second interval, and we will just append these matches
### together.
###

function GetAmOperationsDataForServers(
    [Parameter(Mandatory=$true)] [String[]] $Servers, 
    [Parameter(Mandatory=$true)] [DateTime] $StartTime,
    [Parameter(Mandatory=$true)] [DateTime] $EndTime
    ) {

    $AmOperationsData = @(GetPamOperationalChannelStatsForServers -Servers $Servers -StartTime $StartTime -EndTime $EndTime)
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
            
            MergeStoreMountAndEseTimingEventsFromServers -Servers $Servers -AmOperationsLookupTable $AmOperationsLookupTable -StartTime $StartTime -EndTime $EndTime
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
            $_.TimeRecoveryStarted = if ($_.TimeRecoveryStarted) { [DateTime]::Parse( $_.TimeRecoveryStarted ) } else { $null }
            $_.TimeRecoveryEnded = if ($_.TimeRecoveryEnded) { [DateTime]::Parse( $_.TimeRecoveryEnded)  } else { $null }
        
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
    #Use out-null or else it will be returned to the screen
    mkdir $ReportPath | Out-Null
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
		
			# Filter pingable servers only to avoid hangs
			$servers = @( $Dag.Servers | where {[bool](Test-Connection (Get-ExchangeServer $_).Fqdn -Count 1 -quiet)} )
			
            # For LogMiner, we just return the raw data
            if ($RawOutput) {
        
                # For the RawOutput format, we filter out any times that it returns nulls,
                # and flatten the lists of returned results into a single big list of values.
                GetAmOperationsDataForServers -Servers $servers -StartTime $StartTime -EndTime $EndTime |
                     Where { $_ } |
                     Foreach { $_ | Write-Output }
            
            } else {
        
                $ReportFileName = "$ReportPath\FailoverReport.$Dag.$ReportTimeStamp.csv"
                GetAmOperationsDataForServers -Servers $servers -StartTime $StartTime -EndTime $EndTime  |
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
        $MatchingDatabaseGuids = @(Get-MailboxDatabase $Database -ErrorAction:SilentlyContinue | Foreach { $_.Guid })
        $ReportFilter = { 
                (& $InitialReportFilter) -and 
                ($MatchingDatabaseGuids -contains $_.DatabaseGuid)
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
# MIIdrAYJKoZIhvcNAQcCoIIdnTCCHZkCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUClsKWa3FByVFiCTKn8NgxGId
# Ph6gghhkMIIEwzCCA6ugAwIBAgITMwAAAJzu/hRVqV01UAAAAAAAnDANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTMw
# WhcNMTcwNjMwMTkyMTMwWjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjU4NDctRjc2MS00RjcwMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzCWGyX6IegSP
# ++SVT16lMsBpvrGtTUZZ0+2uLReVdcIwd3bT3UQH3dR9/wYxrSxJ/vzq0xTU3jz4
# zbfSbJKIPYuHCpM4f5a2tzu/nnkDrh+0eAHdNzsu7K96u4mJZTuIYjXlUTt3rilc
# LCYVmzgr0xu9s8G0Eq67vqDyuXuMbanyjuUSP9/bOHNm3FVbRdOcsKDbLfjOJxyf
# iJ67vyfbEc96bBVulRm/6FNvX57B6PN4wzCJRE0zihAsp0dEOoNxxpZ05T6JBuGB
# SyGFbN2aXCetF9s+9LR7OKPXMATgae+My0bFEsDy3sJ8z8nUVbuS2805OEV2+plV
# EVhsxCyJiQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFD1fOIkoA1OIvleYxmn+9gVc
# lksuMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAFb2avJYCtNDBNG3nxss1ZqZEsphEErtXj+MVS/RHeO3TbsT
# CBRhr8sRayldNpxO7Dp95B/86/rwFG6S0ODh4svuwwEWX6hK4rvitPj6tUYO3dkv
# iWKRofIuh+JsWeXEIdr3z3cG/AhCurw47JP6PaXl/u16xqLa+uFLuSs7ct7sf4Og
# kz5u9lz3/0r5bJUWkepj3Beo0tMFfSuqXX2RZ3PDdY0fOS6LzqDybDVPh7PTtOwk
# QeorOkQC//yPm8gmyv6H4enX1R1RwM+0TGJdckqghwsUtjFMtnZrEvDG4VLA6rDO
# lI08byxadhQa6k9MFsTfubxQ4cLbGbuIWH5d6O4wggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBLIwggSuAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBxjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUMZQedDbhe7OmXxVuSAkOhYEg+pcwZgYKKwYB
# BAGCNwIBDDFYMFagLoAsAEMAbwBsAGwAZQBjAHQATwB2AGUAcgBNAGUAdAByAGkA
# YwBzAC4AcABzADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdl
# IDANBgkqhkiG9w0BAQEFAASCAQAoqYhjqyPObdGcbBJjTGub0V44aPlrKLt5yO/w
# iZPNRpObPlExRnKSHkdh18V/IOpUyFqoIrDBv5jRNTWqDjpKV268blwQzQZGQ8Wd
# qXESc+8Zx/49J+7Dvx103FtgRUaUnuMsfR609FUDipCfmHAELhp6eCTfxJAGDhaO
# GjaDWkgZHgmXRUiNdRsRQn4qDn7n1AeJTYlk1lRIoAgXj/MlE6HMe+8WxEhx1xr5
# v7+YCeoGRaSzW1el4yi/8iJXazbYlyy3KWyi7q3NTG3SIlniT7kEe/IA6+kEJ7q6
# 3AAMGbp8hjvr4mHbMYF0G1PhAIPcVvqzh6h3GAUvsRcsLApSoYICKDCCAiQGCSqG
# SIb3DQEJBjGCAhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQIT
# MwAAAJzu/hRVqV01UAAAAAAAnDAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsG
# CSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTYwOTAzMTg0NDI4WjAjBgkqhkiG
# 9w0BCQQxFgQUCICPLd5nkFA5tNzjWSi7WoB55A4wDQYJKoZIhvcNAQEFBQAEggEA
# hUg0i54F50oF4MxBEvuZ0jSYtduArydSt4xypu8KWM8+Ssjo+rB6NyS1fdRq4V3Y
# 7gTxJouQMNXbNBz3Wy/u8z8W6Y7AzivmI+mlgqqjH2DaSToCN0VHfkbEeyM/6MU0
# hGlO1t23CD/BgZR+52PscR5Ww5CXdlXrCP1Q2LzYlAUWFZEKyoNB5pQ3Lli5QNHg
# EHVZuvmynWnZwqQBYKS7gE8Il2NYWl6dK92fTkLHAf6wpt9l/t+Iwfu2LVVbrFt7
# sYUI/hLp0LGnPrxPsjc0D9wxWdj27jAKNVpWTqVYLsJ/ehVzGw7uLhdaEzw6iT3T
# YiIw2mwAXG/haLGfIBOmKQ==
# SIG # End signature block
