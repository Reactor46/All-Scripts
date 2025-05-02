<#
.EXTERNALHELP CollectReplicationMetrics-help.xml
#>

# Copyright (c) Microsoft Corporation. All rights reserved.


param (
    [Parameter(ParameterSetName="Dag",Mandatory=$true)] [object] $DagName,
    
    [Parameter(ParameterSetName="ServerList",Mandatory=$true)] [object] $Servers,
    
    [Parameter(ParameterSetName="Summarise",Mandatory=$true)] [string[]] $SummariseFiles,
    
    [Parameter(Mandatory=$true)] [string] $ReportPath,
    
    [Parameter(ParameterSetName="Dag",Mandatory=$true)]
    [Parameter(ParameterSetName="ServerList",Mandatory=$true)] [TimeSpan] $Duration,
    
    [Parameter(ParameterSetName="Dag",Mandatory=$true)]
    [Parameter(ParameterSetName="ServerList",Mandatory=$true)] [TimeSpan] $Frequency,
    
    [ValidateSet("CollectAndReport", "CollectOnly","ProcessOnly")] [string] $Mode = "CollectAndReport",
    
    [switch] $MoveFilesToArchive,
    
    [switch] $LoadExchangeSnapin
    )


function LoadExchangeSnapin
{
    if (! (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction:SilentlyContinue) )
    {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
    }
}

### Grouping these functions together in a script block that I can then send to the
### jobs that actually execute them.  
$SingleServerCollectionMethods = {

    $CounterToFieldMap = @{
        'database mounted' = 'Mounted';
        'failed' = 'Failed';
        'initializing' = 'Initializing';
        'failedsuspended' = 'FailedSuspended';
        'resynchronizing' = 'Resynchronizing';
        'disconnected' = 'Disconnected';
        'suspended' = 'Suspended';
        'log generation rate on source (generations/sec)' = 'LogGenerationRate';
        'log copy kb/sec' = 'LogCopyRate';
        'log inspection rate (generations/sec)' = 'LogInspectionRate';
        'log replay rate (generations/sec)' = 'LogReplayRate';
        'copyqueuelength' = 'CopyQueueLength';
        'copygenerationnumber' = 'CopyGenerationNumber';
        'inspectorgenerationnumber' = 'InspectorGenerationNumber';
        'replaygenerationnumber' = 'ReplayGenerationNumber';
        'truncatedgenerationnumber' = 'TruncatedGenerationNumber'
    }
    

    function CollectCountersFromServer (
        [Parameter(Mandatory=$true)] [string] $Server,
        [Parameter(Mandatory=$true)] [TimeSpan] $Duration,
        [Parameter(Mandatory=$true)] [TimeSpan] $Frequency
        ) {
    
        $FrequencySeconds = [int] $Frequency.TotalSeconds
        $SampleCount = [int] ($Duration.TotalSeconds / $FrequencySeconds)
    
        $CounterList = @(
            "\processor(_total)\% processor time",
            "\MSExchange Active Manager(*)\*",
            "\MSExchange Replication(*)\*"
        )
    
        return Get-Counter -Counter $CounterList -ComputerName $Server `
                           -MaxSamples $SampleCount -SampleInterval $FrequencySeconds -ErrorAction SilentlyContinue
    }
    
 
    function CreateDatabasePerfSampleObject (
        [Parameter(Mandatory=$true)] [String] $Server,
        [Parameter(Mandatory=$true)] [DateTime] $TimeStamp,
        [Parameter(Mandatory=$true)] [String] $Database
        ) {
 
        $DatabaseSample = New-Object -TypeName Object
        Add-Member -InputObject $DatabaseSample -MemberType NoteProperty -Name TimeStamp -Value $TimeStamp
        Add-Member -InputObject $DatabaseSample -MemberType NoteProperty -Name Server -Value $Server
        Add-Member -InputObject $DatabaseSample -MemberType NoteProperty -Name Database -Value $Database
    
        $CounterToFieldMap.Values | Foreach {
            Add-Member -InputObject $DatabaseSample -MemberType NoteProperty -Name $_ -Value 0
        }
    
        return $DatabaseSample
    }
    
    
    function RegroupSampleFromServer (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)] [Microsoft.PowerShell.Commands.GetCounter.PerformanceCounterSampleSet] $SampleSet,
        [Parameter(Mandatory=$true)] [string] $Server
        ) {
    
        Process {
        
            $SampleSet.CounterSamples |
                Group-Object -Property InstanceName |
                where { $_.Name -ne "_total" } |
                Foreach {
                    $DatabaseGroup = $_
                    $OutputObject = CreateDatabasePerfSampleObject -Server $Server -TimeStamp $SampleSet.TimeStamp -Database $DatabaseGroup.Name
                    $DatabaseGroup.Group |
                        Foreach {
                            $CounterSample = $_
                        
                            ## The counter names are like "\\sinprd0103mb039\\msexchange replication(apcprd01dg003-db020)\copygenerationnumber"
                            ## We want the last value, so we split the string on the slashes.  The extra
                            ## parameters are in case we get a path that has some slashes in the final part
                            ## of the name (although I'm not sure if it's possible for that to happen)
                            $Counter = $CounterSample.Path.Split("\",3,[System.StringSplitOptions]::RemoveEmptyEntries)[2]
                            
                            if ( $CounterToFieldMap.ContainsKey($Counter) ) {
                                $FieldName = $CounterToFieldMap[$Counter]
                                $OutputObject.$FieldName = $CounterSample.CookedValue
                            }
                        }
                    write-output $OutputObject 
                }
        }
    }
 
 
    function SaveCountersFromServer (
        [Parameter(Mandatory=$true)] [string] $Server,
        [Parameter(Mandatory=$true)] [string] $OutputFile,
        [Parameter(Mandatory=$true)] [TimeSpan] $Duration,
        [Parameter(Mandatory=$true)] [TimeSpan] $Frequency
        ) {
        
        CollectCountersFromServer -Server $Server -Duration $Duration -Frequency $Frequency |
               RegroupSampleFromServer -Server $Server |
               Export-Csv -NoTypeInformation -Path $OutputFile
 
    }
     
} # End of SingleServerCollectionMethods



function ProcessCountersFromServers (
    [Parameter(Mandatory=$true)] [string[]] $Servers,
    [Parameter(Mandatory=$true)] [string] $ReportPath,
    [Parameter(Mandatory=$true)] [TimeSpan] $Duration,
    [Parameter(Mandatory=$true)] [TimeSpan] $Frequency
    ) {
    
    $jobs = @()
    
    $CollectionTimeStamp = (Get-Date).ToString('yyyy_MM_dd_HH_mm_ss')
    
    $SingleServerJob = {
        $args = $($Input)
        SaveCountersFromServer -Duration $args.Duration -Frequency $args.Frequency -Server $args.Server -OutputFile $args.OutputFile
    }
    
    foreach ($server in $Servers) {

        $OutputFile = "$ReportPath\CounterData.$Server.$CollectionTimeStamp.csv"
        
        Write-Host ($CollectReplicationMetrics_LocalizedStrings.res_0000 -f $server,$OutputFile)
    
        $ArgumentObject = new-object Object
        Add-Member -InputObject $ArgumentObject -MemberType NoteProperty -Name Server -Value $Server
        Add-Member -InputObject $ArgumentObject -MemberType NoteProperty -Name OutputFile -Value $OutputFile
        Add-Member -InputObject $ArgumentObject -MemberType NoteProperty -Name Frequency -Value $Frequency
        Add-Member -InputObject $ArgumentObject -MemberType NoteProperty -Name Duration -Value $Duration
        
        $jobs += Start-Job -InitializationScript $SingleServerCollectionMethods -ScriptBlock $SingleServerJob -InputObject $ArgumentObject
        
	    # Pause for a moment before spinning up the next job to avoid overloading the machine
	    Start-Sleep -Seconds 10
    }

    Write-Host $CollectReplicationMetrics_LocalizedStrings.res_0001
    Wait-Job $Jobs
    Write-Host $CollectReplicationMetrics_LocalizedStrings.res_0002
    
}


function NewCalculatedPropertyDefinition (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [object] $Property,
            
    [Parameter(Mandatory=$true)]
    [object] $FirstObject
    ) {
    
    Process {
         
        if ($Property -is [String]) {
            
            $MatchingProperties =
               Get-Member -InputObject $FirstObject -MemberType Property,NoteProperty $Property |
               Foreach { $_.Name } |
               where { $p = $_; !($PropertyDefinitions | where { $_["NAME"] -eq $p }) }
                
            if ($MatchingProperties) {
                foreach ($Property in $MatchingProperties) {
                    $NewDefinition = @{
                        NAME = "Delta " + $Property;
                        VALUE = { $_.$Property }.GetNewClosure();
                        PREVIOUS_VALUE = $null;
                        DIFFERENCE_FUNCTION = { param($current,$previous) $current - $previous }
                    }
                
                    write-output $NewDefinition
                }
            }
                
        } elseif ($Property -is [ScriptBlock]) {
            
            $NewDefinition = @{
                NAME = "Delta " + $Property.ToString();
                VALUE = $Property;
                PREVIOUS_VALUE = $null;
                DIFFERENCE_FUNCTION = { param($current,$previous) $current - $previous }
            }
                    
            write-output $NewDefinition
               
        } elseif ( ($Property -is [HashTable]) -and
                    $Property.ContainsKey("NAME") -and ($Property["NAME"] -is [String]) -and
                    $Property.ContainsKey("VALUE") -and ($Property["VALUE"] -is [ScriptBlock])) {
                    
            $NewDefinition = @{
                NAME = $Property["NAME"];
                VALUE = $Property["VALUE"];
                PREVIOUS_VALUE = $null;
                DIFFERENCE_FUNCTION = { param($current,$previous) $current - $previous }
            }
            if ($Property.ContainsKey("DIFFERENCE_FUNCTION")) {
               $NewDefinition["DIFFERENCE_FUNCTION"] = $Property["DIFFERENCE_FUNCTION"]
            }
                
            Write-Output $NewDefinition
            
        } else {
            throw ($CollectReplicationMetrics_LocalizedStrings.res_0008 -f $Property)
        } 
    }
} # end function NewCalculatedPropertyDefinition



function CalculateChanges (
    [Parameter(ValueFromPipeline=$true, Mandatory=$true)]
    [Object] $InputObject,
    
    [Parameter(Mandatory=$true,Position=0)]
    [Object[]] $Property,
    
    [String[]] $IncludeProperty
    ) {
    
    Begin {
        $PropertyDefinitions = @()
        $HasProcessedFirstValue = $false
    }
    
    Process {
        
        $OutputObject = $_
            
        if ($HasProcessedFirstValue) {
        
            foreach ($p in $PropertyDefinitions) {
                $CurrentValue = $InputObject | foreach $p["VALUE"]
                $Difference = &$p["DIFFERENCE_FUNCTION"] $CurrentValue $p["PREVIOUS_VALUE"]
                $OutputObject | Add-Member -MemberType NoteProperty -Name $p["NAME"] -Value $Difference
                $p["PREVIOUS_VALUE"] = $CurrentValue
            }
            
        } else {
        
            # Determine all the property definitions that we're going to collect
            $Property |
                NewCalculatedPropertyDefinition -FirstObject $InputObject |
                Foreach {
                    $NewDefinition = $_
                    if (! ($PropertyDefinitions | where {$_.Name -eq $NewDefinition["NAME"]}) ) {
                        $PropertyDefinitions += $NewDefinition                        
                        $OutputObject | Add-Member -MemberType NoteProperty -Name $NewDefinition["NAME"] -Value 0
                    }
                }
            
            foreach ($p in $PropertyDefinitions) {
                $CurrentValue = $InputObject | % $p["VALUE"]
                $p["PREVIOUS_VALUE"] = $CurrentValue
            }
            $HasProcessedFirstValue = $True
        }
        
        write-output $OutputObject
    }
}


$CopyStateCounters = @("Mounted", "Initializing", "Resynchronizing",
                       "Disconnected", "FailedSuspended", "Suspended", "Failed")

$GenNumberCountersToRateMap =
    @{ CalculatedSourceGeneration = "LogsGenerated";
       CopyGenerationNumber       = "LogsCopied";
       InspectorGenerationNumber  = "LogsInspected";
       ReplayGenerationNumber     = "LogsReplayed" }


function CalculatePerCopyStats (
    [Parameter(Mandatory=$true)]
    [String] $CopyName,
    
    [Parameter(Mandatory=$true)]
    [String] $Day,
    
    [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
    [object] $SampleData
    ) {
    
    Begin {
        $CopyStatObject = New-Object Object
            
        $CopyStatObject | Add-Member -MemberType NoteProperty -Name CopyName -Value $CopyName
        $CopyStatObject | Add-Member -MemberType NoteProperty -Name Date -Value $Day
        
        $CopyStatObject | Add-Member -MemberType NoteProperty -Value 0 -Name "Total Samples"
        
        foreach ($counter in $CopyStateCounters) {
            $CopyStatObject | Add-Member -MemberType NoteProperty -Value 0 -Name "$counter Samples"
        }
        
        $CopyStatObject | Add-Member -MemberType NoteProperty -Value 0 -Name "Healthy Samples"
        $CopyStatObject | Add-Member -MemberType NoteProperty -Value 0 -Name "Out-of-criteria Samples"
        $CopyStatObject | Add-Member -MemberType NoteProperty -Value 0 -Name "High CopyQueue Samples"
        $CopyStatObject | Add-Member -MemberType NoteProperty -Value 0 -Name "High InspectionQueue Samples"
        $CopyStatObject | Add-Member -MemberType NoteProperty -Value 0 -Name "High ReplayQueue Samples"
        $CopyStatObject | Add-Member -MemberType NoteProperty -Value 0 -Name "High GenerationRate Samples"            
    }
    
    Process {
    
        $CopyStatObject."Total Samples" += 1
        
        $CopyIsHealthy = $true
        Foreach ($counter in $CopyStateCounters) {
            if ($SampleData.$counter -eq 1) {
                $CopyStatObject."$counter Samples" += 1
                $CopyIsHealthy = $false
            }
        }
        
        if ($CopyIsHealthy) {
            $CopyStatObject."Healthy Samples" += 1
        }
        
        if ($SampleData.CopyQueueLength -gt 12) {
            $CopyStatObject."Out-of-criteria Samples" += 1
        }
        
        if ($SampleData.CalculatedCopyQueueLength -gt 12) {
            $CopyStatObject."High CopyQueue Samples" += 1
        }
        
        if ($SampleData.CalculatedInspectorQueueLength -gt 12) {
            $CopyStatObject."High InspectionQueue Samples" += 1
        }
        
        if ($SampleData.CalculatedReplayQueueLength -gt 30) {
            $CopyStatObject."High ReplayQueue Samples" += 1
        }
        
        if ($SampleData.LogsGenerated -gt 100) {
            $CopyStatObject."High GenerationRate Samples" += 1
        }
    }
    
    End {
        write-output $CopyStatObject
    }
}


function SanitizeData (
    [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
    [object] $SampleData
    ) {

    Process {
        # When a copy changes from active to passive, counters that were 0
        # now get the current generation numbers.  This means that it looks like
        # we've got massive copy/replay/inspection rates.  We smooth this out
        # by looking for the times that the rate is exactly the value of the
        # current generation number, and setting the rate to 0 in that case.
        #
        # In the other direction, when we get a bad set of counters, many will
        # temporarily go to zero, which makes the rate calculation go negative.
        # Another case is when we have a lossy failover and some generation
        # numbers go backwards.  In each case, we also set the rate counter
        # to zero.
        #
        foreach ($GenNumberCounter in $GenNumberCountersToRateMap.Keys) {
            $rateField = $GenNumberCountersToRateMap[$GenNumberCounter]
            if ($SampleData.$RateField -eq $SampleData.$GenNumberCounter -or
                $SampleData.$RateField -lt 0 ) {
                $SampleData.$RateField = 0
            }
        }
            
        # If the gap between samples was much more than a minute, we'll put the average
        # per-minute amounts for the logs generated etc etc
        if ($SampleData."TimeInterval (seconds)" -gt 90) {
            $IntervalInMinutes = $SampleData."TimeInterval (seconds)" / 60
            foreach ($rateField in $GenNumberCountersToRateMap.Values) {
                $SampleData.$rateField = $SampleData.$rateField / $IntervalInMinutes
            }
        }
        
        write-output $SampleData
    }
    
}


function AddQueueLengths (
    [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
    [object] $SampleData
    ) {
    
    Process {            
    
    
    # There are many times that we get zeroes for some but not all the counters
    # Specifically, the copy generation tends to get initialized to old data
    # when the instance gets started.  So if we see some of the gen-counters
    # are zero, we'll also set the copy generation number back to zero
    if ($SampleData.InspectorGenerationNumber -eq 0 -and $SampleData.ReplayGenerationNumber -eq 0) {
        $SampleData.CopyGenerationNumber = 0
    }
    
    # The scripts were configured to collect:
    #   * Copy, Inspection, and Replay generation numbers
    #   * CopyQueueLength counter (== Copier queue + Inspector queue)
    # From these, we'll calculate some of the other figures
    #   * Replay queue is Inspector gen - Replay gen
    #   * Inspector queue is Copy gen - Inspector gen
    #   * The real copier queue is the copy queue length counter - the inspector queue
    #   * The source generation number is the Copy gen + real copy queue length
    $SampleData | Add-Member -MemberType NoteProperty -Name CalculatedReplayQueueLength -Value ([int] 0)
    if ($SampleData.InspectorGenerationNumber -ne 0 -and $SampleData.ReplayGenerationNumber -ne 0) {
        $SampleData.CalculatedReplayQueueLength = [int]($SampleData.InspectorGenerationNumber - $SampleData.ReplayGenerationNumber)
    }
    
    $SampleData | Add-Member -MemberType NoteProperty -Name CalculatedInspectorQueueLength -Value ([int] 0)
    if ($SampleData.CopyGenerationNumber -ne 0 -and $SampleData.InspectorGenerationNumber -ne 0) {
        $SampleData.CalculatedInspectorQueueLength  = [int]($SampleData.CopyGenerationNumber - $SampleData.InspectorGenerationNumber)
    }
    
    $SampleData | Add-Member -MemberType NoteProperty -Name CalculatedCopyQueueLength -Value ([int] 0)
    if ($SampleData.CopyQueueLength -ne 0 -and $SampleData.CalculatedInspectorQueueLength -ne 0) {
        $SampleData.CalculatedCopyQueueLength = [int]($SampleData.CopyQueueLength - $SampleData.CalculatedInspectorQueueLength)
    }
    
    $SampleData | Add-Member -MemberType NoteProperty -Name CalculatedSourceGeneration -Value ([int] 0)
    if ($SampleData.CopyGenerationNumber -ne 0) {
        $SampleData.CalculatedSourceGeneration = [int]($SampleData.CalculatedCopyQueueLength + $SampleData.CopyGenerationNumber)
    }
    
    # For later filtering purposes, it's useful to break the time stamp down further
    # into both the date and the hour
    $SampleData | Add-Member -MemberType NoteProperty -Name Day -Value $_.TimeStamp.Date
    $SampleData | Add-Member -MemberType NoteProperty -Name Time -Value $_.TimeStamp.TimeOfDay
    
    write-output $SampleData
    }
}
    
       
function ImportCounterData (
    [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
    [object] $CsvFile
    ) {
    
    Process {
        import-csv $_ |
        foreach {
            # The timestamps in the CSV are just strings, they are more useful
            # to us as DateTimes
            $_.TimeStamp = [DateTime]$_.TimeStamp
            
            $_.CopyQueueLength =           [int]$_.CopyQueueLength
            $_.CopyGenerationNumber =      [int]$_.CopyGenerationNumber
            $_.InspectorGenerationNumber = [int]$_.InspectorGenerationNumber
            $_.ReplayGenerationNumber =    [int]$_.ReplayGenerationNumber 
            
            write-output $_
        }
    }
}


function ProcessOneServerDayOfCounterData (
    [Parameter(Mandatory=$true)]
    [object[]] $CounterFiles,
    
    [Parameter(Mandatory=$true)]
    [String] $Server,
    
    [Parameter(Mandatory=$true)]
    [String] $Day
    ) {

    # These definitions are used for the rate-of-change calculations that we'll
    # make from the raw data.  We're interested in:
    #
    #   TimeInterval - the amount of time between the samples
    #   LogsGenerated
    #   LogsCopied
    #   LogsInspected
    #   LogsReplayed
    
    $CalculatedValueDefinitions = @(
        @{ NAME = "TimeInterval (seconds)";
           VALUE = {$_.TimeStamp};
           DIFFERENCE_FUNCTION = {param($c,$p); return ($c-$p).TotalSeconds}
        }
    )
    
    $GenNumberCountersToRateMap.Keys | 
        foreach {
            $local:GenNumberCounter = $_
            $CalculatedValueDefinitions += @{
                NAME = $GenNumberCountersToRateMap[$GenNumberCounter];
                VALUE = { $_.$GenNumberCounter }.GetNewClosure()
            }
        }
    
    $PerCopyStats = @()
    
    $counterFiles |
        ImportCounterData |
        group Database,Server |
        foreach {
        
            # Within each copy, we'll calculate the extra data on each sample,
            # and then calculate the overall stats for that copy.  Then we
            # pass all the data back up to get written into the whole day's
            # data CSV.
            #
            # To do this, after we've done the first set of massaging and
            # calculations, we use "tee" to save off that data before we
            # then do the calculations that summarise the copy.  Then we
            # can pass that original saved data on to get written out to
            # a CSV with all the data.
            
            $CopyName = $_.Name -replace ", ", "\"
            $TempDataBuffer = @()
           
            $PerCopyStats += $_.Group |
                sort TimeStamp |
                AddQueueLengths |
                CalculateChanges -Property $CalculatedValueDefinitions |
                SanitizeData |
                Tee-Object -Variable TempDataBuffer |
                CalculatePerCopyStats -CopyName $CopyName -Day $Day 
            
            # Pass the whole set of data on to the CSV
            $tempDataBuffer | foreach { Write-Output $_ }
               
        } |
        Export-Csv CopyPerfData.$Server.$Day.csv -NoTypeInformation
        
        # And return the summary data back up to the caller:
        $PerCopyStats | foreach { Write-Output $_ } 
}


function GenerateReport (
    [Parameter(ParameterSetName="ServerList",Mandatory=$true)]  [string[]] $Servers,
    [Parameter(ParameterSetName="FileList",Mandatory=$true)]    [string[]] $FileList,
    [string] $DagName,
    [Parameter(Mandatory=$true)] [string] $ReportPath,
    [Parameter(ParameterSetName="ServerList",Mandatory=$true)] [TimeSpan] $Duration,
    [switch] $MoveFilesToArchive
    ) {
        
    # If we got a list of files, then trust that list; otherwise search for files of
    # the form "CounterData.<servername>.<time>.csv"
    
    if ($PsCmdlet.ParameterSetName -eq "ServerList") {
        $ProcessCutoffTime = (Get-Date) - $Duration
        $FilesToProcess = $Servers | 
            foreach { dir "$ReportPath\CounterData.$_.*.csv" }|
            where { ($mode -eq "CollectAndReport") -or ($_.LastWriteTime -lt $ProcessCutoffTime) } 
    } else {
        $FilesToProcess = $FileList | foreach { dir $_ }
    }
      
    
    $ReportTimeStamp = (Get-Date).ToString('yyyy_MM_dd_HH_mm_ss')
    if ($DagName) {
        $ReportSuffix = "$DagName.$ReportTimeStamp.csv"
    } else {
        $ReportSuffix = "$ReportTimeStamp.csv"
    }
    $ReportFilePath = "$ReportPath\HaReplPerfSummary.$ReportSuffix"
    
    write-host ($CollectReplicationMetrics_LocalizedStrings.res_0012 -f $ReportFilePath)
    
    # Group the files by day that they cover, and then within each group
    # group again by server name.    
    $FilesToProcess |
        group { $_.LastWriteTime.ToString("yyyy-MM-dd") } | 
        foreach {
            $Day = $_.Name
            write-host ($CollectReplicationMetrics_LocalizedStrings.res_0009 -f $Day)
            
            $_.Group |
                group {$_.Name.Split(".")[1]} |
                foreach {
                    $Server = $_.Name
                    write-host ($CollectReplicationMetrics_LocalizedStrings.res_0010 -f $Server)
                    
                    ProcessOneServerDayOfCounterData -CounterFiles $_.Group -Server $Server -Day $Day 
                }
        } |
        Export-Csv -NoTypeInformation $ReportFilePath
    
    if ($MoveFilesToArchive) {
        $ArchiveFolder = "$ReportPath\HaReplPerfArchive.$ReportSuffix"
        mkdir $ArchiveFolder
        move $FilesToProcess $ArchiveFolder
        $dirObj = Get-WmiObject -Query "Select * From Win32_Directory Where Name = '$($ArchiveFolder.Replace("\","\\"))'"
        $dirObj.Compress()
    }
}



###################################################################
###  Entry point for the script itself
###################################################################

Import-LocalizedData -BindingVariable CollectReplicationMetrics_LocalizedStrings -FileName CollectReplicationMetrics.strings.psd1
if ($LoadExchangeSnapIn) {
    LoadExchangeSnapin
}

switch($PsCmdlet.ParameterSetName) {

    ("Dag") {
        $DagsForReport = Get-DatabaseAvailabilityGroup $DagName -ErrorAction SilentlyContinue
        if (! $DagsForReport) {
            Throw ($CollectReplicationMetrics_LocalizedStrings.res_0004 -f $DagName)
        }
    }

    ("ServerList") {
        $ServersForReport = Get-MailboxServer $Servers -ErrorAction SilentlyContinue
        if (! $ServersForReport) {
            Throw ($CollectReplicationMetrics_LocalizedStrings.res_0011 -f $servers)
        }
    }
    
    ("Summarise") {
        $Mode = "ProcessOnly"
    }

}


if (! (Test-Path $ReportPath -PathType Container) ) {
    Throw ($CollectReplicationMetrics_LocalizedStrings.res_0005 -f $ReportPath)
}

if (! [IO.Path]::IsPathRooted($ReportPath)) {
    $ReportPath = [IO.Path]::GetFullPath($ReportPath)
}

if ($PsCmdlet.ParameterSetName -ne "Summarise") {

    if ($Duration -lt $Frequency) {
        Throw $CollectReplicationMetrics_LocalizedStrings.res_0006
    }

    if ($Frequency -lt [TimeSpan]::FromSeconds(1)) {
        Throw $CollectReplicationMetrics_LocalizedStrings.res_0007
    }
    
}

### Gather the data from the servers in the DAGs, unless we are in the 'ProcessOnly' mode
if ($mode -ne "ProcessOnly") {

    switch ($PsCmdlet.ParameterSetName) {
    
        ("Dag") {
            $serversToCollect = $DagsForReport |
                                Foreach { $_.Servers } |
                                Foreach { $_.Name }
        }
        
        ("ServerList") {
            $serversToCollect = $ServersForReport
        }
    }
                        
    ProcessCountersFromServers -Servers $serversToCollect -ReportPath $ReportPath -Duration $Duration -Frequency $Frequency

}

### Generate the report, unless we are in the 'CollectOnly' mode
if ($mode -ne "CollectOnly") {

    switch ($PsCmdlet.ParameterSetName) {
    
        ("Dag") {    
            Foreach ($Dag in $DagsForReport) {        
                $ServersForReport = $Dag.Servers | Foreach { $_.Name }
                GenerateReport -Servers $ServersForReport -DagName $Dag.Name -ReportPath $ReportPath -Duration $Duration -MoveFilesToArchive:$MoveFilesToArchive
            }
        }
        
        ("ServerList") {
            GenerateReport -Servers $ServersForReport -ReportPath $ReportPath -Duration $Duration -MoveFilesToArchive:$MoveFilesToArchive
        }
        
        ("Summarise") {
            GenerateReport -FileList $SummariseFiles -ReportPath $ReportPath -MoveFilesToArchive:$MoveFilesToArchive
        }
    }
    
}

# SIG # Begin signature block
# MIIabgYJKoZIhvcNAQcCoIIaXzCCGlsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUMLB2WmmCdBDk8VPAL0B9BEfj
# rKSgghUmMIIEmTCCA4GgAwIBAgITMwAAAJ0ejSeuuPPYOAABAAAAnTANBgkqhkiG
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
# mu5J4PcBZW+JC33Iacjmbuqnl84xKf8OxVtc2E0bodj6L54/LlUWa8kTo/0xggSy
# MIIErgIBATCBkDB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQQITMwAAAJ0ejSeu
# uPPYOAABAAAAnTAJBgUrDgMCGgUAoIHUMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3
# AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEW
# BBQ+1Xp8NFWd1qKpJqIZAkmr5SboSzB0BgorBgEEAYI3AgEMMWYwZKA8gDoAQwBv
# AGwAbABlAGMAdABSAGUAcABsAGkAYwBhAHQAaQBvAG4ATQBlAHQAcgBpAGMAcwAu
# AHAAcwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJ
# KoZIhvcNAQEBBQAEggEARw/YQIS4lECYVUSJbuOH2TjBrtl7gLcO8MiGTsVEhVHI
# iGA7vp604El8Blh49ISUTYIEMSajmTO/wcr7dyAqLSwvTJSsxRkpdy8BOnQHC8kW
# v1ftisLVzCMxLKQR108Iq/xjUdU6QsbuaIZ74EiWWiGCnrZZA9eOpe//PEXUZFs1
# XT3H9pcHMuYcA+K/UD9AeuJtQ3hBKnwwnDOg7ubojKTkDQR7NU2/v/mtm2TLRrjQ
# j2zYpJX4/A/SgmednVbBcx4QlsCcgqqhqChkswbB+BSGPoI72XNoLb7iaOvxsgAc
# 2ZBaWttVQblaH1IuGJ0nEpgvWGgj87B5lhPlWDShhKGCAh8wggIbBgkqhkiG9w0B
# CQYxggIMMIICCAIBATCBhTB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECCmECjkIA
# AAAAAB8wCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJ
# KoZIhvcNAQkFMQ8XDTEzMDEwODA4NDY0N1owIwYJKoZIhvcNAQkEMRYEFN7Ka4N5
# DqWgMQU5Rls4FB+qT/fAMA0GCSqGSIb3DQEBBQUABIIBAFPuQWyb9yI/+wdkSKVp
# T8/dKxNAedGhitIUYXkek01adWD5r4fKh64DNjVAP05f8lBCPSHqTte1tN5rT2vN
# RdL+9rMknfbSjblXfZLTV41Rz/OFgINDm+ge9X01flIZK9JcV6c0l084oZNh5t46
# 4NbH1h9PkHNEZD3NUADqrnNnlxYYrx95iTsNxwPJQ4+dCyZ2ruXiS5GnJ8PZPG0g
# 2AQ/TAGJt6TAhCKADIfoW64wBfNZrS+/YpEgdMfB1jmxXJ1emXM1UiKoezR8sG19
# iAxC543WbL6p16JhQ57rYo6nK9JTJ4R0iAujPpi42GWNWxjVJGFtdHzD2ElHUUgn
# Kcg=
# SIG # End signature block
