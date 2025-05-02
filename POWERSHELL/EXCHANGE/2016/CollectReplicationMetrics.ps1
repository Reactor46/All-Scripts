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
# MIIdugYJKoZIhvcNAQcCoIIdqzCCHacCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUMLB2WmmCdBDk8VPAL0B9BEfj
# rKSgghhkMIIEwzCCA6ugAwIBAgITMwAAAJzu/hRVqV01UAAAAAAAnDANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBMAwggS8AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCB1DAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUPtV6fDRVndaiqSaiGQJJq+Um6EswdAYKKwYB
# BAGCNwIBDDFmMGSgPIA6AEMAbwBsAGwAZQBjAHQAUgBlAHAAbABpAGMAYQB0AGkA
# bwBuAE0AZQB0AHIAaQBjAHMALgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20vZXhjaGFuZ2UgMA0GCSqGSIb3DQEBAQUABIIBAAiKruB0GltaBa8Zv6ER
# yYo4ZQodh8jrq0RqGtRzOujuNGOF6hy/OD/uGhha/XIDGODzVU3WHFrO3/Nsb0z+
# mPtqvFmEN3rlJsP3EIHVD6WF8YsHnrmXSU4j7vL4xpDSJN+kBX+fCO2KtfBR9cKc
# pgWoP9gzg9HPM7d1w1+4F2oxMh6BIluLa7gd43QpzRAMv0ljZPW29wrtAi1UDQd6
# T2YtzSO600v/Rq0x3l+kE5pdbpLnorJ6aKpKGk8CZcM3RWy/zZTPVi/LMeaw8MV4
# dkN1nROix3SG40aPsYui8aD1XTXxZmSLTqV6kzKYSYJXzsTARO2rZ4hDOvFlrwM3
# LbmhggIoMIICJAYJKoZIhvcNAQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBAhMzAAAAnO7+FFWpXTVQAAAAAACcMAkGBSsOAwIaBQCgXTAY
# BgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMx
# ODQ0MjlaMCMGCSqGSIb3DQEJBDEWBBQks1XYeDS6ZuiWcbVVOyXt2wLoUTANBgkq
# hkiG9w0BAQUFAASCAQCaRJX5IeZZah20zX6opH8wv4Kafi7aot9zAoksUYQk3KFb
# 8zQLKCzOm1X2n18oDfxK9Cql6H7t2Eyp9aZWdYlhIt//klzRQnnuhx1Jo1U3Z50f
# C3J6TBkOlP1H0QRRz8pyaiky+1UXk1aYkqxkZ0kvQCz/LEwgU2VotPx8D4sU713W
# MCgN+vofg9SlvPx8jXnhq2Gg39DEIW5uGyu2RarzSrGncSKWAhDAWJoHHTShMmNN
# 9CzVdN4i+EWY7V4fInKn02ylCGhYTbQTs5kXGae1q3LXaKPkjvLg23qPVzaFe3oR
# aYu+O3XXQlYubXMBy5ckTpcij98V0jrS/sJfFJYt
# SIG # End signature block
