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
# MIIa0wYJKoZIhvcNAQcCoIIaxDCCGsACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUMLB2WmmCdBDk8VPAL0B9BEfj
# rKSgghWCMIIEwzCCA6ugAwIBAgITMwAAAHGzLoprgqofTgAAAAAAcTANBgkqhkiG
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
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBLswggS3
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAAAymzVMhI1xOFV
# AAEAAADKMAkGBSsOAwIaBQCggdQwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFD7V
# enw0VZ3WoqkmohkCSavlJuhLMHQGCisGAQQBgjcCAQwxZjBkoDyAOgBDAG8AbABs
# AGUAYwB0AFIAZQBwAGwAaQBjAGEAdABpAG8AbgBNAGUAdAByAGkAYwBzAC4AcABz
# ADGhJIAiaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL2V4Y2hhbmdlIDANBgkqhkiG
# 9w0BAQEFAASCAQCQY7yk9fUPthq+aNfa0RnFNlbcYYn7DsmbQGuLYO6/kSc+xcHl
# zxpvedCKuJN60iD89GTx3fydLpJx5YQm3nxdx/ywXV97T28MGIzfVBdJetlOE19I
# e02dGkLwtgG72JG/LY+zKQw6h7m31UIfR6G+516GJV5pjbmkF0vbDTdHfLzdOnA8
# +QYMe3gGyQXllC1WwG3GWiG1NhxFgmDDv4qIAyQ/UcCQT1FkZsEFN1QnPdn+eqCh
# 03gkSrOM3TdHQOlMRmptgqTc8NA0MI3YdieF+34Wwl3VsfdmfbIOUrZPt1G4/L92
# PCmnL5DD8pC2yL8xdqD4IQDc91pt+1r70entoYICKDCCAiQGCSqGSIb3DQEJBjGC
# AhUwggIRAgEBMIGOMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xITAfBgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQQITMwAAAHGzLopr
# gqofTgAAAAAAcTAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEH
# ATAcBgkqhkiG9w0BCQUxDxcNMTUwNDEwMDI1NzMxWjAjBgkqhkiG9w0BCQQxFgQU
# O8pnDqT3qhU8A4iinUviBrtOCBswDQYJKoZIhvcNAQEFBQAEggEADIDfrPs7WwFg
# Y21zmmNAbLhjdCR8IdO9NV43sfc3aX0/nwZ8EDuGQbwLuEL5y2x5UNVpvRXhVsu3
# PzPu6ybWFOZgzD7hUvNMa2H780xdtowzZtmzeQCxtksj84aFRCNc0lhaBnjYEss6
# l5m5/Yd3yRXfXV/mNARPxXgsaAdOhY1pah0wu9aaPe88LrdVjojdxSfsBz9Hf6HJ
# e7QSmfdGZOTw1FogAqv53X9WIZNZfOhajRH6043OcijKV3VuOdaZMPy9sDnaAWfm
# h4BZU/CvOcncy88TU2S69/pkSSefZjChTdT9C7QqnJOTH7SzSEnfqwfYgCFrOV9Q
# m/DH01IR9w==
# SIG # End signature block
