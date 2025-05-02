[CmdletBinding()]
Param
(
    [parameter(Mandatory=$true,
              ParameterSetName="Server",
              HelpMessage="Server where trace files exist on well known share")]
    [String]
    $Server,

    [parameter(Mandatory=$true,
              ParameterSetName="Database",
              HelpMessage="Server(s) hosting database copy where trace files exist on well known share")]
    [String]
    $DatabaseName,

    [parameter(Mandatory=$true,
              ParameterSetName="Mailbox",
              HelpMessage="Server where trace files exist on well known share")]
    [String]
    $MailboxId,

    [parameter(ParameterSetName="Mailbox",
              HelpMessage="Organization Identity for mailbox (datacenter only)")]
    [String]
    $OrganizationId,

    [parameter(ParameterSetName="Mailbox",
              HelpMessage="Archive Mailbox Switch")]
    [Switch]
    $Archive,

    [parameter(Mandatory=$true,
              ParameterSetName="FolderPath",
              HelpMessage="Full file path to folder containing ETL trace files")]
    [String]
    $TraceFolderPath,

    [parameter(ParameterSetName="FolderPath", HelpMessage="Use transformer modules from a local enlistment.")]
    [switch]
    $UseLocalTransformerModules,

    [parameter(HelpMessage="Path to trace transformer modules (default is '\\redmond\exchange\Build\E16\Latest\sources')")]
    [string]
    $TransformerModulePath,

    [parameter(Mandatory=$true,
              ParameterSetName="FilePath",
              HelpMessage="Full file path to single ETL trace file")]
    [String]
    $TraceFilepath,

    [parameter(HelpMessage="Retrieve trace files from all database copies. By default, only retrieve traces from active database copy.")]
    [Switch]
    $UseAllDatabaseCopies,

    [parameter(HelpMessage="Passive server hosting database copy where trace files exist on well known share")]
    [String]
    $CopyOnServer,

    [parameter(HelpMessage="Use well known share in datacenter topology, used with Server and Mailbox parameter sets ")]
    [Switch]
    $Datacenter,
    
    [parameter(HelpMessage="Do not require Organization with Mailbox parameter sets")]
    [Switch]
    $Dedicated,

    [parameter(HelpMessage="Retrieves trace events greater than or equal to Start timestamp")]
    [DateTime]
    $Start=(get-date).AddDays(-1),
    
    [parameter(HelpMessage="Retrieves trace events less than or equal to End timestamp")]
    [DateTime]
    $End=(get-date).AddHours(1),

    [parameter(HelpMessage="StartCreationTime used in get-MachineLog (default 48 hrs before start timestamp)")]
    [DateTime]
    $StartCreationTime,
    
    [parameter(HelpMessage="Folder path to working directory to convert ETL trace file (default is 'c:\temp\LongOperation')")]
    [String]
    $WorkingPath="$Home\AppData\Local\Temp\StoreTrace",

    [parameter(HelpMessage="Allow script to proceed without Reference Trace Files")]
    [Switch]
    $AllowZeroReferenceFiles,

    [parameter(HelpMessage="Trace type to retrieve. Can be BreadCrumbs, FullTextIndexQuery, HeavyClientActivity, LockContention, LongOperation, RopResource, InstantSearchDocumentId, RopSummary, MailboxInfo, OperationDetail, OperationParameterColumns, OperationParameterSort, OperationParameterRestriction, OperationParameterOther, OperationContext, IntegrityCheckStatus, IntegrityCheckCorruption, DatabaseAggregateRops, MailboxAggregateRops or InstantSearchBigFunnel")]
    [String]
    [ValidateSet("BreadCrumbs", "FullTextIndexQuery", "HeavyClientActivity", "LockContention", "LongOperation", "RopResource", "InstantSearchDocumentId", "RopSummary", "MailboxInfo", "OperationDetail", "OperationParameterColumns", "OperationParameterSort", "OperationParameterRestriction", "OperationParameterOther", "OperationContext", "IntegrityCheckStatus", "IntegrityCheckCorruption","DatabaseAggregateRops", "MailboxAggregateRops", "InstantSearchBigFunnel")]
    $TraceType="RopSummary",

    [parameter(HelpMessage="Include detail if TraceType is LongOperation.")]
    [Switch]
    $IncludeDetailTrace,

    [parameter(HelpMessage="Exclude LongOperation TraceType when not resource intense")]
    [Switch]
    $ExcludeNonResourceIntensiveOperations
)

Function CleanWorkingPath ()
{
    If (test-path $WorkingPath)
    {
        $LogFiles = @(Get-ChildItem "$WorkingPath\*.*")
        If (test-path "$WorkingPath\ReferenceData")
        {
            $RefLogFiles = @(Get-ChildItem "$WorkingPath\ReferenceData\*.*")
            If ($RefLogFiles.count -gt 0) {$LogFiles += $RefLogFiles}
        }

        ForEach ($LogFile in $LogFiles) 
        {
            write-debug ("$(get-date) Removing $($LogFile.Name)")
            Remove-item $LogFile
        }
    }
}

Function CopyTransformerModules
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [string[]]$Modules,

        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$DestinationPath
    )

    write-host "$(get-date) Copying transformer modules to $DestinationPath"
    If (test-path $SourcePath -ErrorAction SilentlyContinue)
    {
        If (!(test-Path $DestinationPath -ErrorAction SilentlyContinue)) {[void](New-Item -ItemType Directory $DestinationPath)}

        $Files = @(Get-ChildItem -Path $SourcePath)

        ForEach ($File in $Files)
        {
            copy $File.Fullname $DestinationPath -ErrorAction SilentlyContinue
        }
    }
}

Function LoadTransformerModules
{

    Param
    (
        [Parameter(Mandatory=$true)]
        [string[]]$Modules,

        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    write-host "$(get-date) Loading transformer modules from $path"

    If (!$Server -or ($Server -and !$Datacenter))
    {
        [string]$distribDebugPath = join-Path -Path $path -ChildPath "Distrib\PRIVATE\BIN\DEBUG\AMD64\"
        [string]$distribRetailPath = join-Path -Path $path -ChildPath "Distrib\PRIVATE\BIN\RETAIL\AMD64\"
    }
    Else
    {
        [string]$distribDebugPath = join-Path -Path $path -ChildPath "Distrib\PRIVATE\BIN\DEBUG\AMD64\"
        [string]$distribRetailPath = join-Path -Path $workingPath -ChildPath "bin"
        CopyTransformerModules -Modules:$Modules -SourcePath $path -DestinationPath $distribRetailPath
    }

    $LoadedModules = @()

    if (test-Path -Path $distribDebugPath -PathType Container)
    {
        [string]$distribPath = $distribDebugPath
        write-debug ("Using debug/amd64")
    }
    elseif (test-Path -Path $distribRetailPath -PathType Container)
    {
        [string]$distribPath = $distribRetailPath
        write-debug ("Using retail/amd64")
    }
    else
    {
        throw "distrib path not found for either debug/retail amd64"
    }

    foreach ($module in $modules)
    {
        [string]$modulePath = join-Path -Path $distribPath -ChildPath $module

        if (test-Path -Path $modulePath -PathType Leaf)
        {
            [IO.FileInfo]$info = get-Item $modulePath
            [string]$version = $info.VersionInfo.ProductVersion
            write-Host "$(get-date) Loading $version $module"
            import-Module $modulePath
        }
        else
        {
            throw "assembly $module not found in $distribPath"
        }
    }

    # By default, data mining code will try to log events which may fail if running with non-administrator permissions
    # set DebugLevel to 0 to avoid writing any debug logs.  (The setter is private so this uses reflection.)
    [Microsoft.Exchange.Management.DataMining.SharedConfiguration].GetProperty("DebugLevel").SetMethod.Invoke($null, 0)
}

Function LoadReferenceData ()
{
    param ($FilePath, $DataType, $HashTable)

    Switch ($DataType.ToUpper())
    {
        DATABASEINFO
        {
            import-csv -path $FilePath | % { If (!$HashTable.Contains($_.DatabaseHash)) {$HashTable[$_.DatabaseHash] = $_.DatabaseName} }
            break;
        }

        OPERATIONDETAIL
        {
            import-csv -path $FilePath | % { If (!$HashTable.Contains($_.key)) {$HashTable[$_.key] = $_} }
            break;
        }        
        
        Default
        {
            import-csv -path $FilePath | % { If (!$HashTable.Contains($_.key)) {$HashTable[$_.key] = $_.Value} }
            break;
        }
    }

    Return $HashTable
}

Function LoadMailboxInfo ()
{
    param ($FilePath, $HashTable)

    import-csv -path $FilePath | % { `
        $key = "$($_.DatabaseHash),$($_.MailboxNumber)"
        If (!$HashTable.Contains($key))
        {
            $HashTable[$key] = [guid]$_.MailboxGuid
        }
        
    }

    Return $HashTable
}

Function GetSourceFolder()
{
    param ($ExServer, $Datacenter)

    $server = $ExServer.FQDN
    If ($Datacenter)
    {
        if (Check-ServerNewLogLocation($ExServer))
        {
            $checkPath = "\\$server\LOGS_D\store"
        }
        else
        {
            $checkPath = "\\$server\Exchange\logging\store"
        }

        If (Test-Path $checkPath)
        {
            $SourceFolder = $checkPath
        }
        Else
        {
            write-error "Could not find logging folder $checkPath share on $server!"            
        }
    }
    Else
    {
        If (Test-Path "\\$server\logging\store")
        {
            $SourceFolder = "\\$server\logging\store"
        }
        ElseIf (Test-Path "\\$server\exchsrvr\logging\store")
        {
            $SourceFolder = "\\$server\exchsrvr\logging\store"
        }
        Else
        {
            write-error "Could not find logging folder share on $server!"
        }
    }

    return $SourceFolder   
}

Function GetMailboxNumber ()
{
    Param ($DatabaseName, $MailboxGuid)
    If (test-path $ExScripts\ManagedStoreDiagnosticFunctions.ps1)
    {
        . $ExScripts\ManagedStoreDiagnosticFunctions.ps1
        $MailboxNumber = (get-storeQuery -database $DatabaseName -query "Select MailboxNumber from Mailbox where MailboxGuid = '$MailboxGuid'").MailboxNumber
        Return $MailboxNumber
    }
    ElseIf (test-path \\redmond\exchange\Build\E16\LATEST\sources\sources\dev\Management\src\CentralAdmin\Scripts\ManagedStore\ManagedStoreDiagnosticFunctions.ps1)
    {
        . \\redmond\exchange\Build\E16\LATEST\sources\sources\dev\Management\src\CentralAdmin\Scripts\ManagedStore\ManagedStoreDiagnosticFunctions.ps1
        $MailboxNumber = (get-storeQuery -database $DatabaseName -query "Select MailboxNumber from Mailbox where MailboxGuid = '$MailboxGuid'").MailboxNumber
        Return $MailboxNumber
    }
    Else
    {
        return $NULL
    }
}

function GetFlagsEnumString
{
    param(
        [Parameter(Mandatory=$true)]
        [Hashtable]
        $ReferenceData,

        [Parameter(Mandatory=$true)]
        [ValidatePattern("[0-9]*")]
        [string]
        $Value
    )

    $intValue = [Int]$Value
    if ($intValue -eq 0)
    {
        $stringValue = $ReferenceData[$intValue]
    }
    else
    {
        $knownFlags = 0
        $stringValue = `
            @($ReferenceData.Keys | %{ $_ -band $intValue} | ?{$_ -ne 0} | %{$knownFlags = $_ -bor $knownFlags; $ReferenceData[$_]}) + `
            @(if ($knownFlags -ne $intValue) { "0x{0:x}" -f ($knownFlags -bxor $intValue) }) `
            -join ", "
    }

    return $stringValue
}

Function Get-TraceDetails ()
{
    Param ([System.Array]$TraceFileContent, $Operation, $Filename)
    $CountTrace=1
    $FirstTraceTimestamp = get-date
    $LastTraceTimestamp = [DateTime]::MinValue
    ForEach ($Line in $TraceFileContent)
    {
        If ($Line -match 'Correlation Id')
        {
            $DetailTrace+="`n$line"
            $id = $line.split(":")[1].trim();`
            if ($id -eq $Operation.CorrelationId) 
            {
                $CollectTrace=$true
            } 
        }
        ElseIf ($Line -eq "")
        {
            If ($DetailTrace.length -gt 0 -and $CollectTrace)
            {
                #write-host "$(get-date) Found operation CorrelationId $($operation.CorrelationId) detail in $Script:CurrentTraceFileContent"
                $Operation.Trace = $DetailTrace
                Return $operation
                #Write-Output $DetailTrace;`
            }
            $DetailTrace=""
            $DBPlanSection=$false
            $CollectTrace=$false
        } 
        ElseIf ($Line -notmatch 'Sequence,TimeStamp,Data')
        {
            if ($DetailTrace -eq "" -and $Line -match ",")
            {
                $TraceTimestamp = get-date("$($Line.split(",")[1])z")
                If ($TraceTimestamp -lt $FirstTraceTimestamp)
                {
                    $FirstTraceTimestamp = $TraceTimestamp
                }
                If ($TraceTimestamp -gt $LastTraceTimestamp)
                {
                    $LastTraceTimestamp = $TraceTimestamp
                }
                If ($TraceTimestamp -gt $Operation.Timestamp)
                {
                    #Trace detail not in this file, return NULL and continue to next file
                    #write-host "exiting because detail trace timestamp ($TraceTimestamp) > operation timestamp ($($operation.Timestamp))"
                    return $NULL
                }
            }
            if ($line -match 'Executed DB plans')
            {
                $DetailTrace+="`n$line"
                $DBPlanSection=$true
            }
            Elseif ((!$Line.StartsWith("`t") -and $DBPlanSection) -or $DetailTrace -eq "")
            {
                $DetailTrace+="$($line.replace("`r",''))";`
            }
            Else
            {
                $DetailTrace+="`n$line";`
            }
        }
        Else
        {
            $CollectTrace=$false;`
        }
    }
    #write-host "operation ($($Operation.timestamp)) not found in current file (last entry '$TraceTimestamp')"
    If (!$fileTimeSpan.contains($Filename)) 
    {        
        $fileTimeSpan[$Filename]=@($FirstTraceTimestamp,$LastTraceTimestamp)
    }
}

Function Get-TraceFromFiles ()
{
    param($operation, [System.Array]$files, $TraceType)
    $Result=$NULL
    Add-Member -InputObject $operation -MemberType NoteProperty -Name Trace -Value $NULL -ErrorAction SilentlyContinue

    If ($TraceType -eq 'LongOperation' -and $Operation.BuildNumber)
    {
        $BuildNumber=$operation.BuildNumber.split(".")
        If ([int]$BuildNumber[1] -eq 0 -and [int]$BuildNumber[2] -lt 392)
        {
            $Version=1
        }
        If ([int]$BuildNumber[1] -eq 0 -and [int]$BuildNumber[2] -lt 954)
        {
            $Version=2
        }
        Else
        {
            $Version=3
        }
    }
    ElseIf ($TraceType -eq 'LongOperation')
    {
        $Version=1
    }
    Else
    {
        $Version=3
    }

    If ($Version -eq 3)
    {
        Add-Member -InputObject $operation -MemberType NoteProperty -Name TotalChunks -Value 0 -ErrorAction SilentlyContinue
        Add-Member -InputObject $operation -MemberType NoteProperty -Name ChunksAdded -Value 0 -ErrorAction SilentlyContinue
    }
    ForEach ($file in $files)
    {
        If ($Version -le 2)
        {
            If ($script:fileTimeSpan.Contains($file.name))
            {
                $Timespan = $script:fileTimeSpan[$file.name]
                If ($operation.timestamp -ge $Timespan[0] -and $operation.timestamp -le $Timespan[1])
                {
                    If ($file.name -ne $Script:CurrentTraceFileContent) 
                    {
                        #write-host "CurrentTraceFileContent is $CurrentTraceFileContent, reading $($file.name)"
                        $Script:TraceFileContent = Get-Content -Path $file.FullName
                        $Script:CurrentTraceFileContent = $file.name
                    }
                    $Result = Get-TraceDetails -TraceFileContent $Script:TraceFileContent -Operation $operation -FileName $File.Name
                }
                Else
                {
                    #write-host "operation CorrelationId $($operation.CorrelationId) does not exist in '$($file.FullName)' (outside timespan)"
                }
            }
            Else
            {
                If ($file.name -ne $Script:CurrentTraceFileContent) 
                {
                    #write-host "CurrentTraceFileContent is $CurrentTraceFileContent, reading $($file.name)"
                    $Script:TraceFileContent = Get-Content -Path $file.FullName
                    $Script:CurrentTraceFileContent = $file.name
                }
                $Result = Get-TraceDetails -TraceFileContent $Script:TraceFileContent -Operation $operation -FileName $File.Name
            }
        }
        Else
        {
            $DetailRecords = @(Import-csv $file.Fullname | ? {$_.CorrelationId -eq $operation.CorrelationId} | ? {$_.TotalChunks -gt 0} | Sort ChunkIndex)
            If ($DetailRecords.count -gt 0)
            {
                ForEach ($Chunk in $DetailRecords)
                {
                    $Operation.Trace += $Chunk.Content.Replace("``t", "`t").Replace("``r", "`r").Replace("``n", "`n")
                    $Operation.TotalChunks = $Chunk.TotalChunks
                    $Operation.ChunksAdded++
                }
                $Result = $operation
            }         
        }
        if ($Result) 
        {
            $Script:DetailTraceFoundCount++
            return $result
            Break
        }
    }
    if (!$Result) 
    {
        $Script:DetailTraceNotFoundCount++
        write-debug ("$(get-date) Operation with CorrelationId $($operation.CorrelationId) not found in detail traces")
        $operation.trace=$NULL
        return $operation
    }
}

function CopyEtlFiles
{
    [CmdletBinding(DefaultParameterSetName="Default")]
    param(
	[Parameter(ParameterSetName="Default")]
        [Parameter(ParameterSetName="UseTorus")]
        [Parameter(Mandatory=$true)]
        [System.Array]
        $EtlFiles,

	[Parameter(ParameterSetName="Default")]
        [Parameter(ParameterSetName="UseTorus")]
        [Parameter(Mandatory=$true)]
        [string]
        $WorkingPath,

        [Parameter(ParameterSetName="UseTorus", Mandatory=$true)]
        [switch]
        $UseTorus,

        [Parameter(ParameterSetName="UseTorus", Mandatory=$true)]
        #[Microsoft.Exchange.Data.Directory.Management.ADPresentationObject]
        $Server)

    write-host "$(get-date) Copying $($EtlFiles.count) trace files to '$WorkingPath'"
    if ($UseTorus)
    {
        $downloadList = [String]::Join(",", @($EtlFiles | %{$_.Name}))

        if (Check-ServerNewLogLocation($Server))
        {
            Get-MachineLog -Target $Server.Name -Log StoreDatacenter -Filter $downloadList -DownloadToLocalFolderPath $WorkingPath | Out-Null
        }
        else
        {
            Get-MachineLog -Target $Server.Name -Log Store -Filter $downloadList -DownloadToLocalFolderPath $WorkingPath | Out-Null
        }
    }
    else
    {
        ForEach ($EtlFile in $Etlfiles)
        {
            # Copy ETL files to working directory
            write-debug ("$(get-date) Working with $($EtlFile.FullName) $($EtlFile.LastWriteTime)")
            $LocalEtlFileName = "$WorkingPath\$($EtlFile.Name)"
            write-debug ("$(get-date) Copying $($EtlFile.Name) to $LocalEtlFileName")
            copy $EtlFile $LocalEtlFileName
        }
    }
}

function TransformEtlFiles
{
    param([string]$WorkingPath)

    [Microsoft.Exchange.Management.DataMining.UploaderConfiguration]$ulc = new-object Microsoft.Exchange.Management.DataMining.UploaderConfiguration
    [Microsoft.Exchange.Management.DataMining.UploaderLog]$ull = new-object Microsoft.Exchange.Management.DataMining.UploaderLog
    [string]$transformPath = join-Path -Path ($WorkingPath) -ChildPath "*.etl"

    $ull.SourcePath = $WorkingPath
    $ull.TransformationPath = $WorkingPath
    $ull.LogFileSelection = "*.etl"

    #write-host "UploaderLog SourcePath =" $ull.SourcePath
    #write-host "UploaderLog TransformationPath =" $ull.TransformationPath
    #write-Host "UploaderLog LogFileSelection =" $ull.LogFileSelection

    [Microsoft.Exchange.Management.DataMining.EventTracingTransformer]$ett = New-Object Microsoft.Exchange.Management.DataMining.EventTracingTransformer -ArgumentList $ulc,$ull

    #write-Host "Removing transformer bookmark"
    remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Datamining\Uploader" -Name "EventTracingTransformerBookmark" -ErrorAction SilentlyContinue

    write-Host "$(get-date) Transforming $transformPath"

    $ett.TransformLogs()
    $ett.Cleanup()

}

function Check-ServerNewLogLocation
{
    param(
        [Parameter(Mandatory=$true)]
        $ExServer
    )

    # defaults to version to check so if we go above 15.1 and/or 
    # the format is not matching in the future, we default to new location
    $Build = "15.01.0372.000"
    $BuildNewLocation = "15.01.0372.000"

    If ($ExServer.AdminDisplayVersion.Tostring().contains("Version 15.0 (Build"))
    {
        $ExMajorVersion = [Int]$ExServer.AdminDisplayVersion.Tostring().replace("Version 15.0 (Build","").replace(")","").trim().Split(".")[0]
        $ExMinorVersion = [Int]$ExServer.AdminDisplayVersion.Tostring().replace("Version 15.0 (Build","").replace(")","").trim().Split(".")[1]
        $Build = [String]::Format("15.00.{0:0000}.{1:000}", $ExMajorVersion, $ExMinorVersion) #15.00.0651.006
    }
    If ($ExServer.AdminDisplayVersion.Tostring().contains("Version 15.1 (Build"))
    {
        $ExMajorVersion = [Int]$ExServer.AdminDisplayVersion.Tostring().replace("Version 15.1 (Build","").replace(")","").trim().Split(".")[0]
        $ExMinorVersion = [Int]$ExServer.AdminDisplayVersion.Tostring().replace("Version 15.1 (Build","").replace(")","").trim().Split(".")[1]
        $Build = [String]::Format("15.01.{0:0000}.{1:000}", $ExMajorVersion, $ExMinorVersion) #15.01.0651.006
    }

    $Build -ge $BuildNewLocation
}

function Get-TraceFileList
{
    param(
        [Parameter(Mandatory=$true)]
        $ExServer,

        [Parameter(Mandatory=$true)]
        [bool]
        $Datacenter,

        [Parameter(Mandatory=$true)]
        [bool]
        $UseTorus
    )

    if ($UseTorus)
    {
        #filter filelist returned by Get-MachineLog to reduce latency (retention period is accumulating more files)
        if (!$StartCreationTime) {$StartCreationTime = get-date($script:start).AddDays(-2).ToUniversalTime()}
        write-host "$(get-date) Looking for trace files on $($ExServer.name) ($($ExServer.AdminDisplayVersion)) using Get-MachineLog (StartCreationTime: $StartCreationTime UTC)"
        if (Check-ServerNewLogLocation($ExServer))
        {
            @(Get-MachineLog -Target $ExServer.Name -Log StoreDatacenter -StartCreationTime $StartCreationTime )
        }
        else
        {
            @(Get-MachineLog -Target $ExServer.Name -Log Store -StartCreationTime $StartCreationTime )
        }
    }
    else
    {
        $SourceFolder = GetSourceFolder -Server $ExServer -Datacenter $Datacenter
        write-host "$(get-date) Looking for trace files at '$SourceFolder'"
        @(Get-ChildItem "$($SourceFolder)\*.etl")
    }
}

#If ($Debug) {$DebugPreference = "SilentlyContinue"}

$StartPath = $pwd
$SourceFolder = $NULL
$EtlFiles=@()
$RefEtlFiles=@()
$MinRefFiles=4

If ($MyInvocation.BoundParameters.ContainsKey("IncludeDetailTrace") -and $TraceType -notin ("FullTextIndexQuery", "HeavyClientActivity", "LongOperation")) 
{
    write-error "IncludeDetailTrace parameter only supported with FullTextIndexQuery, HeavyClientActivity and LongOperation TraceTypes"
    return
}

If ($MyInvocation.BoundParameters.ContainsKey("ExcludeNonResourceIntensiveOperations") -and $TraceType -ne "LongOperation") 
{
    write-error "ExcludeNonResourceIntensiveOperations parameter only supported with LongOperation TraceType"
    return
}

$useTorus = $false
$torus = Get-Variable TorusConnection -ErrorAction SilentlyContinue
if ($torus -ne $null)
{
    if ($torus.Value.CapacityOnly.ToBool() -or -not $torus.Value.ConnectCapacity)
    {
        Write-Error "Get-StoreTrace.ps1 only works with Torus in management-capacity dual sessions."
        return
    }
    else
    {
        $useTorus = $true
    }
}

$TraceTypeFileName = $TraceType

switch($TraceType)
{
    "RopSummary"
    {
        $TraceTypeUsesReferenceData = $true
        break
    }
    
    "RopResource"
    {
        $TraceTypeUsesReferenceData = $true
        break
    }
    
    "LockContention"
    {
        $TraceTypeUsesReferenceData = $true
        break
    }

    "HeavyClientActivity"
    {
        $TraceTypeUsesReferenceData = $false
        break
    }

    "BreadCrumbs"
    {
        $TraceTypeUsesReferenceData = $true
        break
    }

    "FullTextIndexQuery"
    {
        $TraceTypeUsesReferenceData = $true

        # We have two types of traces with the same name but different formats
        # this is used to specify which one we want.
        # This can be FullTextIndexQuery or FullTextIndexSingleLine
        # TODO - This really should be a parameter to the script...we probably want
        # parameter sets for each query type, but I will leave that for a later change
        $FullTextIndexTraceType = "FullTextIndexSingleLine" 
        break
    }

    "LongOperation"
    {
        $TraceTypeUsesReferenceData = $false
        break
    }

    "InstantSearchDocumentId"
    {
        $TraceTypeUsesReferenceData = $false
        $TraceTypeFileName = "FullTextIndexQuery"
        break
    }

    "MailboxInfo"
    {
        $TraceTypeUsesReferenceData = $true
        $TraceTypeFileName = "ReferenceData"
    }

    "OperationDetail"
    {
        $TraceTypeUsesReferenceData = $true
        $TraceTypeFileName = "ReferenceData"
    }

    "OperationParameterColumns"
    {
        $TraceTypeUsesReferenceData = $true
        $TraceTypeFileName = "ReferenceData"
    }

    "OperationParameterSort"
    {
        $TraceTypeUsesReferenceData = $true
        $TraceTypeFileName = "ReferenceData"
    }

    "OperationParameterRestriction"
    {
        $TraceTypeUsesReferenceData = $true
        $TraceTypeFileName = "ReferenceData"
    }

    "OperationParameterOther"
    {
        $TraceTypeUsesReferenceData = $true
        $TraceTypeFileName = "ReferenceData"
    }

    "OperationContext"
    {
        $TraceTypeUsesReferenceData = $true
        $TraceTypeFileName = "ReferenceData"
    }    

    "IntegrityCheckStatus"
    {
        $TraceTypeUsesReferenceData = $false
        $TraceTypeFileName = "IntegrityCheck"
    }

    "IntegrityCheckCorruption"
    {
        $TraceTypeUsesReferenceData = $false
        $TraceTypeFileName = "IntegrityCheck"
    }

    "DatabaseAggregateRops"
    {
        $TraceTypeUsesReferenceData = $true
        $TraceTypeFileName = "RopSummary"
    }

    "MailboxAggregateRops"
    {
        $TraceTypeUsesReferenceData = $true
        $TraceTypeFileName = "RopSummary"
    }

    "InstantSearchBigFunnel"
    {
        $TraceTypeUsesReferenceData = $false
        $TraceTypeFileName = "FullTextIndexQuery"
        break
    }

    default
    {
        throw "Unexpected TraceType"
    }
}

CleanWorkingPath

#Create WorkingPath if it doesn't exist
If (!(Test-Path -Path $WorkingPath -ErrorAction SilentlyContinue)) {[void](New-Item -ItemType Directory $WorkingPath)}
If (!(Test-Path -Path "$WorkingPath\ReferenceData" -ErrorAction SilentlyContinue)) {[void](New-Item -ItemType Directory "$WorkingPath\ReferenceData")}

If ($TraceFolderPath -ne $null -and $TraceFolderPath.Length -gt 0)
{
    write-debug ("TraceFolderPath specified")
    $SourceFolder = $TraceFolderPath

    if ($UseLocalTransformerModules)
    {
        [string]$TransformerModulePath = $env:SDROOT
    }

    write-host "$(get-date) Looking for $TraceType trace files at '$SourceFolder'"
    $TempEtlFiles = @(Get-ChildItem "$($SourceFolder)\$TraceTypeFileName*.etl"  | ? {($_.LastWriteTime -ge $Start -and $_.LastWriteTime -le $End) -or ($_.CreationTime -le $End -and $_.LastWriteTime -ge $End)} | Sort CreationTime)

    If ($TempEtlFiles.count -eq 0)
    {
        $EtlFiles += @(Get-ChildItem "$($SourceFolder)\$TraceTypeFileName*.etl" | sort CreationTime | Select -Last 1)
    }
    Else
    {
        $EtlFiles += $TempEtlFiles
    }

    if ($TraceTypeUsesReferenceData)
    {
        write-host "$(get-date) Looking for ReferenceData trace files at '$SourceFolder'"
        
        $TempRefEtlFiles = @(Get-ChildItem "$($SourceFolder)\ReferenceData*.etl" | sort CreationTime | % { `
            Add-Member -InputObject $_ -MemberType NoteProperty -Name TraceFileGroup -Value ($_.name.replace("ReferenceData_","").split("_")[0]) -PassThru;`
                })
        #Select first reference file containing many of the reference traces logged at service startup plus last four (containing MailboxInfo and Activity)
        $RefEtlFiles += $TempRefEtlFiles | Group TraceFileGroup | Sort Name | Select -last 1 -ExpandProperty Group | Select -first 1 -last $MinRefFiles
        If ($RefEtlFiles.count -lt $MinRefFiles)
        {
            write-host "Latest ReferenceData trace file series only contains $($RefEtlFiles.count) files, downloading last $MinRefFiles files by CreationTime"
            $RefEtlFiles = $TempRefEtlFiles | Select -last $MinRefFiles  
        }
    }

}
ElseIf ($TraceFilepath)
{
    write-debug ("TraceFilepath specified")
    If (test-Path $TraceFilepath)
    {
        $EtlFiles = @(get-item $TraceFilepath)
        write-host ("$(get-date) Found file at '$TraceFilepath'")
    }
    Else
    {
        write-error "No LongOperation trace files found at '$TraceFilepath'"
        return
    }
}
ElseIf ($Server -or $MailboxId -or $DatabaseName)
{
    If ($MailboxId)
    {
        If ($Datacenter -and (!$OrganizationId -and !$Dedicated))
        {
            write-error "Must provide `$OrganizationId parameter with `$Datacenter and `$Mailbox parameters"
            return
        }
        ElseIf ($Datacenter -and $OrganizationId)
        {
            $FullMailboxId = "$($OrganizationId)\$($MailboxId)"
        }
        Else
        {
            $FullMailboxId = $MailboxId
        }

        If ($OrganizationId -in ("outlook.com","hotmail.com","live.com","msn.com","84df9e7f-e9f6-40af-b435-aaaaaaaaaaaa"))
        {
            $Mailbox = get-ConsumerMailbox $MailboxId -MservDataOnly
        }

        If (!$Mailbox -and $Archive)
        {
            #check for archive mailbox
            $Mailbox = get-Mailbox $FullMailboxId -Archive -ErrorAction SilentlyContinue
        }
        
        If (!$Mailbox)
        {
            #check for primary mailbox
            $Mailbox = get-Mailbox $FullMailboxId -ErrorAction SilentlyContinue
            If (!$Mailbox)
            {
                #check for monitoring mailbox
                $Mailbox = get-Mailbox $FullMailboxId -Monitoring -ErrorAction SilentlyContinue
                If (!$Mailbox)
                {
                    #check for arbitration mailbox
                    $Mailbox = get-Mailbox $FullMailboxId -Arbitration -ErrorAction SilentlyContinue
                    If (!$Mailbox)
                    {
                        #check for public folder mailbox
                        $Mailbox = get-Mailbox $FullMailboxId -PublicFolder -ErrorAction SilentlyContinue
                        If (!$Mailbox)
                        {
                            #check for site mailbox
                            $Mailbox = get-SiteMailbox $FullMailboxId -ErrorAction SilentlyContinue
                            If (!$Mailbox)
                            {
                                write-error "Mailbox not found for '$MailboxId', exiting"
                                return
                            }
                        }
                    }
                }
            }
        }
        If ($Mailbox)
        {
            If ($Archive)
            {
                $MailboxGuid = $Mailbox.ArchiveGuid
                $database = get-MailboxDatabase -status "$($mailbox.ArchiveDatabase.ToString())"
            }
            Else
            {
                $MailboxGuid = $Mailbox.ExchangeGuid
                $database = get-MailboxDatabase -status "$($mailbox.Database.ToString())"
            }
        }

    }
    ElseIf ($DatabaseName)
    {
        $database = get-MailboxDatabase -status $DatabaseName       
    }

    If ($server)
    {
        $servers = @(Get-ExchangeServer $server)
        $database=$NULL
        $MailboxGuid = $NULL
    }
    ElseIf ($database)
    {
        If ($UseAllDatabaseCopies)
        {
            $servers = @($database.servers | get-ExchangeServer)
        }
        If ($CopyOnServer)
        {
            $servers = @(Get-ExchangeServer $CopyOnServer)
        }
        Else
        {
            $server = $database.MountedOnServer.split(".")[0]
            $servers = @(Get-ExchangeServer $server)
        }
    }
    Else
    {
        Write-error "Unable to find server, database, or mailbox"
        return
    }

    Foreach ($ExServer in $servers)
    {
        If ($ExServer.AdminDisplayVersion.Tostring().contains("Version 15.0 (Build"))
        {
            $ExMajorVersion = [Int]$ExServer.AdminDisplayVersion.Tostring().replace("Version 15.0 (Build","").replace(")","").trim().Split(".")[0]
            $ExMinorVersion = [Int]$ExServer.AdminDisplayVersion.Tostring().replace("Version 15.0 (Build","").replace(")","").trim().Split(".")[1]
            $Build = [String]::Format("15.00.{0:0000}.{1:000}", $ExMajorVersion, $ExMinorVersion) #15.00.0651.006
        }
        If ($ExServer.AdminDisplayVersion.Tostring().contains("Version 15.1 (Build"))
        {
            $ExMajorVersion = [Int]$ExServer.AdminDisplayVersion.Tostring().replace("Version 15.1 (Build","").replace(")","").trim().Split(".")[0]
            $ExMinorVersion = [Int]$ExServer.AdminDisplayVersion.Tostring().replace("Version 15.1 (Build","").replace(")","").trim().Split(".")[1]
            $Build = [String]::Format("15.01.{0:0000}.{1:000}", $ExMajorVersion, $ExMinorVersion) #15.01.0651.006
        }
        write-debug ("$(get-date) $($ExServer.FQDN) Version = $ExMajorVersion.$ExMinorVersion")

        if (!$TransformerModulePath)
        {
            If ($useTorus)
            {
                [string]$TransformerModulePath = "\\redmond\exchange\Build\E16\$Build\sources"
            }
            ElseIf ($Datacenter)
            {
                [string]$TransformerModulePath = "\\$($ExServer.FQDN)\exchange\datacenter\DataMining\Cosmos"
            }
            Else
            {
                [string]$TransformerModulePath = "\\redmond\exchange\Build\E16\$Build\sources"
            }   
        }
        Else
        {
            If ($TransformerModulePath -eq "latest")
            {
                $TransformerModulePath = "\\redmond\exchange\Build\E16\LATEST\sources"
            }
        }

        $fileList = Get-TraceFileList -ExServer $ExServer -Datacenter $Datacenter -UseTorus $useTorus
        $FileListSizeMb = ($filelist | measure Length -sum).sum/1mb
        write-host "$(get-date) Found $($fileList.count) trace files on $ExServer ($FileListSizeMb MB)"

        $TempEtlFiles = @($fileList | ? {$_.Name -match "$TraceTypeFileName.*\.etl" -and (($_.LastWriteTime -ge $Start -and $_.LastWriteTime -le $End) -or ($_.CreationTime -le $End -and $_.LastWriteTime -ge $End))} | Sort CreationTime)

        If ($TempEtlFiles.count -eq 0)
        {
            $EtlFiles += @($fileList | ?{$_.Name -match "$TraceTypeFileName.*\.etl"} | sort CreationTime | Select -Last 1)
        }
        Else
        {
            $EtlFiles += $TempEtlFiles
        }

        #Retrieve Reference traces
        if ($TraceTypeUsesReferenceData)
        {        
            $TempRefEtlFiles = @($fileList | ?{$_.Name -match "ReferenceData.*\.etl"} | sort CreationTime | % { `
                Add-Member -InputObject $_ -MemberType NoteProperty -Name TraceFileGroup -Value ($_.name.replace("ReferenceData_","").split("_")[0]) -PassThru;`
                })
            #Select first reference file containing many of the reference traces logged at service startup plus last four (containing MailboxInfo and Activity)
            $LocalRefEtlFiles = $TempRefEtlFiles | Group TraceFileGroup | Sort Name | Select -last 1 -ExpandProperty Group | Select -first 1 -last 4
            If ($LocalRefEtlFiles.count -lt $MinRefFiles)
            {
                write-host "$(get-date) Latest ReferenceData trace file series only contains $($LocalRefEtlFiles.count) file(s), downloading last $MinRefFiles files by CreationTime"
                $LocalRefEtlFiles = $TempRefEtlFiles | Select -last $MinRefFiles                 
            }
            $RefEtlFiles += $LocalRefEtlFiles
        }
    }
}

If ($EtlFiles.count -eq 0)
{
    write-error "No $TraceType trace files found"
    return
}
ElseIf ($TraceTypeUsesReferenceData -and $RefEtlFiles.count -lt 2 -and !$AllowZeroReferenceFiles)
{
    write-error "Insufficient ReferenceData trace files found"
    return
}
Else
{
    write-host "$(get-date) Found $($EtlFiles.count) $TraceType trace files"

    if ($TraceTypeUsesReferenceData)
    {
        write-host "$(get-date) Found $($RefEtlFiles.count) ReferenceData trace files"
    }
}

if ($UseTorus)
{
    CopyEtlFiles -EtlFiles $EtlFiles -WorkingPath $WorkingPath -UseTorus -Server $ExServer
    if ($TraceTypeUsesReferenceData)
    {
        CopyEtlFiles -EtlFiles $RefEtlFiles -WorkingPath "$WorkingPath\ReferenceData" -UseTorus -Server $ExServer
    }
}
else
{
    CopyEtlFiles -EtlFiles $EtlFiles -WorkingPath $WorkingPath
    if ($TraceTypeUsesReferenceData)
    {
        CopyEtlFiles -EtlFiles $RefEtlFiles -WorkingPath "$WorkingPath\ReferenceData"
    }
}

[string[]]$modules = "Microsoft.Exchange.Diagnostics.dll", `
"Microsoft.Exchange.Management.DataMining.Configuration.dll", `
"Microsoft.Exchange.Management.DataMining.LogTransformerCommon.dll", `
"Microsoft.Exchange.Management.DataMining.EventTracingTransformer.dll"
    
$LoadedModules=@()

ForEach ($Module in $Modules)
{
    $ModuleName = $Module.replace(".dll","")
    $LoadedModules += get-module $ModuleName
}
If ($LoadedModules.count -eq $Modules.count)
{
    Write-Host "$(get-date) Transformer modules are already loaded in PowerShell process"
}
ElseIf ($LoadedModules.count -ne $Modules.count -and (test-path -Path $TransformerModulePath -PathType Container))
{
    LoadTransformerModules -Module:$modules -path $TransformerModulePath
}
Else
{
    write-warning "Invalid TransformerModulePath = '$TransformerModulePath', path must exit when modules are not loaded"
    return
}

TransformEtlFiles -WorkingPath $WorkingPath
TransformEtlFiles -WorkingPath "$WorkingPath\ReferenceData"

#Load Reference Data
if ($TraceTypeUsesReferenceData)
{
    write-Host "$(get-date) Loading reference data"

    #Initialize Hash Tables
    $ACTIVITY=@{}
    $ADMINMETHOD=@{}
    $BREADCRUMBKIND=@{}
    $CLIENTTYPE=@{}
    $DATABASEINFO=@{}
    $ERRORCODE=@{}
    $MAILBOXINFO=@{}
    $OPERATIONDETAIL=@{}
    $OPERATIONSOURCE=@{}
    $OPERATIONTYPE=@{}
    $ROPID=@{}
    $TASKTYPE=@{}
    $MAILBOXSTATUS=@{}
    $MAILBOXTYPE=@{}
    $MAILBOXTYPEDETAIL=@{}

    # TODO - Need reference traces for these types
    $SEARCHSTATE=@{
        0x00000000 = "None";
        0x00000001 = "Running";
        0x00000002 = "Rebuild";
        0x00000004 = "Recursive";
        0x00000010 = "Foreground";
        0x00001000 = "AccurateResults";
        0x00002000 = "PotentiallyInaccurateResults";
        0x00010000 = "Static";
        0x00020000 = "InstantSearch";
        0x00080000 = "StatisticsOnly";
        0x00100000 = "CiOnly";
        0x00200000 = "FullTextIndexQueryFailed";
        0x00400000 = "EstimateCountOnly";
        0x01000000 = "CiTotally";
        0x02000000 = "CiWithTwirResidual";
        0x04000000 = "TwirMostly";
        0x08000000 = "TwirTotally";
        0x10000000 = "Error"}
    $SETSEARCHCRITERIAFLAGS=@{
        0x00000000 = "None";
        0x00000001 = "Stop";
        0x00000002 = "Restart";
        0x00000004 = "Recursive";
        0x00000008 = "Shallow";
        0x00000010 = "Foreground";
        0x00000020 = "Background";
        0x00004000 = "UseCIForComplexQueries";
        0x00010000 = "ContentIndexed";
        0x00020000 = "NonContentIndexed";
        0x00040000 = "Static";
        0x00800000 = "FailOnForeignEID";
        0x01000000 = "StatisticsOnly";
        0x02000000 = "FailNonContentIndexedSearch";
        0x04000000 = "EstimateCountOnly"}
    $LOGOPERATIONTYPE=@{
        "1" = "SearchFolderPopulationStart";
        "2" = "RequestSentToFAST";
        "3" = "ResponseReceivedFromFAST";
        "4" = "ReceviedErrorFromFAST";
        "5" = "FirstResultsLinked";
        "6" = "RequestCompleted";
        "7" = "ViewOperation";
        "8" = "ViewSetSearchCriteria";
        "9" = "OtherOperation"}

    $ReferenceDataFiles = @(Get-ChildItem "$WorkingPath\ReferenceData\*.csv")

    ForEach ($file in $ReferenceDataFiles)
    {
        $DataType = "$($file.basename.split("_")[0])"

        Switch ($DataType.ToUpper())
        {
            ACTIVITY
            {
                $ACTIVITY = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $ACTIVITY
                break;
            }

            ADMINMETHOD
            {
                $ADMINMETHOD = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $ADMINMETHOD
                break;
            }

            BREADCRUMBKIND
            {
                $BREADCRUMBKIND = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $BREADCRUMBKIND
                break;
            }

            CLIENTTYPE
            {
                $CLIENTTYPE = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $CLIENTTYPE
                break;
            }

            DATABASEINFO
            {
                $DATABASEINFO = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $DATABASEINFO
                break;
            }

            ERRORCODE
            {
                $ERRORCODE = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $ERRORCODE
                break;
            }

            MAILBOXINFO
            {
                $MAILBOXINFO = LoadMailboxInfo -FilePath $File.FullName -HashTable $MAILBOXINFO
                break;
            }

            OPERATIONDETAIL
            {
                $OPERATIONDETAIL = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $OPERATIONDETAIL
                break;
            }

            OPERATIONSOURCE
            {
                $OPERATIONSOURCE = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $OPERATIONSOURCE
                break;
            }

            OPERATIONTYPE
            {
                $OPERATIONTYPE = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $OPERATIONTYPE
                break;
            }

            ROPID
            {
                $ROPID = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $ROPID
                break;
            }

            TASKTYPE
            {
                $TASKTYPE = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $TASKTYPE
                break;
            }

            MAILBOXSTATUS
            {
                $MAILBOXSTATUS = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $MAILBOXSTATUS
                break;
            }

            MAILBOXTYPE
            {
                $MAILBOXTYPE = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $MAILBOXTYPE
                break;
            }

            MAILBOXTYPEDETAIL
            {
                $MAILBOXTYPEDETAIL = LoadReferenceData -FilePath $File.FullName -DataType:$datatype -HashTable $MAILBOXTYPEDETAIL
                break;
            }
        }
    }

    If ($MailboxGuid)
    {
        #Verify that MailboxGuid exists in MAILBOXINFO before proceeding
        If (!$MAILBOXINFO)
        {
            Write-Warning "$(get-date) No Mailboxes found in ReferenceData trace files, filtering will find no results"
            return
        }
        Else
        {
            $MailboxInfoEntry = $MAILBOXINFO.values | ? {$_ -eq $MailboxGuid}
            If (!$MailboxInfoEntry)
            {
                $MailboxNumber = GetMailboxNumber -DatabaseName $mailbox.Database.ToString() -MailboxGuid $MailboxGuid
                If ($MailboxNumber)
                {
                    $DatabaseHash = $database.guid.getHashCode()
                    $key = "$($DatabaseHash),$($MailboxNumber)"
                    write-host "$(get-date) MailboxGuid '$MailboxGuid' not found in ReferenceData, adding '$key' to `$MAILBOXINFO"  
                    $MAILBOXINFO[$key] = [guid]$MailboxGuid
                }
                Else
                {
                    Write-Warning "$(get-date) MailboxGuid '$MailboxGuid' not found in ReferenceData, filtering will find no results"                
                }
            }
        }
    }

    If ($Database)
    {
        #Verify that Database exists before proceeding
        If (!$DATABASEINFO)
        {
            Write-Warning "$(get-date) No Databases found in ReferenceData trace files, filtering will find no results"
            return
        }
        Else
        {
            $DatabaseReferences = @{}

            $DATABASEINFO.values | % { if (!$DatabaseReferences.Contains($_)) { $DatabaseReferences[$_] = 1 } }

            If (!$DatabaseReferences.contains($Database.name))
            {
                Write-Warning "$(get-date) Database '$($Database.name)' not found in ReferenceData trace files, filtering will find no results"
                return
            }
        }
    }
}

$ExtraColumns = @("*")

If ($TraceType -eq 'FullTextIndexQuery')
{
    [string]$TransformFilePath = join-Path -Path ($WorkingPath) -ChildPath "FullText*.csv"
}
Else
{
    [string]$TransformFilePath = join-Path -Path ($WorkingPath) -ChildPath "$TraceType*.csv"
}
$LogFiles = @(get-ChildItem $TransformFilePath)

if ($TraceType -eq "HeavyClientActivity")
{
    if($IncludeDetailTrace)
    {
        $fileTimeSpan=@{}
        $DetailFiles = $LogFiles | ?{$_.FullName -match "HeavyClientActivityDetail_"}
    }    
    $LogFiles = $LogFiles | ?{$_.FullName -match "HeavyClientActivity_"}
}
elseif ($TraceType -eq "FullTextIndexQuery")
{
    $FullTextIndexQueryHeaders = @(
        "Sequence,TimeStamp,TraceVersion,CorrelationId,DatabaseGuid,MailboxNumber,OperationType,ClientType,LogString")

    $FullTextIndexSingleLineHeaders = @(
        "Sequence,TimeStamp,TraceVersion,CorrelationId,DatabaseGuid,MailboxGuid,QueryStartTime,QueryEndTime,MailboxNumber,ClientType,QueryString,Failed,ErrorMessage,MaxLinkCountReached,StoreResidual,NumberFastTripes,Pulsing,FirstScopeFolder,MaxCount,ScoprFolderCount,InitialSearchState,FinalSearchState,SetSearchCriteriaFlags,IsNestedSearchFolder,First1000FastResults,FirstNotificationFastResults,FastResults,TotalResults,SearchRestrictionTime,SearchPlanTime,First1000FastTime,FirstResultsFastTime,FastTimes,FirstResultsTime,FastTime,ExpandedScopeFolderCount,FriendlyFolderName,TotalRowsProcessed,ClientActionString,EncryptedQuery,Replacements",
        "Sequence,TimeStamp,TraceVersion,CorrelationId,DatabaseGuid,MailboxGuid,QueryStartTime,QueryEndTime,MailboxNumber,ClientType,QueryString,Failed,ErrorMessage,MaxLinkCountReached,StoreResidual,NumberFastTripes,Pulsing,FirstScopeFolder,MaxCount,ScopeFolderCount,InitialSearchState,FinalSearchState,SetSearchCriteriaFlags,IsNestedSearchFolder,First1000FastResults,FirstNotificationFastResults,FastResults,TotalResults,SearchRestrictionTime,SearchPlanTime,First1000FastTime,FirstResultsFastTime,FastTimes,FirstResultsTime,FastTime,ExpandedScopeFolderCount,FriendlyFolderName,TotalRowsProcessed,ClientActionString,EncryptedQuery,Replacements",
        "Sequence,TimeStamp,TraceVersion,CorrelationId,DatabaseGuid,MailboxGuid,QueryStartTime,QueryEndTime,MailboxNumber,ClientType,QueryString,Failed,ErrorMessage,MaxLinkCountReached,StoreResidual,NumberFastTripes,Pulsing,FirstScopeFolder,MaxCount,ScopeFolderCount,InitialSearchState,FinalSearchState,SetSearchCriteriaFlags,IsNestedSearchFolder,First1000FastResults,FirstNotificationFastResults,FastResults,TotalResults,SearchRestrictionTime,SearchPlanTime,First1000FastTime,FirstResultsFastTime,FastTimes,FirstResultsTime,FastTime,ExpandedScopeFolderCount,FriendlyFolderName,TotalRowsProcessed,ClientActionString,EncryptedQuery,Replacements,DatabaseReadWaitTime")

    if($IncludeDetailTrace)
    {
        $fileTimeSpan=@{}
        $DetailFiles = $LogFiles | ?{$_.FullName -match "FullTextQueryDetail_"}
    }    

    $LogFiles = $LogFiles | ?{
        if ($_.FullName -match "FullTextIndexQuery_")
        {
            $headerLine = Get-Content $_.FullName -TotalCount 1

            if ($FullTextIndexTraceType -eq "FullTextIndexSingleLine")
            {
                $headerLine -in $FullTextIndexSingleLineHeaders
            }
            elseif ($FullTextIndexTraceType -eq "FullTextIndexQuery")
            {
                $headerLine -in $FullTextIndexQueryHeaders
            }
            else
            {
                throw "Unexpected value for FullTextIndexTraceType: $FullTextIndexTraceType"
            }
        }
        else
        {
            $false
        }
    }
}
elseif ($TraceType -eq "LongOperation")
{
    if($IncludeDetailTrace)
    {
        $fileTimeSpan=@{}
        $DetailFiles = $LogFiles | ?{$_.FullName -match "LongOperation_"}
    }    
    $LogFiles = $LogFiles | ?{$_.FullName -match "LongOperationSummary_"}
}
elseif ($TraceType -eq "RopSummary")
{
    $ExtraColumns += @("ViewSignatureHash");    
    if($IncludeDetailTrace)
    {
        $ExtraColumns += @("SelectColumnsHash", 
            "SortColumnsHash", 
            "SelectColumns", 
            "SortColumns", 
            "RestrictionHash", 
            "Restriction", 
            "ConfigFlags", 
            "OperationSignatureHash", 
            "OperationSignature");
    }    
}

write-Host "$(get-date) Importing $($LogFiles.count) $TraceType files"
write-Host "$(get-date) Start of time frame to collect is $Start ($($Start.ToUniversalTime()) UTC)"
write-Host "$(get-date) End of time frame to collect is $End ($($End.ToUniversalTime()) UTC)"

$RopSummaryHasSchemaVersion = $null
$FullTextIndexSingleLineHasTypo = $null
$LongOperationSummaryVersion = $null
$MailboxInfoVersion = $null
$DatabasePageSize = 32KB
$MailboxSizeUnits = 1GB
$MessageSizeUnits = 1MB
$RulesSizeUnits = 1MB

$LogFiles | % {write-host "$(get-date) Importing $($_.FullName)"; Import-csv -Path $_.FullName} |`
    foreach `
    { `
        $_.Timestamp = get-date("$($_.timestamp)Z"); `
        `
        if ($TraceType -eq 'InstantSearchBigFunnel') `
        { `
            Add-Member -InputObject $_ -MemberType NoteProperty -Name DatabaseGuid -Value $_.Database; `
            $_.PSObject.Properties.Remove('Database'); `
        } `
        if ($TraceTypeUsesReferenceData -and $TraceTypeFileName -ne 'ReferenceData') `
        { `
            if ($TraceType -eq 'BreadCrumbs') `
            { `
                Add-Member -InputObject $_ -MemberType NoteProperty -Name Database -Value $_.DatabaseHash; `
                Add-Member -InputObject $_ -MemberType NoteProperty -Name Mailbox -Value $_.MailboxNumber; `
                Add-Member -InputObject $_ -MemberType NoteProperty -Name Client -Value $_.ClientType; `
                $_.PSObject.Properties.Remove('DatabaseHash'); `
                $_.PSObject.Properties.Remove('MailboxNumber'); `
                $_.PSObject.Properties.Remove('ClientType'); `
            } `
            elseif ($TraceType -eq 'FullTextIndexQuery') `
            { `
                Add-Member -InputObject $_ -MemberType NoteProperty -Name Database -Value ([Guid]$_.DatabaseGuid); `
                Add-Member -InputObject $_ -MemberType NoteProperty -Name Client -Value $_.ClientType; `
                $_.PSObject.Properties.Remove('DatabaseGuid'); `
                $_.PSObject.Properties.Remove('ClientType'); `

                if ($FullTextIndexTraceType -eq 'FullTextIndexSingleLine') `
                { `
                    if ($FullTextIndexSingleLineHasTypo -eq $null) `
                    { `
                        if ((Get-Member -InputObject $_ -Name 'ScoprFolderCount') -eq $null) `
                        { `
                            $FullTextIndexSingleLineHasTypo = $false `
                        } `
                        else `
                        { `
                            $FullTextIndexSingleLineHasTypo = $true `
                        } `
                    } `

                    if ($FullTextIndexSingleLineHasTypo)
                    {
                        Add-Member -InputObject $_ -MemberType NoteProperty -Name ScopeFolderCount -Value $_.ScoprFolderCount; `
                        $_.PSObject.Properties.Remove('ScoprFolderCount'); `
                    }

                    Add-Member -InputObject $_ -MemberType NoteProperty -Name Mailbox -Value ([Guid]$_.MailboxGuid); `
                    $_.PSObject.Properties.Remove('MailboxNumber'); `
                    $_.PSObject.Properties.Remove('MailboxGuid'); `
                } `
                elseif ($FullTextIndexTraceType -eq 'FullTextIndexQuery') `
                { `
                    Add-Member -InputObject $_ -MemberType NoteProperty -Name Mailbox -Value $_.MailboxNumber; `
                    $_.PSObject.Properties.Remove('MailboxNumber'); `
                } `
            } `

            if ($MAILBOXINFO) `
            { `
                if ($_.Database -is [Guid]) `
                { `
                    $key = "$($_.Database.GetHashCode()),$($_.Mailbox)"; 
                } `
                else `
                { `
                    $key = "$($_.Database),$($_.Mailbox)"; `
                } `
                `
                if ($MAILBOXINFO.contains($key)) `
                { `
                    $_.Mailbox = $MAILBOXINFO[$key] `
                } `
            } `
            `
            if ($_.Database -and $DATABASEINFO.contains($_.Database)) `
            { `
                $_.Database = $DATABASEINFO[$_.Database]
            } `
        } `

        write-output $_ `
    } | `
    where `
    { `
        ( $_.Timestamp -ge $Start -and $_.Timestamp -le $End ) `
    } | `
    where `
    { `
        if ($TraceTypeUsesReferenceData -and $TraceTypeFileName -ne 'ReferenceData') `
        { `
            ( $MailboxGuid -and $_.Mailbox -is [System.Guid] -and $_.Mailbox -eq $MailboxGuid ) -or `
            ( !$MailboxGuid) `
        } `
        else `
        { `
            ( $MailboxGuid -and $_.MailboxGuid -eq $MailboxGuid ) -or `
            ( !$MailboxGuid) `
        } `
    } | `
    where `
    { `
        if ($TraceTypeUsesReferenceData -and $TraceTypeFileName -ne 'ReferenceData') `
        { `
            ( $Database -and $_.Database -eq $Database.Name) -or `
            ( !$Database) `
        } `
        else `
        { `
            ( $Database -and $_.DatabaseGuid -eq $Database.Guid) -or `
            ( !$Database) 
        } `
    } | select $ExtraColumns |`
    foreach `
    { `
        $_.Sequence = [Long]$_.Sequence; `
        `
        if ($TraceTypeUsesReferenceData -and $TraceTypeFileName -ne 'ReferenceData') `
        {
            if ($CLIENTTYPE) `
            { `
                $_.Client = $CLIENTTYPE[$_.Client] `
            } `
            `
            if ($TraceType -ne 'BreadCrumbs' -and `
                $TraceType -ne 'MailboxAggregateRops' -and `
                $TraceType -ne 'FullTextIndexQuery') `
            { `
                if ($ACTIVITY) `
                { `
                    $_.Activity = $ACTIVITY[$_.Activity] `
                } `
            } `
        } `
        `
        if ($TraceType -eq 'RopSummary') `
        { `
            `
            if ($RopSummaryHasSchemaVersion -eq $null) `
            { `
                if ((Get-Member -InputObject $_ -Name 'SchemaVersion') -eq $null) `
                { `
                    $RopSummaryHasSchemaVersion = $false `
                } `
                else `
                { `
                    $RopSummaryHasSchemaVersion = $true `
                } `
            } `

            if ($RopSummaryHasSchemaVersion) `
            { `
                $_.SchemaVersion = [Int]$_.SchemaVersion; `
                
                if ($_.SchemaVersion -eq 100) `
                { `
                    $_.ContextFlags = [Int]$_.ContextFlags; `
                    
                    if ($_.ContextFlags -band 0x1 -eq 0x1) `
                    { `
                        Add-Member -InputObject $_ -MemberType NoteProperty -Name ActiveCopy -Value $false `
                    } `
                    else `
                    { `
                        Add-Member -InputObject $_ -MemberType NoteProperty -Name ActiveCopy -Value $true `
                    } `

                    $_.PSObject.Properties.Remove('ContextFlags'); `

                    if ($OPERATIONTYPE) `
                    { `
                        $_.OperationType = $OPERATIONTYPE[$_.OperationType] `
                    } `

                    if ($ROPID -and $_.OperationType -eq 'Rop') `
                    { `
                        $_.Operation = $ROPID[$_.Operation] `
                    } `
                    elseif ($TASKTYPE -and $_.OperationType -eq 'Task') `
                    { `
                        $_.Operation = $TASKTYPE[$_.Operation] `
                    } `
                    elseif ($ADMINMETHOD -and $_.OperationType -eq 'Admin') `
                    { `
                        $_.Operation = $ADMINMETHOD[$_.Operation] `
                    } `
                    `
                    if ($OPERATIONDETAIL) `
                    { `
                        $detail = $OPERATIONDETAIL[$_.Detail];`
                        $_.Detail = $detail.Value;`
                    } `
                    `
                    $_.SharedLock = if ([int]$_.SharedLock -eq 0) { $false } else { $true }; `
                    
                    $_.TotalLogBytes = [Int]$_.TotalLogBytes; `
                    $_.TotalPagesPreread = [Int]$_.TotalPagesPreread; `
                    $_.TotalPagesRead = [Int]$_.TotalPagesRead; `
                    $_.TotalPagesDirtied = [Int]$_.TotalPagesDirtied; `
                    $_.TotalPagesRedirtied = [Int]$_.TotalPagesRedirtied; `
                    $_.TotalJetReservedAlpha = [Int]$_.TotalJetReservedAlpha; `
                    $_.TotalJetReservedBeta = [Int]$_.TotalJetReservedBeta; `
                    $_.TotalDirectoryOperations = [Int]$_.TotalDirectoryOperations; `
                    $_.TotalOffPageHits = [Int]$_.TotalOffPageHits; `
                    $_.TotalCpuTimeKernel = [Int]$_.TotalCpuTimeKernel; `
                    $_.TotalCpuTimeUser = [Int]$_.TotalCpuTimeUser; `
                    $_.TotalChunks = [Int]$_.TotalChunks; `
                    $_.MaxChunkTime = [Int]$_.MaxChunkTime; `
                    $_.TotalLockWaitTime = [Int]$_.TotalLockWaitTime; `
                    $_.TotalDirectoryWaitTime = [Int]$_.TotalDirectoryWaitTime; `
                    $_.TotalDatabaseTime = [Int]$_.TotalDatabaseTime; `
                    $_.TotalFastWaitTime = [Int]$_.TotalFastWaitTime; `
                    $_.TotalDirectoryMServOperations = [Int]$_.TotalDirectoryMServOperations; `
                    $_.TotalUndefinedBeta = [Int]$_.TotalUndefinedBeta; `
                    $_.TotalUndefinedGamma = [Int]$_.TotalUndefinedGamma; `
                    $_.TotalUndefinedDelta = [Int]$_.TotalUndefinedDelta; `
                    $_.TotalUndefinedOmega = [Int]$_.TotalUndefinedOmega; `
                } `
                elseif ($_.SchemaVersion -eq 101 -or $_.SchemaVersion -eq 102 -or $_.SchemaVersion -eq 103) `
                { `
                    $_.ContextFlags = [Int]$_.ContextFlags; `
                    
                    if ($_.ContextFlags -band 0x1 -eq 0x1) `
                    { `
                        Add-Member -InputObject $_ -MemberType NoteProperty -Name ActiveCopy -Value $false `
                    } `
                    else `
                    { `
                        Add-Member -InputObject $_ -MemberType NoteProperty -Name ActiveCopy -Value $true `
                    } `

                    $_.PSObject.Properties.Remove('ContextFlags'); `

                    if ($OPERATIONTYPE) `
                    { `
                        $_.OperationType = $OPERATIONTYPE[$_.OperationType] `
                    } `

                    if ($ROPID -and $_.OperationType -eq 'Rop') `
                    { `
                        $_.Operation = $ROPID[$_.Operation] `
                    } `
                    elseif ($TASKTYPE -and $_.OperationType -eq 'Task') `
                    { `
                        $_.Operation = $TASKTYPE[$_.Operation] `
                    } `
                    elseif ($ADMINMETHOD -and $_.OperationType -eq 'Admin') `
                    { `
                        $_.Operation = $ADMINMETHOD[$_.Operation] `
                    } `
                    `
                    if ($OPERATIONDETAIL) `
                    { `
                        $detail = $OPERATIONDETAIL[$_.Detail];`
                        $_.Detail = $detail.Value;`
                        if($detail.SelectColumnsHash)`
                        {`
                            $_.ViewSignatureHash = $detail.Key;`
                            if($IncludeDetailTrace)`
                            {`
                                $_.SelectColumnsHash = $detail.SelectColumnsHash;`
                                $_.SortColumnsHash = $detail.SortColumnsHash;`
                                $_.SelectColumns = $detail.SelectColumns;`
                                $_.SortColumns = $detail.SortColumns;`
                                $_.RestrictionHash = $detail.RestrictionHash;`
                                $_.Restriction = $detail.Restriction;`
                                $_.ConfigFlags = $detail.ConfigFlags;`
                                $_.OperationSignatureHash = $detail.OperationSignatureHash;`
                                $_.OperationSignature = $detail.OperationSignature;`
                            }`
                        }`
                    } `
                    `
                    $_.SharedLock = if ([int]$_.SharedLock -eq 0) { $false } else { $true }; `
                    
                    $_.TotalLogBytes = [Int]$_.TotalLogBytes; `
                    $_.TotalPagesPreread = [Int]$_.TotalPagesPreread; `
                    $_.TotalPagesRead = [Int]$_.TotalPagesRead; `
                    $_.TotalPagesDirtied = [Int]$_.TotalPagesDirtied; `
                    $_.TotalPagesRedirtied = [Int]$_.TotalPagesRedirtied; `
                    $_.TotalDatabaseReadWaitTime = [Int]$_.TotalDatabaseReadWaitTime; `
                    $_.TotalJetReservedBeta = [Int]$_.TotalJetReservedBeta; `
                    $_.TotalDirectoryOperations = [Int]$_.TotalDirectoryOperations; `
                    $_.TotalOffPageHits = [Int]$_.TotalOffPageHits; `
                    $_.TotalCpuTimeKernel = [Int]$_.TotalCpuTimeKernel; `
                    $_.TotalCpuTimeUser = [Int]$_.TotalCpuTimeUser; `
                    $_.TotalChunks = [Int]$_.TotalChunks; `
                    $_.MaxChunkTime = [Int]$_.MaxChunkTime; `
                    $_.TotalLockWaitTime = [Int]$_.TotalLockWaitTime; `
                    $_.TotalDirectoryWaitTime = [Int]$_.TotalDirectoryWaitTime; `
                    $_.TotalDatabaseTime = [Int]$_.TotalDatabaseTime; `
                    $_.TotalFastWaitTime = [Int]$_.TotalFastWaitTime; `
                    $_.TotalDirectoryMServOperations = [Int]$_.TotalDirectoryMServOperations; `
                    $_.TotalUndefinedBeta = [Int]$_.TotalUndefinedBeta; `
                    $_.TotalUndefinedGamma = [Int]$_.TotalUndefinedGamma; `
                    $_.TotalUndefinedDelta = [Int]$_.TotalUndefinedDelta; `
                    $_.TotalUndefinedOmega = [Int]$_.TotalUndefinedOmega; `
                } `
                else `
                { `
                    throw "Unexpected SchemaVersion: $($_.SchemaVersion)" `
                } `
            } `
            else `
            { `
                if ($ROPID) `
                { `
                    $_.Operation = $ROPID[$_.Operation] `
                } `
            } `
            `
            if ($ERRORCODE -and $_.LastError -ne 0) `
            { `
                $_.LastError = $ERRORCODE[$_.LastError]
            } `

            $_.TotalCalls = [Int]$_.TotalCalls; `
            $_.NumSlow = [Int]$_.NumSlow; `
            $_.MaxElapsed = [Int]$_.MaxElapsed; `
            $_.NumErrored = [Int]$_.NumErrored; `
            $_.TotalTime = [Int]$_.TotalTime; `
            $_.NumActivities = [Int]$_.NumActivities; `
        } `
        `
        elseif ($TraceType -eq 'DatabaseAggregateRops') `
        { `
            `
            $_.SchemaVersion = [Int]$_.SchemaVersion; `
            $_.ContextFlags = [Int]$_.ContextFlags; `
            `
            if ($_.ContextFlags -band 0x1 -eq 0x1) `
            { `
                Add-Member -InputObject $_ -MemberType NoteProperty -Name ActiveCopy -Value $false `
            } `
            else `
            { `
                Add-Member -InputObject $_ -MemberType NoteProperty -Name ActiveCopy -Value $true `
            } `
            `
            $_.PSObject.Properties.Remove('ContextFlags'); `
            `
            if ($_.MailboxCategory -eq 0)
            { `
                $_.MailboxCategory = "None" `
            } `
            elseif ($_.MailboxCategory -eq 1)
            { `
                $_.MailboxCategory = "Consumer" `
            } `
            elseif ($_.MailboxCategory -eq 2)
            { `
                $_.MailboxCategory = "Business" `
            } `
            `
            if ($OPERATIONTYPE) `
            { `
                $_.OperationType = $OPERATIONTYPE[$_.OperationType] `
            } `

            if ($ROPID -and $_.OperationType -eq 'Rop') `
            { `
                $_.Operation = $ROPID[$_.Operation] `
            } `
            elseif ($TASKTYPE -and $_.OperationType -eq 'Task') `
            { `
                $_.Operation = $TASKTYPE[$_.Operation] `
            } `
            elseif ($ADMINMETHOD -and $_.OperationType -eq 'Admin') `
            { `
                $_.Operation = $ADMINMETHOD[$_.Operation] `
            } `
            `
            $_.SharedLock = if ([int]$_.SharedLock -eq 0) { $false } else { $true }; `
            if ($OPERATIONDETAIL) `
            { `
                $detail = $OPERATIONDETAIL[$_.Detail];`
                $_.Detail = $detail.Value;`
                if($detail.SelectColumnsHash)`
                {`
                    if($IncludeDetailTrace)`
                    {`
                        $_.SelectColumnsHash = $detail.SelectColumnsHash;`
                        $_.SortColumnsHash = $detail.SortColumnsHash;`
                        $_.SelectColumns = $detail.SelectColumns;`
                        $_.SortColumns = $detail.SortColumns;`
                        $_.RestrictionHash = $detail.RestrictionHash;`
                        $_.Restriction = $detail.Restriction;`
                        $_.ConfigFlags = $detail.ConfigFlags;`
                        $_.OperationSignatureHash = $detail.OperationSignatureHash;`
                        $_.OperationSignature = $detail.OperationSignature;`
                    }`
                }`
            } `
            `
            $_.TotalLogBytes = [Int]$_.TotalLogBytes; `
            $_.TotalPagesPreread = [Int]$_.TotalPagesPreread; `
            $_.TotalPagesRead = [Int]$_.TotalPagesRead; `
            $_.TotalPagesDirtied = [Int]$_.TotalPagesDirtied; `
            $_.TotalPagesRedirtied = [Int]$_.TotalPagesRedirtied; `
            $_.TotalDatabaseReadWaitTime = [Int]$_.TotalDatabaseReadWaitTime; `
            $_.TotalJetReservedBeta = [Int]$_.TotalJetReservedBeta; `
            $_.TotalDirectoryOperations = [Int]$_.TotalDirectoryOperations; `
            $_.TotalOffPageHits = [Int]$_.TotalOffPageHits; `
            $_.TotalCpuTimeKernel = [Int]$_.TotalCpuTimeKernel; `
            $_.TotalCpuTimeUser = [Int]$_.TotalCpuTimeUser; `
            $_.TotalChunks = [Int]$_.TotalChunks; `
            $_.MaxChunkTime = [Int]$_.MaxChunkTime; `
            $_.TotalLockWaitTime = [Int]$_.TotalLockWaitTime; `
            $_.TotalDirectoryWaitTime = [Int]$_.TotalDirectoryWaitTime; `
            $_.TotalDatabaseTime = [Int]$_.TotalDatabaseTime; `
            $_.TotalFastWaitTime = [Int]$_.TotalFastWaitTime; `
            $_.TotalDirectoryMServOperations = [Int]$_.TotalDirectoryMServOperations; `
            `
            if ($ERRORCODE -and $_.LastError -ne 0) `
            { `
                $_.LastError = $ERRORCODE[$_.LastError]
            } `
            `
            $_.TotalCalls = [Int]$_.TotalCalls; `
            $_.NumSlow = [Int]$_.NumSlow; `
            $_.MaxElapsed = [Int]$_.MaxElapsed; `
            $_.NumErrored = [Int]$_.NumErrored; `
            $_.TotalTime = [Int]$_.TotalTime; `
            $_.NumActivities = [Int]$_.NumActivities; `
        } `
        `
        elseif ($TraceType -eq 'MailboxAggregateRops') `
        { `
            `
            $_.SchemaVersion = [Int]$_.SchemaVersion; `
            $_.ContextFlags = [Int]$_.ContextFlags; `
            `
            if ($_.ContextFlags -band 0x1 -eq 0x1) `
            { `
                Add-Member -InputObject $_ -MemberType NoteProperty -Name ActiveCopy -Value $false `
            } `
            else `
            { `
                Add-Member -InputObject $_ -MemberType NoteProperty -Name ActiveCopy -Value $true `
            } `
            `
            $_.PSObject.Properties.Remove('ContextFlags'); `
            `
            if ($_.MailboxCategory -eq 0)
            { `
                $_.MailboxCategory = "None" `
            } `
            elseif ($_.MailboxCategory -eq 1)
            { `
                $_.MailboxCategory = "Consumer" `
            } `
            elseif ($_.MailboxCategory -eq 2)
            { `
                $_.MailboxCategory = "Business" `
            } `
            `
            $_.TotalLogBytes = [Int]$_.TotalLogBytes; `
            $_.TotalPagesPreread = [Int]$_.TotalPagesPreread; `
            $_.TotalPagesRead = [Int]$_.TotalPagesRead; `
            $_.TotalPagesDirtied = [Int]$_.TotalPagesDirtied; `
            $_.TotalPagesRedirtied = [Int]$_.TotalPagesRedirtied; `
            $_.TotalDatabaseReadWaitTime = [Int]$_.TotalDatabaseReadWaitTime; `
            $_.TotalJetReservedBeta = [Int]$_.TotalJetReservedBeta; `
            $_.TotalDirectoryOperations = [Int]$_.TotalDirectoryOperations; `
            $_.TotalOffPageHits = [Int]$_.TotalOffPageHits; `
            $_.TotalCpuTimeKernel = [Int]$_.TotalCpuTimeKernel; `
            $_.TotalCpuTimeUser = [Int]$_.TotalCpuTimeUser; `
            $_.TotalChunks = [Int]$_.TotalChunks; `
            $_.MaxChunkTime = [Int]$_.MaxChunkTime; `
            $_.TotalLockWaitTime = [Int]$_.TotalLockWaitTime; `
            $_.TotalDirectoryWaitTime = [Int]$_.TotalDirectoryWaitTime; `
            $_.TotalDatabaseTime = [Int]$_.TotalDatabaseTime; `
            $_.TotalFastWaitTime = [Int]$_.TotalFastWaitTime; `
            $_.TotalDirectoryMServOperations = [Int]$_.TotalDirectoryMServOperations; `
            `
            if ($ERRORCODE -and $_.LastError -ne 0) `
            { `
                $_.LastError = $ERRORCODE[$_.LastError]
            } `

            $_.TotalCalls = [Int]$_.TotalCalls; `
            $_.NumSlow = [Int]$_.NumSlow; `
            $_.MaxElapsed = [Int]$_.MaxElapsed; `
            $_.NumErrored = [Int]$_.NumErrored; `
            $_.TotalTime = [Int]$_.TotalTime; `
            $_.NumActivities = [Int]$_.NumActivities; `
        } `
        `
        elseif ($TraceType -eq 'RopResource') `
        { `
            if ($ROPID) `
            { `
                $_.Operation = $ROPID[$_.Operation] `
            } `
            `
            $_.TotalCalls = [Int]$_.TotalCalls; `
            $_.TotalLogBytes = [Int]$_.TotalLogBytes; `
            $_.TotalLogRecords = [Int]$_.TotalLogRecords; `
            $_.TotalPagesDirtied = [Int]$_.TotalPagesDirtied; `
            $_.TotalPagesPreread = [Int]$_.TotalPagesPreread; `
            $_.TotalPagesRead = [Int]$_.TotalPagesRead; `
            $_.TotalPagesRedirtied = [Int]$_.TotalPagesRedirtied; `
            $_.TotalPagesReferenced = [Int]$_.TotalPagesReferenced; `
            $_.TotalDirectoryOperations = [Int]$_.TotalDirectoryOperations; `
            $_.TotalOffPageHits = [Int]$_.TotalOffPageHits; `
            $_.TotalCpuTimeKernel = [Int]$_.TotalCpuTimeKernel; `
            $_.TotalCpuTimeUser = [Int]$_.TotalCpuTimeUser; `
        } `
        `
        elseif ($TraceType -eq 'LockContention') `
        { `
            if ($CLIENTTYPE) `
            { `
                $_.MailboxClient = $CLIENTTYPE[$_.MailboxClient]; `
                $_.ComponentClient = $CLIENTTYPE[$_.ComponentClient]; `
                $_.OtherClient = $CLIENTTYPE[$_.OtherClient]; `
            } `
            `
            if ($ROPID) `
            { `
                $_.Operation = $ROPID[$_.Operation]; `
                $_.MailboxOperation = $ROPID[$_.MailboxOperation]; `
                $_.ComponentOperation = $ROPID[$_.ComponentOperation]; `
                $_.OtherOperation = $ROPID[$_.OtherOperation]; `
            } `
            `
            $_.MailboxTotal = [Int]$_.MailboxTotal; `
            $_.ComponentTotal = [Int]$_.ComponentTotal; `
            $_.OtherTotal = [Int]$_.OtherTotal; `
        } `
        `
        elseif ($TraceType -eq 'HeavyClientActivity') `
        {
            $_.DatabaseGuid = [Guid]$_.DatabaseGuid; `
            $_.MailboxGuid = [Guid]$_.MailboxGuid; `
            $_.TotalRpcCalls = [Int]$_.TotalRpcCalls; `
            $_.TotalRops = [Int]$_.TotalRops; `
        } `
        `
        elseif ($TraceType -eq 'BreadCrumbs') `
        { `
            if ($BREADCRUMBKIND) `
            { `
                $_.CrumbKind = $BREADCRUMBKIND[$_.CrumbKind] `
            } `
            `
            if ($ERRORCODE -and `
                ($_.CrumbKind -eq 'AdminError' -or `
                 $_.CrumbKind -eq 'Exception' -or `
                 $_.CrumbKind -eq 'RopError')) `
            { `
                $_.CrumbValue = $ERRORCODE[$_.CrumbValue] `
            } `
            `
            if ($OPERATIONSOURCE) `
            { `
                $_.Source = $OPERATIONSOURCE[$_.Source] `
            } `
            `
            if ($ROPID -and $_.Source -eq 'Mapi') `
            { `
                $_.Operation = $ROPID[$_.Operation] `
            } `
            elseif ($TASKTYPE -and `
                ($_.Source -eq 'MailboxTask' -or `
                 $_.Source -eq 'MapiTimedEvent' -or `
                 $_.Source -eq 'PerUserCacheFlush')) `
            { `
                $_.Operation = $TASKTYPE[$_.Operation] `
            } `
            elseif ($ADMINMETHOD -and `
                ($_.Source -eq 'AdminRpc' -or `
                 $_.Source -eq 'LogicalIndexCleanup' -or `
                 $_.Source -eq 'MailboxCleanup' -or `
                 $_.Source -eq 'MailboxMaintenance' -or `
                 $_.Source -eq 'OnlineIntegrityCheck' -or `
                 $_.Source -eq 'SearchFolderAgeOut' -or `
                 $_.Source -eq 'SubobjectsCleanup')) `
            { `
                $_.Operation = $ADMINMETHOD[$_.Operation] `
            } `
        } `
        elseif ($TraceType -eq 'FullTextIndexQuery') `
        { `
            $_.TraceVersion = [Int16]$_.TraceVersion; `
            $_.CorrelationId = [Guid]$_.CorrelationId; `

            if ($FullTextIndexTraceType -eq 'FullTextIndexSingleLine') `
            { `
                if ($SEARCHSTATE) `
                { `
                    $_.InitialSearchState = GetFlagsEnumString -ReferenceData $SEARCHSTATE -Value $_.InitialSearchState;
                    $_.FinalSearchState = GetFlagsEnumString -ReferenceData $SEARCHSTATE $_.FinalSearchState; `
                } `

                if ($SETSEARCHCRITERIAFLAGS) `
                { `
                    $_.SetSearchCriteriaFlags = GetFlagsEnumString -ReferenceData $SETSEARCHCRITERIAFLAGS -Value $_.SetSearchCriteriaFlags `
                } `

                $_.QueryStartTime = [DateTime]$_.QueryStartTime; `
                $_.QueryEndTime = [DateTime]$_.QueryEndTime; `
                $_.Failed = [Bool][Byte]$_.Failed; `
                $_.MaxLinkCountReached = [Bool][Byte]$_.MaxLinkCountReached; `
                $_.NumberFastTripes = [Int]$_.NumberFastTripes; `
                $_.Pulsing = [Bool][Byte]$_.Pulsing; `
                $_.MaxCount = [Int]$_.MaxCount; `
                $_.ScopeFolderCount = [Int]$_.ScopeFolderCount; `
                $_.IsNestedSearchFolder = [Bool][Byte]$_.IsNestedSearchFolder; `
                $_.First1000FastResults = [Int]$_.First1000FastResults; `
                $_.FirstNotificationFastResults = [Int]$_.FirstNotificationFastResults; `
                $_.FastResults = [Int]$_.FastResults; `
                $_.TotalResults = [Int]$_.TotalResults; `
                $_.SearchRestrictionTime = [Int]$_.SearchRestrictionTime; `
                $_.SearchPlanTime = [Int]$_.SearchPlanTime; `
                $_.First1000FastTime = [Int]$_.First1000FastTime; `
                $_.FirstResultsFastTime = [Int]$_.FirstResultsFastTime; `
                $_.FirstResultsTime = [Int]$_.FirstResultsTime; `
                $_.FastTime = [Int]$_.FastTime; `
                $_.ExpandedScopeFolderCount = [Int]$_.ExpandedScopeFolderCount; `
                $_.TotalRowsProcessed = [Int]$_.TotalRowsProcessed; `
                if ($_.TraceVersion -gt 6) `
                { `
                    $_.DatabaseReadWaitTime = [Int]$_.DatabaseReadWaitTime; `
                } `
            } `
            elseif ($FullTextIndexTraceType -eq 'FullTextIndexQuery') `
            { `
                if ($LOGOPERATIONTYPE) `
                { `
                    $_.OperationType = $LOGOPERATIONTYPE[$_.OperationType] `
                } ` 
            } `
        } `
        elseif ($TraceType -eq 'LongOperation') `
        { `
            if ($LongOperationSummaryVersion -eq $null) `
            { `
                if ((Get-Member -InputObject $_ -Name 'DatabaseReadWaitTime') -ne $null) `
                { `
                    $LongOperationSummaryVersion = 5 `
                } `
                elseif ((Get-Member -InputObject $_ -Name 'HashCode') -ne $null) `
                { `
                    $LongOperationSummaryVersion = 4 `
                } `
                elseif ((Get-Member -InputObject $_ -Name 'CorrelationId') -ne $null) `
                { `
                    $LongOperationSummaryVersion = 3 `
                } `
                elseif ((Get-Member -InputObject $_ -Name 'TimeInDatabase') -ne $null) `
                { `
                    $LongOperationSummaryVersion = 2 `
                } `
                else `
                { `
                    $LongOperationSummaryVersion = 1 `
                } `
            } `

            $_.DatabaseGuid = [Guid]$_.DatabaseGuid; `
            $_.MailboxGuid = [Guid]$_.MailboxGuid; `
            $_.ChunkElapsedTime = [Long]$_.ChunkElapsedTime; `
            $_.InteractionTime = [Long]$_.InteractionTime; `
            $_.PagesPreread = [Int]$_.PagesPreread; `
            $_.PagesRead = [Int]$_.PagesRead; `
            $_.PagesDirtied = [Int]$_.PagesDirtied; `
            $_.LogBytesWritten = [Int]$_.LogBytesWritten; `
            $_.SortOrderCount = [Int]$_.SortOrderCount; `
            $_.NumberPlansExecuted = [Int]$_.NumberPlansExecuted; `
            $_.PlansExecutionTime = [Long]$_.PlansExecutionTime; `
            $_.NumberDirectoryOps = [Int]$_.NumberDirectoryOps; `
            $_.DirectoryOpsTime = [Long]$_.DirectoryOpsTime; `
            $_.NumberLocksAttempted = [Int]$_.NumberLocksAttempted; `
            $_.NumberLocksSucceeded = [Int]$_.NumberLocksSucceeded; `
            $_.LocksWaitTime = [Long]$_.LocksWaitTime; `
            $_.IsLongOperation = [Bool][Byte]$_.IsLongOperation; `
            $_.IsResourceIntense = [Bool][Byte]$_.IsResourceIntense; `
            $_.IsContested = [Bool][Byte]$_.IsContested; `

            if ($LongOperationSummaryVersion -gt 1) `
            { `
                $_.TimeInDatabase = [Long]$_.TimeInDatabase; `
            } `

            if ($LongOperationSummaryVersion -gt 2) `
            { `
                $_.CorrelationId = [Guid]$_.CorrelationId; `
                $_.TimeInCpuKernel = [Long]$_.TimeInCpuKernel; `
                $_.TimeInCpuUser = [Long]$_.TimeInCpuUser; `
            } `

            if ($LongOperationSummaryVersion -gt 3) `
            { `
                $_.HashCode = [Int]$_.HashCode; `
            } `

            if ($LongOperationSummaryVersion -gt 4) `
            { `
                $_.DatabaseReadWaitTime = [Long]$_.DatabaseReadWaitTime; `
            } `

            if ($ExcludeNonResourceIntensiveOperations -and -not $_.IsResourceIntense) `
            { `
                $_ = $NULL; `
            } `

        } `
        elseif ($TraceType -eq 'InstantSearchDocumentId') `
        {
            $_.TraceVersion = [Int]$_.TraceVersion; `
            $_.DatabaseGuid = [Guid]$_.DatabaseGuid; `
            $_.MailboxGuid = [Guid]$_.MailboxGuid; `
            $_.ErrorCode = [Int]$_.ErrorCode; `
            $_.QueryRowsInstance = [Int]$_.QueryRowsInstance; `
            $_.QueryRowsTimeLimit = [Long]$_.QueryRowsTimeLimit; `
            $_.IsPreviewPromoted = [Bool]$_.IsPreviewPromoted; `
            $_.PrereadAge = [Long]$_.PrereadAge; `
            $_.PrereadBytes = [Long]$_.PrereadBytes; `
            $_.TotalDocumentIdCount = [Int]$_.TotalDocumentIdCount; `
            $_.Group1DocumentIdCount = [Int]$_.Group1DocumentIdCount; `
            $_.Group2DocumentIdCount = [Int]$_.Group2DocumentIdCount; `
            $_.Group3DocumentIdCount = [Int]$_.Group3DocumentIdCount; `
            $_.Group4DocumentIdCount = [Int]$_.Group4DocumentIdCount; `
            $_.ViewStartTime = [Long]$_.ViewStartTime; `
            $_.QueryRowsElapsedTime = [Long]$_.QueryRowsElapsedTime; `
            $_.QueryRowsDirectoryTime = [Long]$_.QueryRowsDirectoryTime; `
            $_.QueryRowsLockTime = [Long]$_.QueryRowsLockTime; `
            $_.QueryRowsTimeInDatabase = [Long]$_.QueryRowsTimeInDatabase; `
            $_.QueryRowsDatabaseReadWaitTime = [Long]$_.QueryRowsDatabaseReadWaitTime; `
            $_.QueryRowsPagesPreread = [Int]$_.QueryRowsPagesPreread; `
            $_.QueryRowsPagesRead = [Int]$_.QueryRowsPagesRead; `
            $_.QueryRowsOffPageBlobHits = [Int]$_.QueryRowsOffPageBlobHits; `
            $_.OtherCount = [Int]$_.OtherCount; `
            $_.OtherElapsedTime = [Long]$_.OtherElapsedTime; `
            $_.OtherDirectoryTime = [Long]$_.OtherDirectoryTime; `
            $_.OtherLockTime = [Long]$_.OtherLockTime; `
            $_.OtherTimeInDatabase = [Long]$_.OtherTimeInDatabase; `
            $_.OtherDatabaseReadWaitTime = [Long]$_.OtherDatabaseReadWaitTime; `
            $_.OtherPagesPreread = [Int]$_.OtherPagesPreread; `
            $_.OtherPagesRead = [Int]$_.OtherPagesRead; `
            $_.OtherOffPageBlobHits = [Int]$_.OtherOffPageBlobHits; `
        } `
        elseif ($TraceType -eq 'InstantSearchBigFunnel') `
        {
            $_.TraceVersion = [Int]$_.TraceVersion; `
            $_.DatabaseGuid = [Guid]$_.DatabaseGuid; `
            $_.MailboxGuid = [Guid]$_.MailboxGuid; `
            $_.ErrorCode = [Int]$_.ErrorCode; `
            $_.QueryRowsInstance = [Int]$_.QueryRowsInstance; `
            $_.RowCount = [Int]$_.RowCount; `
            $_.FiltersChecked = [Int]$_.FiltersChecked; `
            $_.FiltersMatched = [Int]$_.FiltersMatched; `
            $_.PoisChecked = [Int]$_.PoisChecked; `
            $_.PoisMatched = [Int]$_.PoisMatched; `
            $_.PLDocuments = [Int]$_.PLDocuments; `
            $_.PLQueryTime = [Long]$_.PLQueryTime; `
            $_.TotalPagePreread = [Long]$_.TotalPagePreread; `
            $_.TotalPageRead = [Long]$_.TotalPageRead; `
            $_.TotalPageCacheMiss = [Long]$_.TotalPageCacheMiss; `
            $_.TotalDBWaitTime = [Long]$_.TotalDBWaitTime; `
            $_.PLPagePreread = [Long]$_.PLPagePreread; `
            $_.PLPageRead = [Long]$_.PLPageRead; `
            $_.PLPageCacheMiss = [Long]$_.PLPageCacheMiss; `
            $_.PLDBWaitTime = [Long]$_.PLDBWaitTime; `
            $_.FPPagePreread = [Long]$_.FPPagePreread; `
            $_.FPPageRead = [Long]$_.FPPageRead; `
            $_.FPPageCacheMiss = [Long]$_.FPPageCacheMiss; `
            $_.FPDBWaitTime = [Long]$_.FPDBWaitTime; `
        } `
        elseif ($TraceType -eq 'MailboxInfo') `
        {
            if ($MailboxInfoVersion -eq $null) `
            { `
                if ((Get-Member -InputObject $_ -Name 'NamedPropertyCount') -ne $null) `
                { `
                    Write-Warning "Units for RulesQuota is KB. Units for MaxSendSize, and MaxReceiveSize are MB. Units for MessageTotal, MessageFree, AttachmentTotal, AttachmentFree, OtherTotal, OtherFree, MailboxWarningQuota, MailboxShuttoffQuota, MailboxSendQuota, DumpsterWarningQuota, and DumpsterShutoffQuota are GB."
                    $MailboxInfoVersion = 5 `
                } `
                elseif ((Get-Member -InputObject $_ -Name 'CorruptionCheckCount') -ne $null) `
                { `
                    $MailboxInfoVersion = 4 `
                } `
                elseif ((Get-Member -InputObject $_ -Name 'MessageTotal') -ne $null) `
                { `
                    Write-Warning "Units for MessageTotal, MessageFree, AttachmentTotal, AttachmentFree, OtherTotal, and OtherFree are GB."
                    $MailboxInfoVersion = 3 `
                } `
                elseif ((Get-Member -InputObject $_ -Name 'SchemaVersion') -ne $null) `
                { `
                    $MailboxInfoVersion = 2 `
                } `
                else `
                { `
                    $MailboxInfoVersion = 1 `
                } `
            } `

            if ($MAILBOXSTATUS) `
            { `
                $_.Status = $MAILBOXSTATUS[$_.Status] `
            } `

            if ($MAILBOXTYPE) `
            { `
                $_.Type = $MAILBOXTYPE[$_.Type] `
            } `

            if ($MAILBOXTYPEDETAIL) `
            { `
                $_.Detail = $MAILBOXTYPEDETAIL[$_.Detail] `
            } `

            $_.DatabaseGuid = [Guid]$_.DatabaseGuid; `
            $_.MailboxGuid = [Guid]$_.MailboxGuid; `
            $_.DatabaseHash = [Int]$_.DatabaseHash; `
            $_.MailboxNumber = [Int]$_.MailboxNumber; `
            $_.IsArchiveMailbox = [Bool][Byte]$_.IsArchiveMailbox; `
            $_.IsSystemMailbox = [Bool][Byte]$_.IsSystemMailbox; `
            $_.IsHealthMailbox = [Bool][Byte]$_.IsHealthMailbox; `
            $_.MessageSize = [Long]$_.MessageSize; `
            $_.HiddenMessageSize = [Long]$_.HiddenMessageSize; `
            $_.DeletedSize = [Long]$_.DeletedSize; `
            $_.HiddenDeletedSize = [Long]$_.HiddenDeletedSize; `
            $_.MessageCount = [Long]$_.MessageCount; `
            $_.HiddenMessageCount = [Long]$_.HiddenMessageCount; `
            $_.DeletedCount = [Long]$_.DeletedCount; `
            $_.HiddenDeletedCount = [Long]$_.HiddenDeletedCount; `
            $_.FolderCount = [Long]$_.FolderCount; `

            if ($MailboxInfoVersion -gt 1) `
            {
                $_.SchemaVersion = [Int]$_.SchemaVersion; `
            } `

            if ($MailboxInfoVersion -gt 2) `
            {
                $_.MessageTotal = [Long]$_.MessageTotal * $DatabasePageSize / $MailboxSizeUnits; `
                $_.MessageFree = [Long]$_.MessageFree * $DatabasePageSize / $MailboxSizeUnits; `
                $_.AttachmentTotal = [Long]$_.AttachmentTotal * $DatabasePageSize / $MailboxSizeUnits; `
                $_.AttachmentFree = [Long]$_.AttachmentFree * $DatabasePageSize / $MailboxSizeUnits; `
                $_.OtherTotal = [Long]$_.OtherTotal * $DatabasePageSize / $MailboxSizeUnits; `
                $_.OtherFree = [Long]$_.OtherFree * $DatabasePageSize / $MailboxSizeUnits; `
                if (![String]::IsNullOrEmpty($_.OwnerGuid)) `
                {
                    $_.OwnerGuid = [Guid]$_.OwnerGuid; `
                } `
            } `

            if ($MailboxInfoVersion -gt 3) `
            {
                $_.CorruptionCheckCount = [Int]$_.CorruptionCheckCount; `
                $_.CorruptionCheckTime = [TimeSpan]::FromMilliseconds([Int]$_.CorruptionCheckTime); `
            } `

            if ($MailboxInfoVersion -gt 4) `
            {
                $_.NamedPropertyCount = [Int]$_.NamedPropertyCount; `
                $_.NextDocumentId = [Int]$_.NextDocumentId; `
                $_.ReplidCount = [Int]$_.ReplidCount; `
                $_.ReceiveFoldersCount = [Int]$_.ReceiveFoldersCount; `
                $_.LCID = [Int]$_.LCID; `
                $_.LocaleId = [Int]$_.LocaleId; `
                if (![String]::IsNullOrEmpty($_.OrganizationId)) `
                {
                    $_.OrganizationId = [Guid]$_.OrganizationId; `
                } `

                $_.MailboxWarningQuota = [Long]$_.MailboxWarningQuota / $MailboxSizeUnits; `
                $_.MailboxShutoffQuota = [Long]$_.MailboxShutoffQuota / $MailboxSizeUnits; `
                $_.MailboxSendQuota = [Long]$_.MailboxSendQuota / $MailboxSizeUnits; `
                $_.MailboxMessagesPerFolderCountWarningQuota = [Long]$_.MailboxMessagesPerFolderCountWarningQuota; `
                $_.MailboxMessagesPerFolderCountQuota = [Long]$_.MailboxMessagesPerFolderCountQuota; `
                $_.FolderCountWarningQuota = [Long]$_.FolderCountWarningQuota; `
                $_.FolderCountQuota = [Long]$_.FolderCountQuota; `
                $_.FolderHierarchyDepthWarningQuota = [Long]$_.FolderHierarchyDepthWarningQuota; `
                $_.FolderHierarchyDepthQuota = [Long]$_.FolderHierarchyDepthQuota; `
                $_.FolderChildrenCountWarningQuota = [Long]$_.FolderChildrenCountWarningQuota; `
                $_.FolderChildrenCountQuota = [Long]$_.FolderChildrenCountQuota; `
                $_.DumpsterWarningQuota = ([Long]$_.DumpsterWarningQuota) / $MailboxSizeUnits; `
                $_.DumpsterShutoffQuota = ([Long]$_.DumpsterShutoffQuota) / $MailboxSizeUnits; `
                $_.DumpsterMessagesPerFolderCountWarningQuota = [Long]$_.DumpsterMessagesPerFolderCountWarningQuota; `
                $_.DumpsterMessagesPerFolderCountQuota = [Long]$_.DumpsterMessagesPerFolderCountQuota; `
                $_.NamedPropertyCountQuota = [Long]$_.NamedPropertyCountQuota; `
                $_.RulesQuota = ([Int]$_.RulesQuota) / $RulesSizeUnits; `
                $_.MaxSendSize = ([Long]$_.MaxSendSize) / $MessageSizeUnits; `
                $_.MaxReceiveSize = ([Long]$_.MaxReceiveSize) / $MessageSizeUnits; `
                $_.IsPreviewPromoted = [Bool][Byte]$_.IsPreviewPromoted; `
                $_.SystemMessageWarningQuota = ([Long]$_.SystemMessageWarningQuota) / $MailboxSizeUnits; `
                $_.SystemMessageShutoffQuota = ([Long]$_.SystemMessageShutoffQuota) / $MailboxSizeUnits; `
                $_.SystemMessageSize = ([Long]$_.SystemMessageSize) / $MailboxSizeUnits; `
                $_.SystemMessageCount = [Long]$_.SystemMessageCount; `
                $_.LastLogonTime = [DateTime]$_.LastLogonTime; `
                $_.IsEncrypted = [Bool][Byte]$_.IsEncrypted; `
                $_.IsBigFunnelEnabled = [Bool][Byte]$_.IsBigFunnelEnabled; `
                $_.ScheduledISIntegLastFinished = [DateTime]$_.ScheduledISIntegLastFinished; `
            } `
        } `
        elseif ($TraceTypeFileName -eq 'ReferenceData') `
        { `
            $_.Key = [Int]$_.Key; `
        } `
        elseif ($TraceType -eq 'IntegrityCheckStatus') `
        {
                $_.MailboxGuid = [Guid]$_.MailboxGuid; `
                $_.RequestGuid = [Guid]$_.RequestGuid; `
                $_.JobGuid = [Guid]$_.JobGuid; `
                $_.DetectOnly = [Bool]$_.DetectOnly; `
                $_.CreationTime = [DateTime]$_.CreationTime; `
                $_.QueuedInAssistantTime = [DateTime]$_.QueuedInAssistantTime; `
                $_.StartTime = [DateTime]$_.StartTime; `
                $_.CompletedTime = [DateTime]$_.CompletedTime; `
                $_.LastExecutionTime = [DateTime]$_.LastExecutionTime; `
                $_.TimeInServerMs = [Long]$_.TimeInServerMs; `
                $_.CountBusyPreempted = [Int]$_.CountBusyPreempted; `
                $_.CorruptionsDetected = [Int]$_.CorruptionsDetected; `
                $_.CorruptionsFixed = [Int]$_.CorruptionsFixed; `
        } `
        elseif ($TraceType -eq 'IntegrityCheckCorruption') `
        {
                $_.MailboxGuid = [Guid]$_.MailboxGuid; `
                $_.RequestGuid = [Guid]$_.RequestGuid; `
                $_.JobGuid = [Guid]$_.JobGuid; `
                $_.DetectOnly = [Bool]$_.DetectOnly; `
                $_.Count = [Int]$_.Count; `
                $_.IsFixed = [Bool]$_.IsFixed; `
                $_.IsExternal = [Bool]$_.IsExternal; `
        } `

        if ($_) `
        { `
            if ($IncludeDetailTrace -and $DetailFiles) `
            { `
                $_ = (Get-TraceFromFiles -operation $_ -files $DetailFiles -TraceType $TraceType); `
            } `                    
            write-output $_; `
        } `
    }

write-host "$(get-date) Done importing $TraceType traces"


# SIG # Begin signature block
# MIIdpAYJKoZIhvcNAQcCoIIdlTCCHZECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUJlZOU3+PGilLzmJtvI/GgLaw
# MyOgghhkMIIEwzCCA6ugAwIBAgITMwAAAJvgdDfLPU2NLgAAAAAAmzANBgkqhkiG
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
# bWrJUnMTDXpQzTGCBKowggSmAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBvjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUsq3IWkK4iQXCh3+xKSTE1hW7MWswXgYKKwYB
# BAGCNwIBDDFQME6gJoAkAEcAZQB0AC0AUwB0AG8AcgBlAFQAcgBhAGMAZQAuAHAA
# cwAxoSSAImh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9leGNoYW5nZSAwDQYJKoZI
# hvcNAQEBBQAEggEAaxXxOworOTW/I/dreK9h7G2k63BE59YHnZ6whFzBIV5Yw075
# XP8ibcueg6YGd5gCw/HFqHtCzX5UrMVe92v85/KuvTZoOuhFtfGAAx+iDryoRE8s
# FgN6jVzQS2taXaZJ3+WQ3awDREYHImZtQ7i/awSpy/Wiom+rSPCloWdaYaYWl3cX
# /VliHxiSezSNtl+3ZZSfu7+t4YwShCGw3M2CtvovgJjJhzryOq9heGll+NIkq+xv
# TvLyRFpd5cmYqrwzTZxU9PgS3x8K8dkmvKYLrBxvmAKncK696p+IvgY2HVS71hXi
# TFoSqgg8FgFZ/oQxQrEsSRjpp5LKEBih5kmCbqGCAigwggIkBgkqhkiG9w0BCQYx
# ggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAACb4HQ3
# yz1NjS4AAAAAAJswCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0B
# BwEwHAYJKoZIhvcNAQkFMQ8XDTE2MDkwMzE4NDQwOFowIwYJKoZIhvcNAQkEMRYE
# FAgOCHaA0ikYBaL90ZEekBwM/mAjMA0GCSqGSIb3DQEBBQUABIIBAIHhzaeNcic1
# 55Mq021B5DDugS/8bvGZlCBw+3ZatMTWO/loFp0NrAAgOH1wsG+Em/oSr0cNfiHZ
# uriYUhV0Ob2EFCD9AZ/2mxnZCVWvUr2lAVdqb1IJolBoVriKxSiEPBmsrQvFuvPx
# +3aH3lqjQorl0/nwpac3X3ct1GTwYnIt6gCywzaaHGcOZEakTu15rd+kGp5FW4f3
# ib/VrF3HUpk3297iuVbV8OdMX8kz0Am3IjBU4yjTx/lXYeJiurNAIgiXb3Glwm2X
# ss892iJJbX2NqcR1jzHJgaqtKhTdShdObORmSwQh/VnThFkBmN2h2gChoIs1oy+l
# YRq4JEeTHYY=
# SIG # End signature block
