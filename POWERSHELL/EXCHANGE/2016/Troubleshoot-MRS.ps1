# Copyright (c) 2011 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

# Requires -Version 2

<#
    .Synopsis
    Can be used to attempt to complete specific failed move requests or used in a scan/fix mode that discovers move requests and fixes them.
    Notice: CA script only support -MoveRequest parameterset
    
    .Example TroubleShoot-MRS.ps1
    Scans all databases in the forest for move requests and attempts to repair or clean them up
    
    .Example TroubleShoot-MRS.ps1 -FileName <filename.csv>
    Attempts to repair all move requests in the file that are currently failed
    
    .Example TroubleShoot-MRS.ps1 -Database <database>
    Scans specified database for move requests and attempts to repair or clean them up
    
    .Example TroubleShoot-MRS.ps1 -Server <server>
    Scans all databases mounted on the server for move requests and attempts to repair or clean them up
    
    .Example TroubleShoot-MRS.ps1 -Organization <organization>
    Scans all move requests for given organization and attempts to repair or clean them up
    
    .Example TroubleShoot-MRS.ps1 -OrgList <filename.txt>
    Scans all move requests for given organizations in file and attempts to repair or clean them up
    
    .Example Troubleshoot-MRS.ps1 -RunningCA -MoveRequest -LogFolder \\ch1prd0310lg002\MailboxMoves\QhuTest -Target ch1prd0310ca002 
    Run this as CA script and scans all failed move requests for the current forest and attemps to repair or clean them up, using  ch1prd0310ca002 as Target machine and \\ch1prd0310lg002\MailboxMoves\TS-MRSLog to put logs
    
    .Example Troubleshoot-MRS.ps1 -RunningCA -MoveRequest -LogFolder \\ch1prd0310lg002\MailboxMoves\QhuTest -Target ch1prd0310ca002 -MoveRequestIdentities @("425a844f-7de0-4212-a1a9-3c083ced1074","c5721123-050f-4300-b800-340d603e9b91")
    Run this as CA script and scans 100 failed move request for the current forest and attemps to repair or clean them up, using  ch1prd0310ca002 as Target machine and \\ch1prd0310lg002\MailboxMoves\TS-MRSLog to put logs
    
    .Example Get-MoveRequest <identity> |TroubleShoot-MRS.ps1
    Pipelined single move request is attempted to be repaired or cleaned up
    
    .Example TroubleShoot-MRS.ps1 -MRGuid <guid> -MRDatabase <targetDatabase>
    Attempts to fix move request for specified user to targetdatabase
    
    .Parameter FileName
    CSV file that includes a UserGuid and MDB field.  UserGuid can be AD User or Mailbox Identity.  MDB field is target Mailbox Database.
    
    .Parameter Database
    Name of Database to scan
    
    .Parameter Server
    Scan local server
    
    .Parameter Organization
    Name of Organization to scan
    
    .Parameter OrgList
    File with one organization list per line.  Each organization is scanned.
    
    .Parameter MRGuid
    AD User or Mailbox Identity
    
    .Parameter MRDatabase
    Target mailbox database
    
    .Parameter LogFolder
    Location of logfiles.Default is current location
    
    .Parameter StatusLogFileName
    Name of logfile.  Default is location of excution\MRSMoveTroubleShooterStatus-yyyy-MM-dd.log
    
    .Parameter ErrorLogFileName
    Name of error logfile.  Default is location of excution\MRSMoveTroubleShooterError-yyyy-MM-dd.log
    
    .Parameter BatchSize
    Number of moves to simultaneously process - default 1000
    
    .Parameter Author
    E-mail address for any error reports
    
    .Parameter MRSMovers
    E-mail address for all status reports
    
    .Parameter Sleep
    Duration to sleep between scans of moves in progress
    
    .Parameter IgnoreNMR
    WARNING - This flag skips validation checks for existing move requests and will blindly force a move to the target database.  Verify your target is correct before using this flag.
    
    .Parameter Test
    This flag skips all write operations and reports them as success.  Useful mainly for validation in production where the operator does not have permissions to change move requests.
    
    .Parameter ADReplicationDelay
    In scan mode, any move request that have been touched within this duration are skipped.  Default is 2 hours.
    
    .Parameter QueueDelayAlert
    Length of time Move Request can stay in queue before generating an alert.  Default is 24 hours.
    
    .Parameter AgedCleanupObjectDelay
    Length of time before move request is removed from the queue if it is determined to be junk.  Default is 30 days.
    
    .Notes
    NAME: MRSMoveTroubleShooter
    AUTHOR: jamesweb@microsoft.com
    LASTEDIT: 5/23/2011 12:32:00
    KEYWORDS:
#>


[CmdletBinding(DefaultParameterSetName="Scan")]
param(

# CSV file that includes a UserGuid and MDB field.  UserGuid can be AD User or Mailbox Identity.  MDB field is target Mailbox Database.
[Parameter(Mandatory=$false, ParameterSetName="Filename")]
[ValidateNotNullOrEmpty()]
[string]$FileName,

# Name of Database to scan
[Parameter(Mandatory=$false, ParameterSetName="Database")]
[ValidateNotNullOrEmpty()]
[string]$database,

# Name of Server to scan
[Parameter(Mandatory=$false, ParameterSetName="Server")]
[ValidateNotNullOrEmpty()]
[switch]$server,

# Name of Organization to scan
[Parameter(Mandatory=$false, ParameterSetName="Organization")]
[ValidateNotNullOrEmpty()]
[string]$organization,

# File with one organization list per line.  Each organization is scanned.
[Parameter(Mandatory=$false, ParameterSetName="Organizations")]
[ValidateNotNullOrEmpty()]
[string]$orgList,

# Piped in Move request
[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="Pipeline")]
[ValidateNotNullOrEmpty()]
$mr,

# AD User or Mailbox Identity
[Parameter(Position=0, Mandatory=$false, ParameterSetName="ByGuid")]
[ValidateNotNullOrEmpty()]
[string]$MRGuid,

# Target mailbox database
[Parameter(Position=1, Mandatory=$false, ParameterSetName="ByGuid")]
[ValidateNotNullOrEmpty()]
[string]$MRDatabase,

# AD User or Mailbox Identity
[Parameter(Position=0, Mandatory=$false, ParameterSetName="Unlock")]
[ValidateNotNullOrEmpty()]
[string]$UnlockUserGuid,

# Scan via MR
[Parameter(Position=0, Mandatory=$false, ParameterSetName="MoveRequest")]
[ValidateNotNullOrEmpty()]
[switch]$MoveRequest,

# The list of move requests to go over. If none are provided, all the failed ones on local forest are scanned
[Parameter(Position=1, Mandatory=$false, ParameterSetName="MoveRequest")]
[ValidateNotNullOrEmpty()]
[string[]]$MoveRequestIdentities = @(),

# Name of logfile.
[Parameter(Mandatory=$false)]
[ValidateNotNullOrEmpty()]
[string]$StatusLogFileName="MRSMoveTroubleShooterStatus-$(get-date -f 'yyyy-MM-dd-HH-mm').log",

# Name of error logfile.
[Parameter(Mandatory=$false)]
[ValidateNotNullOrEmpty()]
[string]$ErrorLogFileName="MRSMoveTroubleShooterError-$(get-date -f 'yyyy-MM-dd-HH-mm').log",

# Location of logfiles.
[Parameter(Mandatory=$false)]
[ValidateNotNullOrEmpty()]
[string]$LogFolder=$(get-location),

# Number of moves to simultaneously process
[Parameter(Mandatory=$false)]
[ValidateRange(1, 10000)]
[int]$BatchSize=1000,

# E-mail address for any error reports
[Parameter(Mandatory=$false)]
[ValidatePattern("\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b")]
[string]$author="exmigtst@microsoft.com",

# E-mail address for all status reports
[Parameter(Mandatory=$false)]
[ValidatePattern("\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b")]
[string]$MRSMovers="exchdcmoves@microsoft.com",

# Duration to sleep in seconds between scans of moves in progress
[Parameter(Mandatory=$false)]
[ValidateRange(0, 480)]
[int]$Sleep=30,

# WARNING - This flag skips validation checks for existing move requests and will blindly force a move to the target database.  Verify your target is correct before using this flag.
[Parameter(Mandatory=$false)]
[ValidateNotNullOrEmpty()]
[switch]$IgnoreNMR,

# This flag skips all write operations and reports them as success.  Useful mainly for validation in production where the operator does not have permissions to change move requests.
[Parameter(Mandatory=$false)]
[ValidateNotNullOrEmpty()]
[switch]$Test,

# Running under Monitorying context or not
[Parameter(Mandatory=$false)]
[ValidateNotNullOrEmpty()]
[switch]$MonitoringContext,

# Running under CA context or not
[Parameter(Mandatory=$false)]
[ValidateNotNullOrEmpty()]
[switch]$RunningCA,

# Length of time in Hours an objects have been last touched before an action is taken (reduces replication issues)
[Parameter(Mandatory=$false)]
[ValidateRange(0, 24)]
[int]$ADReplicationDelay=2,

# Length of time in Hours a Move Request can stay in queue before generating an alert.
[Parameter(Mandatory=$false)]
[ValidateRange(0, 720)]
[int]$QueueDelayAlert=24,

# Length of time in Days before move request is removed from the queue if it is determined to be junk.
[Parameter(Mandatory=$false)]
[ValidateRange(0, 365)]
[int]$AgedCleanupObjectDelay=30,

#Max bad items we allowed when trying to skip bad items
[Parameter(Mandatory=$false)]
[int]$maxBadItems=5,

#Allow change reg key or not
[Parameter(Mandatory=$false)]
[switch]$allowKeyChange
)

#########################################################
#                                                                                                                           #
#                                                     Script Functions                                               #
#                                                                                                                           #
#########################################################

begin
{
    # version
    $script:version = "1.06"
    # initialize
    if(! [System.IO.Path]::IsPathRooted($StatusLogFileName))
    {
        $StatusLogFileName="{0}\{1}" -f $LogFolder, $StatusLogFileName
    }
    if(! [System.IO.Path]::IsPathRooted($ErrorLogFileName))
    {
        $ErrorLogFileName="{0}\{1}" -f $LogFolder, $ErrorLogFileName
    }
    $fm = [System.IO.FileMode]::Create
    $fmo = [System.IO.FileMode]::Open
    $fa = [System.IO.FileAccess]::ReadWrite
    $far = [System.IO.FileAccess]::Read
    $fs = [System.IO.FileShare]::Read
    $count = 0
    $status = @{}
    $start = (Get-Date)
    
    $script:list = New-Object System.Collections.ArrayList
    $script:ADlist = New-Object System.Collections.ArrayList
    $script:removelist = New-Object System.Collections.ArrayList
    $script:errorcount = 0
    $script:total = 0
    $script:cleanup = 0
    $script:processing = 0
    $script:unclassified = 0
    $script:ADOrphan = 0
    $script:LastUpdateNull = 0
    $script:skippedList = @{}    
    $maxRetryCount = 5
    
    if($RunningCA)
    {
        $maxRetryCount = 1
    }
    
    $exchangeInstallPath=(get-itemproperty hklm:\software\microsoft\exchangeServer\v14\setup).MsiInstallPath
    
    # Load Datacenter Common Libraries
    if ($MonitoringContext)
    {
        $scriptPath = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
        . "$scriptPath\CITSLibrary.ps1"
        . "$scriptPath\DiagnosticScriptCommonLibrary.ps1"
    }
    else
    {
        . "$exchangeInstallPath\Datacenter\DatacenterHealthCommonLibrary.ps1"
        . "$exchangeInstallPath\Scripts\CITSLibrary.ps1"
        . "$exchangeInstallPath\Scripts\DiagnosticScriptCommonLibrary.ps1"
    }

    Load-ExchangeSnapin

    # Event log source name for application log
    $appLogSourceName = "Mailbox Replication Troubleshooter"

    # Event log source name for crimson log
    $crimsonLogSourceName = "Mailbox Replication"

    # The Arguments object is needed for logging events
    $Arguments = Validate-Arguments `
        -Server "$env:ComputerName" `
        -Database $database `
        -Action $PSCmdlet.ParameterSetName `
        -MonitoringContext:$MonitoringContext

    # Event log entry dictionary
    #
    $MRSLogEntries = @{
        #
        # Events logged to application log and windows event (crimson) log
        # Information: 5000-5299; Warning: 5300-5599; Error: 5600-5999;
        #
        #   Informational events
        #
        TSStarted=(5000,"Information","The Troubleshooter started successfully.")
        TSSuccess=(5001,"Information", "The troubleshooter finished successfully.")
        TSSummary=(5002,"Information", "%1")
        TSInfo=(5003,"Information", "%1")
        #
        #   Warning events
        #
        TSWarning=(5300,"Warning", "%1")
        #
        #   Error events
        #
        TSFailed=(5600,"Error", "The troubleshooter failed with exception %1.")
        TSError =(5601,"Error", "%1")
        #
        #   Events logged only to crimson (windows) event log
        #   Information: 6000-6299; Warning: 6300-6599; Error: 6600-5999;
        #
        #
        #   Informational events
        #

        #
        #   Warning events
        #
        TSCrimsonWarning =(6300,"Warning", "%1")
        
        #
        #   Error events
        #


    }


    # mail notifications
    $header = "MRS TroubleShooter notification from $env:ComputerName`r`n`r`nDetails:`r`n"
    $from = "MRSTS@$env:UserDNSDomain"

$code =  @"
namespace Internal.Exchange.MoveQueueStat
{
    using System;
    using System.Collections.Generic;    
    using System.Linq;
    using System.Text;

   
    public class MoveUnhealthyStat
    {
        public String statType;
        public int count;
        public MoveUnhealthyStat(String statType)
        {
            this.statType = statType;
            this.count = 1;
        }
    }
    
    public class MoveQueueStat
    {
        public int queueDelayedCount;        
        public int relinquishedDelyedCount;        
        public List<MoveUnhealthyStat> stalledStat;
        public int failedCount;
        
        public MoveQueueStat()
        {
            this.queueDelayedCount = 0;
            this.relinquishedDelyedCount = 0;
            this.failedCount = 0;
        }                

        public void InsertStalledStat(string stalledType)
        {
            if (this.stalledStat == null)
            {
                this.stalledStat = new List<MoveUnhealthyStat>();
            }
            
            foreach(MoveUnhealthyStat stat in this.stalledStat)
            {
                if(String.Compare(stat.statType,stalledType) ==0)
                {
                    stat.count++;
                    return;
                }
            }
            this.stalledStat.Add(new MoveUnhealthyStat(stalledType));
        }
    }
}
"@
    Add-Type -Language CSharpVersion3 -TypeDefinition $code         
    
    $script:QueueStatMap = @{}
    $script:stalledTypeMap = @{}    
    $script:queuedMoveCount = 0
    $script:relinquishedMoveCount = 0
    $script:stalledMoveCount = 0
    $script:stalledQueued = 0
    $script:failedMoveCount = 0
    
    function RemoveOldLogs
    {
        param ($TargetFolder, $days)
        $LastWrite = (get-date).AddDays(-$days)
        $files = @(get-childitem $TargetFolder * | Where {$_.LastWriteTime -le "$LastWrite"})
        foreach ($file in $files)
        {
            remove-Item $file.FullName
        }
    }
        
    function DumpQueueStatMap ($sendAlert)
    {
        $summary = "Unhealthy Cross Org Moves Summary`r`n"
        $summary += "`r`nQueued moves over 4 hours count is " + $script:queuedMoveCount
        $summary += "`r`nRelinquished moves over 6 hours count is " + $script:relinquishedMoveCount
        $summary += "`r`nFailed moves in last 2 hours count is " +  $script:failedMoveCount
        $summary += "`r`n{0} moves are stalled, they distributed in {1} different databases `r`n" -f $script:stalledMoveCount, $script:stalledQueued
        
        foreach($key in @($script:stalledTypeMap.keys))
        {
           $summary += "{0} count {1} `r`n " -f $key, $script:stalledTypeMap[$key]
        }        
        
        
        
        $summary += "`r`n`r`n ******DETAILS FOR RELATED DB******`r`n`r`n"
       
        $summary += $script:crossorgmsg        
        
        if($sendAlert)
        {
            $title = "MRST: {0}: Unhealthy cross org moves notification from {1}" -f $script:version, $env:ComputerName
            TSNotify $summary "MailAlert" $MRSLogEntries.TSCrimsonWarning $script:queueStatsLog $title
        }
        
        $summaryFile = "$LogDirectory\QueueStats-Summary-{0}.log" -f (get-date).ToString("yyyyMMdd_hh")        
        $summary | out-file $summaryFile        
    }
    
    function QueueStatMapAnalysis
    {      
        
        foreach($key in @($script:QueueStatMap.keys))
        {
            [Internal.Exchange.MoveQueueStat.MoveQueueStat]$queueStat =  $script:QueueStatMap[$key]
            $script:queuedMoveCount += $queueStat.queueDelayedCount            
            $script:relinquishedMoveCount += $queueStat.relinquishedDelyedCount
            $script:failedMoveCount += $queueStat.failedCount
            
            if($queueStat.stalledStat -ne $null)
            {
                $script:stalledQueued++                
            
                foreach($stalledStatItem  in @($queueStat.stalledStat))
                {
                    $script:stalledMoveCount += $stalledStatItem.count
                    $type = $stalledStatItem.statType
                    if($script:stalledTypeMap[$type] -eq $null)
                    {
                        $script:stalledTypeMap[$type] = $stalledStatItem.count
                    }
                    else
                    {
                        $script:stalledTypeMap[$type] = [int32]::parse($script:stalledTypeMap[$type]) + $stalledStatItem.count
                    }
                }
            }
        }
        
        if($script:queuedMoveCount -gt 0)
        {
            return $true
        }
        if($script:relinquishedMoveCount -gt 0)
        {
            return $true
        }
        if($script:stalledQueued -gt 3)
        {
            return $true
        }
        if($script:failedMoveCount -gt 15)
        {
            return $true
        }
        
        foreach($key in @($script:stalledTypeMap.keys))
        {
           if($script:stalledTypeMap[$key] -gt 5)
           {
                return $true
           }
        }
        
        return $false
    }
    # Helper for formating notification messages
    function TSNotify
    {
        param ([ValidateNotNull()]$message, $type, $eventInfo, $attachments, $title="" )
        
        if($title -eq "" -or $title -eq $null)
        {
            $title = "MRST: {0}: {1} notification from {2}" -f $script:version, $type, $env:ComputerName
        }
        $body = "{0}{1}" -f $header, $message
        
        if ($MonitoringContext)
        {
            if ($body.Length -gt 32000) { $body = $body.Substring(0,32000) }
            try
            {
                Log-Event -Arguments $Arguments -EventInfo $eventInfo -Parameters @($body)
            }
            catch
            {
                [string]$message =  "Unhandled Exception in Log-Event`r`n"
                $message += "EventInfo: {0}, Body Length: {1}" -f $eventInfo, $body.Length
                $message +=  "Error text: {0}`r`n" -f ($_ |fl |out-string)
                $message +=  "Exception: {0}`r`n" -f $_.Exception.ToString()
                WriteError ($message)
            }
        }
        else
        {
            switch ($type)
            {
                "MailAlert"
                {
                    send-mail -from $from -tos $MRSMovers -title $title -body $body -attachments $attachments
                }
                "PageAlert"
                {
                    #send-supportmail - need to figure this out for moves that can't be fixed, i.e. locked mailboxes
                    send-mail -from $from -tos $MRSMovers -title $title -body $body  -attachments $attachments
                }
                default
                {
                    send-mail -from $from -tos $author -title $title -body $body  -attachments $attachments
                }
            }
        }
    }    
    
    function TellMatch($s, $key)
    {            
        if($s.message.toString() -match $key)
        {
            return $true
        }
        else
        {
            foreach($failure in $s.Report.Failures)
            {
                if($failure -match $key)
                {
                    return $true
                }
            }
        }
        
        return $false        
    }    
    
    # Classify the failure type
    # $obj can MoveRequestStatistics object with Report or just object with User Guid (generates report)
    function ClassifyFail ($obj)
    {
        Write-Verbose "Enter ClassifyFail"              

        [string]$err = ""
        if ((Get-Member -InputObject $obj -Name Report) -eq $null)
        {
            $obj = Get-MoveRequestStatistics $obj.UserGuid -IncludeReport -ErrorAction SilentlyContinue
        }
        #extract ErrorMessage
        if ($obj.Report -match "<errorMessage>(?<content>.*)") { [string]$err = $matches['content'] }
        if ($err -match "Message \(size .*\) exceeds the maximum allowed size") { return "MaxSize" }
        Write-Verbose $obj.Message
        
        if (TellMatch $obj "This mailbox exceeded the maximum number of large items that were specified") { return "MaxSize" }# new log MaxSize
        if (TellMatch $obj ".*: Unable to SetSearchCriteria.") { return "SetSearchCriteria" }
        
        if (TellMatch $obj "tombstone is present in the target database") { return "Tombstone" }
        if (TellMatch $obj "Database .* doesn't satisfy .*constraint CISecond") { return "CIThrottle" }
        if (TellMatch $obj "Database .* doesn't satisfy .*constraint Second") { return "HAThrottle" }        
        if (TellMatch $obj "Database .* does not satisfy .*constraint CISecond") { return "CIThrottle" }
        if (TellMatch $obj "Database .* does not satisfy .*constraint Second") { return "HAThrottle" }
        if (TellMatch $obj "The request has been temporarily postponed because a database has failed over") { return "HAThrottle" }
        if (TellMatch $obj "Couldn't switch the mailbox into Sync Source mode.") { return "SyncSource" }
        if (TellMatch $obj "Lid: 25000") { return "NonCanonicalACL" }
        if (TellMatch $obj "Lid: 24952") { return "NonCanonicalACL" }
        if (TellMatch $obj "Lid: 24936") { return "NonCanonicalACL" }
        if (TellMatch $obj "Lid: 25092") { return "NonCanonicalACL" }        
        if (TellMatch $obj "MapiExceptionPartialCompletion: Unable to copy to target") { return "CopyTarget" }
        if (TellMatch $obj "MapiExceptionCallFailed: Unable to copy to target") { return "CopyTarget" }
        if (TellMatch $obj "Active directory response: The LDAP server is unavailable.") { return "ADUnavailable" }        
        if (TellMatch $obj "Unable to make connection to the server") { return "NetworkError" }       
        if (TellMatch $obj "MapiExceptionRpcServerTooBusy: Unable to make connection to the server") { return "RpcServerTooBusy" }
        if (TellMatch $obj "MapiExceptionShutoffQuotaExceeded") { return "QuotaExceeded" }
        if (TellMatch $obj "MapiExceptionWrongServer: Unable to open mailbox") { return "WrongServer" }
        if (TellMatch $obj "Error: Active Directory operation failed on") { return "ADFailed" }
        if (TellMatch $obj "MapiExceptionCallFailed: Unable to save changes") { return "Resubmit" }
        if (TellMatch $obj "MapiExceptionPartialCompletion: Unable to delete folder") { return "Resubmit" }
        if (TellMatch $obj "The active server for database .* could not be found") { return "Resubmit" } 
        if (TellMatch $obj "An error occurred while updating a user object after the move .* The domain controller .* has the Maintenance Mode flag on") { return "Resubmit" } 
        if (TellMatch $obj "An error occurred while updating a user object after the move .* The domain controller .* is not available for use at the moment") { return "Resubmit" }         
        if (TellMatch $obj "Unable to read mailbox signature basic info") { return "SignatureError" }    
        if (TellMatch $obj "This mailbox exceeded the maximum number of corrupted items that were specified for this request") 
        {
            if($obj.report.baditems.count -gt $maxBaditems)
            {
                return "BadItem"            
            }
            else
            {
                #bug 544705, #bug 535441
                if ((TellMatch $obj "StoreEc: 0x471")  -or (TellMatch $obj "Lid: 40691"))
                {
                    return "IncreaseBIL"
                }
                else
                {
                    return "BadItem"    
                }
            }
        }
        #bug 536294
        if (TellMatch $obj "MapiExceptionBadValue: Unable to synchronize manifest")
        {
            return "badOOFResubmit"
        }
        #bug 536294
        if (TellMatch $obj "MapiExceptionNotFound: Unable to synchronize manifest")
        {
            return "fixMBXResubmit"
        }        
        if(TellMatch $obj "MapiExceptionNotFound: Unable to get view information")
        {
            #no solution yet
            return "bug532524"
        }
        if(TellMatch $obj "MapiExceptionNotFound: Unable to query table rows")
        {        
            if($obj.Report.Failures[0].datacontext -match "Archive")
            {
            	return "skipAchiverule"
            }
            else
            {   
                #need manual check to see whehter it is safe to skip thse rules
                return "PrimaryruleCorrupted"
            } 
        }
    
        # Moves that need to be removed and resubmitted - Experienced in R4->R5, disabled for now.
        #if (TellMatch $obj "Error: Unexpected error 0x8004010F") { return "Resubmit" }
        #if (TellMatch $obj "The source mailbox is in the wrong mailbox database") { return "Resubmit" }
        if($obj.Status -eq "Failed")
        {
            return "Retry"
        }
        else
        {
            return "None"
        }
    }
    

    function WriteLog
    {
        param ([ValidateNotNull()]$mbo, $message)
        
        $output = "`"{0}`", `"{1}`", {2}, {3}, {4}, {5}" -f $message, $mbo.UserGuid, $mbo.MDB, (get-date), $mbo.FailClass, $mbo.RetryCount
        WriteLine $output "Status"
        Write-Verbose $output        
    }

    function WriteLogEx
    {
        param ([ValidateNotNull()]$message, $UserGuid, $MDB, $FailClass)
        $output = "`"{0}`", `"{1}`", {2}, {3}, {4}, {5}" -f $message, $UserGuid, $MDB, (get-date), $FailClass, 0
        WriteLine $output "Status"
        Write-Verbose $output        
    }

    function WriteError($data)
    {
        Write-Verbose $data        
        WriteLine $data "Error"
        TSNotify $data "Exception" $MRSLogEntries.TSFailed
    }
    
    function WriteLine
    {
        param ([ValidateNotNull()]$message, $type)

        # No logging in SCOM
        if ($MonitoringContext) { return }
        
        [byte[]] $messageByte = [System.Text.Encoding]::ascii.getbytes("{0}`r`n" -f $message)
        switch ($type)
        {
            "Status"
            {
                $fswrite.Write($messageByte, 0, $messageByte.Length)
                $fswrite.Flush()            
            }
            "Error"
            {
                $fewrite.Write($messageByte, 0, $messageByte.Length)
                $fewrite.Flush()
            }
            default
            {
                Write-Error "Invalid Log Type - $type"
            }
        }
    }
    
    function ScanMDB([string]$mdb)
    {
        Write-Verbose "Enter ScanMDB: $mdb"
        Write-Progress -Activity " " -Status "Requesting Move Statistics..." -id 1
        
        $stats = @(Get-MoveRequestStatistics -MoveRequestQueue $mdb -includereport -ErrorAction SilentlyContinue) 
        [Internal.Exchange.MoveQueueStat.MoveQueueStat]$queueStat = New-Object Internal.Exchange.MoveQueueStat.MoveQueueStat
        $count = 0
        $script:relinquishCount = 0
        $script:queueDelayCount = 0                
              
        foreach ($obj in $stats)
        {
            if ($obj.RequestStyle -eq "CrossOrg")
            {                
                if ($obj.StatusDetail.Tostring().startswith("Stalled"))
                {
                    $queueStat.InsertStalledStat($obj.StatusDetail);
                }
                
                if ($obj.StatusDetail -eq "Queued") 
                { 
                    $queuedHours = ($(get-date) - $obj.lastupdatetimestamp).TotalHours
                    if ($queuedHours  -gt 4) 
                    {                        
                        $queueStat.queueDelayedCount++;
                    }   
                }
                
                # Check for Relinquished jobs older than an 6 hours
                if ($obj.StatusDetail -eq "Relinquished")
                {
                    $relinquishedDuration =$obj.OverallDuration
                    
                    if($obj.TotalInProgressDuration -ne $null) 
                    {
                        $relinquishedDuration -= $obj.TotalInProgressDuration
                    }
                    if($obj.TotalQueuedDuration -ne $null) 
                    {
                        $relinquishedDuration -= $obj.TotalQueuedDuration
                    }                    
                    if($obj.TotalDataReplicationWaitDuration -ne $null) 
                    {
                        $relinquishedDuration -= $obj.TotalDataReplicationWaitDuration
                    }                    
                    if($obj.TotalSuspendedDuration -ne $null) 
                    {
                        $relinquishedDuration -= $obj.TotalSuspendedDuration
                    } 
                    if ($relinquishedDuration.TotalHours -gt 6)
                    {
                        $queueStat.relinquishedDelyedCount++;
                    }
                }
                
                if ($obj.StatusDetail.Tostring().startswith("Failed") -and ($obj.LastUpdateTimeStamp.AddHours(2)) -ge (Get-Date)) 
                {   
                    $queueStat.failedCount++;
                }                
            } 
       }
       
       
       $script:QueueStatMap[$mdb] = $queueStat 
       
       $crossOrgStats = @($stats |?{$_.RequestStyle -eq "CrossOrg"}|group statusdetail |select Name, Count)       
       $intraOrgStats = @($stats |?{$_.RequestStyle -eq "IntraOrg"}|group statusdetail |select Name, Count)       
       
       
       if($crossOrgStats.count -gt 0)
       {           
           $crossOrgMsgTitle = "`r`n======================`r`n"
           $crossOrgMsgTitle += "`r`n{0}: `r`n" -f $mdb
           $crossOrgMsg = ""
           
           foreach($stat in $crossOrgStats)
           {
               "{0}, {1}, {2}, {3}" -f $mdb, "CrossOrg", $stat.Name, $stat.Count | out-file $script:queueStatsLog  -append -Encoding "ASCII"
               if($stat.Name -eq "Queued" -and $queueStat.queueDelayedCount -gt 0)
               {
                   $unhealthymovemsg = "`r`n{0} out of {1} queued moves queued more than 4 hours" -f $queueStat.queueDelayedCount, $stat.Count                   
                  "{0}, {1}, {2}, {3}" -f $mdb, "CrossOrg", "Queued4hours", $queueStat.queueDelayedCount| out-file $script:queueStatsLog  -append -Encoding "ASCII"                   
               }
               if($stat.Name -eq "Relinquished" -and $queueStat.relinquishedDelyedCount -gt 0)
               {
                   $unhealthymovemsg = "`r`n{0} out of {1} relinquished moves relinquished more than 6 hours" -f $queueStat.relinquishedDelyedCount, $stat.Count
                  "{0}, {1}, {2}, {3}" -f $mdb, "CrossOrg", "Relinquished6hours", $queueStat.relinquishedDelyedCount| out-file $script:queueStatsLog  -append -Encoding "ASCII"                   
               }
               if($stat.Name.StartsWith("Failed") -and $queueStat.failedCount -gt 0)
               {
                   $unhealthymovemsg = "`r`n{0} out of {1} failed moves failed in last 2 hours" -f $queueStat.failedCount, $stat.Count
                   "{0}, {1}, {2}, {3}" -f $mdb, "CrossOrg", "FailedInLast2Hours", $queueStat.failedCount| out-file $script:queueStatsLog  -append -Encoding "ASCII"
               }                   
               if($stat.Name.StartsWith("Stalled"))
               {
                   $unhealthymovemsg = "`r`n{0} count {1}" -f $stat.Name, $stat.Count               
               }               
               
               $crossOrgMsg += $unhealthymovemsg
               $unhealthymovemsg = ""
           }    
           
           
           if($crossOrgMsg -ne "")
           {
               $crossOrgMsg = $crossOrgMsgTitle + $crossOrgMsg
               $script:crossOrgMsg += $crossOrgMsg
           }           
       }
       
       
       if($intraOrgStats.count -gt 0)
       {           
           foreach($stat in $intraOrgStats)
           {
               "{0}, {1}, {2}, {3}" -f $mdb, "IntraOrg", $stat.Name, $stat.Count | out-file $script:queueStatsLog -append -Encoding "ASCII"
           }
       }
       
       foreach ($obj in $stats)
       {
            if ($obj.Status -eq "InProgress" -and $obj.priority -eq "Lowest" -and $requestStyle -eq "IntraOrg" -and $obj.TotalInProgressDuration -gt (New-Timespan -hours 1))
            {
                Suspend-moverequest $obj.MailboxIdentity -confirm:$false
                sleep 30
                Resume-moverequest $obj.MailboxIdentity -confirm:$false
            }
            if ($obj.LastUpdateTimeStamp -eq $null)
            {
                $script:LastUpdateNull++
                Continue
            }
                
            # Ignore everything that was touched in last X hours (default 2) - use 0 to skip check for testing
            if (($obj.LastUpdateTimeStamp.AddHours($ADReplicationDelay)) -ge (Get-Date)) 
            {
                $script:processing++
                Continue
            }

            if ($stats -is [System.Array])
            {
                $count++
                $complete = [System.Math]::Round($count/$stats.Count * 100, 1)
                Write-Progress -Activity " " -Status ("Check for Orphans - $count out of " + $stats.Count) -percentcomplete $complete -id 1 -CurrentOperation "$complete% complete"
            }
            ProcessMRS $obj            
        }
        
        if ($script:relinquishCount -gt 0)
        {
            TSNotify ("[Requires Investigation] {0} Reqlinquished jobs detected on Database {1}" -f $script:relinquishCount, $mdb) "MailAlert" $MRSLogEntries.TSCrimsonWarning
        }
        if ($script:queueDelayCount -gt 0)
        {
            TSNotify ("[Requires Investigation] {0} moves in queued state > {1} hours detected on database {2}" -f $script:queueDelayCount, $QueueDelayAlert, $mdb) "MailAlert" $MRSLogEntries.TSCrimsonWarning
        }          
    }

    function ProcessMRS($obj)
    {
            Write-verbose ("Scanning MRS for {0}" -f $obj.MailboxIdentity)
            
            # Clear vars
            $failClass = "None"
            $err = ""
            $global:ex = ""
            $requestStyle = ""
            $userID = ""

            # check if R4 or R5 move
            if ($obj.RequestStyle -eq $null)
            {
                $requestStyle = $obj.MoveType
                $userID = $obj.UserIdentity
            }
            else
            {
                $requestStyle = $obj.RequestStyle
                $userID = $obj.MailboxIdentity
            }           

            #Default GUID to userID in event GUID can't be located
            $UserGUID = $userID
            $mdb = $obj.TargetDatabase

            #Check if object exitsts in AD
            $gmr = get-moverequest $userID -ErrorVariable global:ex -ErrorAction SilentlyContinue
            $orphan = ($gmr -eq $null)
            if (! $orphan) 
            { 
                $UserGUID = $gmr.GUID 
                # Check MR and MRS are in sync (otherwise orphan)
                write-verbose "Check MR/MRS: $($gmr.TargetDatabase) == $($obj.TargetDatabase)"
                
                $orphan = ($gmr.TargetDatabase -ne $obj.TargetDatabase )
                
                # Check for Relinquished jobs older than an hour
                if (!$orphan -and $obj.StatusDetail -eq "Relinquished" -and ($obj.OverallDuration - $obj.TotalInProgressDuration - $obj.TotalQueuedDuration) -gt (New-Timespan -hours 2) -and ($obj.Message -match "because a database has failed over"))
                {
                    $script:relinquishCount++
                }
            }

            #Skip not Orphan and Queued
            if ((! $orphan) -and ($obj.Status -eq "Queued")) 
            { 
                if (($obj.LastUpdateTimeStamp.AddHours($QueueDelayAlert)) -ge (Get-Date)) 
                {
                    $script:processing++
                    Continue 
                }
                # Request queued for > (default 24)  $AgedObjectDelay hours, need to alert
                $script:queueDelayCount++
            }

            # check if mailbox present
            $mbx = Get-Mailbox -Identity $userID -ErrorAction SilentlyContinue -ResultSize 1
            if ($mbx -eq $null) 
            { 
                $mbx = Get-Mailbox -Identity $userID -arbitration -ErrorAction SilentlyContinue -ResultSize 1
                if ($mbx -eq $null) 
                { 
                    $mbx = Get-Mailbox -Identity $userID -archive -ErrorAction SilentlyContinue -ResultSize 1
                }
            }
            $mailboxPresent = $mbx -ne $null

            # Try to get AD Guid for orphans, otherwise default to UserIdentity (which might contain non-ascii)
            if ($orphan -and $mailboxPresent) { $UserGUID = $mbx.Guid }
            
            #Check Conectivity
            $canConnect = ($(Test-MAPIConnectivity $userID -ErrorVariable +global:ex -ErrorAction SilentlyContinue).Result.Value -eq "Success")
            
            if ($obj.Status -eq "Failed" -and $mailboxpresent)
            {
                #Set Error type
                $failClass = ClassifyFail ($obj)
            }
            if ($obj.Status -eq "Completed" -and $requestStyle -eq "IntraOrg")
            {
                if ($canConnect)
                {
                    #AD object missing, but mailbox can connect so object is just junk in queue.
                    if ($orphan)
                    {
                        if (!$Test) { remove-moverequest -MoveRequestQueue $mdb -MailboxGuid $obj.ExchangeGuid -confirm:$false }
                        WriteLogEx "Remove-MoveRequest Queue Orphan" $UserGuid $MDB "CompleteQueueOrphan"
                        $script:cleanup++
                        continue
                    }                    
                }
                #Mailbox that was deleted while a move request was still in completed state is safe to remove - they should all be orphans
                elseif (!$mailboxPresent -and $orphan)
                {
                    if (!$Test) { remove-moverequest -MoveRequestQueue $mdb -MailboxGuid $obj.ExchangeGuid -confirm:$false }
                    WriteLogEx "Remove-MoveRequest No Mailbox Orphan" $UserGuid $MDB "CompleteNoMailboxOrphan"
                    $script:cleanup++
                    continue
                }
            }

            #Remove objects of certain types older than $AgedCleanupObjectDelay
            $IsOld = ($obj.LastUpdateTimeStamp.AddDays($AgedCleanupObjectDelay)) -le (Get-Date);
            if ($IsOld)
            {
                #IntraOrg orphan no mailbox
                if ($Orphan -and !$MailboxPresent -and $requestStyle -eq "IntraOrg")
                {
                    if (!$Test) { remove-moverequest -MoveRequestQueue $mdb -MailboxGuid $obj.ExchangeGuid -confirm:$false }
                    WriteLogEx "Remove-MoveRequest Queue Orphan" $UserGuid $MDB "NoMailboxQueueOrphan"
                    $script:cleanup++
                    continue
                }
                # Others TBD - i.e. any where mailboxpresent -eq false, and orphan -eq true (currently just doing that for completed)
                # Maybe InterOrg, CrossOrg
            }
            
            if($obj.Status -eq "CompletedWithWarning" -and !$orphan -and $requestStyle -eq "IntraOrg")            
            {     
                if($mailboxpresent -and $mbx.Database -eq $obj.TargetDatabase)
                {
                    if($canConnect)
                    {    
                         if (!$Test)
                         {
                            remove-moverequest $userGUID -confirm:$false
                         }
                         WriteLogEX "Remove completedwithwarning requests that canConnect and homeMDB equals targetDB" $UserGuid $MDB "Completed requests although with warnings"                                                    
                         $script:cleanup++
                         continue
                    }
                    
                    #else TargetLocked, which has been handled by afterwards code                    
                }
                else
                {
                    $failClass = "Resubmit"
                    WriteLogEx "CompletedWithWarning move requests need resubmit, as its homeMDB != targetDB or mailbox is not present yet"                
                }            
            }            

            # Mailbox exists on Target & Target Locked 
            # Orphan in any state indicates Lossy Failover condition
            # CompletedWithWarning indicates transient error
            # New -autosuspend, remove - Need eventlog to store previous Move Report            
            if(! $canConnect -and $mailboxpresent -and $mbx.Database -eq $obj.TargetDatabase -and ($orphan -or $obj.Status -eq "CompletedWithWarning"))
            {
                #Need to eventlog report before continuing, or send e-mail after failing
                TSNotify ("[Notification] Locked mailbox detecting - attempting to repair user: {0}`r`n`r`nPrevious Move Request:`r`n{1}" -f $UserGuid, ($obj|fl|Out-String )) "MailAlert" $MRSLogEntries.TSInfo
                $failClass = "TargetLocked"
                WriteLogEx "Locked mailbox detected" $UserGuid $MDB "TargetLocked"
            }      

            # Target Locked
            # IntraOrg AutoSuspended indicates leaked TargetLock
            # TODO - In future, should check batchname -eq "MRSTS-TargetLock"
            if(! $canConnect -and $mailboxpresent -and $obj.Status -eq "AutoSuspended" -and $requestStyle -eq "IntraOrg")
            {
                #Need to eventlog report before continuing, or send e-mail after failing
                TSNotify ("[Notification] Locked mailbox detecting - attempting to repair user: {0}`r`n`r`nPrevious Move Request:`r`n{1}" -f $UserGuid, ($obj|fl|Out-String )) "MailAlert" $MRSLogEntries.TSInfo
                $failClass = "TargetLocked"
                WriteLogEx "Leaked Locked mailbox detected" $UserGuid $MDB "TargetLocked"
            }

            #Add items to fix up queue if classified and not orphaned unless IgnoreNMR flag set
            if ($failClass -ne "None" -and (!$orphan -or $IgnoreNMR))
            {
                $script:total++
                WriteLogEx "Found object to clean up" $UserGuid $mdb $failclass
                $null = $script:list.Add((New-Object PSObject -Property @{ Status = "Init"; MDB = [String]$mdb; UserGUID = [String]$UserGuid; FailClass = $failClass; RequestStyle = $requestStyle}))
            }
            else
            {
                # Capture everything we didn't process
                $script:unclassified++
                $script:skippedList[("{0,-10}{1,-10}{2,-12}{3,-18}{4,-22}{5,-10}" -f $orphan, $mailboxPresent, $canConnect, $requestStyle, $obj.Status, $IsOld)]++
                if ($failclass -ne "None")
                {
                    $failclassInfo = $failclass
                }
                else
                {
                    $failclassInfo = $obj.Message
                }
                WriteLogEx ("Skipped item - Orphan={0}; Mailbox={1}; CanConnect={2}; Type={3}; MoveState={4}; IsOld={5}; LastUpdate {6}" -f  $orphan, $mailboxPresent, $canConnect, $requestStyle, $obj.Status, $IsOld, $obj.LastUpdateTimestamp) $UserGuid $mdb $failclassInfo
            }
    }

    function ProcessMailbox([ValidateNotNull()]$mbo)
    {
        Write-Verbose "Enter ProcessMailbox"
        
        # switch on status - Pending -> Processing Job -> Waiting to complete -> reset size -> remove
        switch ($mbo.Status)
        {
            "Pending" 
            {
                # Wait for timestamp to expire
                if ($mbo.TimeStamp -ge (Get-Date)) { return }
                $mbo.TimeStamp = (Get-Date).AddMinutes(1)

                # Alert on Non-Canonical ACLS
                if ($mbo.FailClass -eq "NonCanonicalACL")
                {
                    TSNotify ("[Notification] Non-Canonical ACL, requires store team investigation: {0}`r`n`r`nPrevious Move Request:`r`n{1}" -f $mbo.UserGuid, ($mbo|fl|Out-String )) "MailAlert" $MRSLogEntries.TSInfo
                    WriteLog $mbo "Non-Canonical ACL detected"
                    $null = $script:removelist.Add($mbo)
                    return
                }

                if ($Test)
                {
                    $mbo.Status = "Complete"
                    $null = $script:removelist.Add($mbo)
                    WriteLog $mbo "Skipping move because of Test flag"
                    return
                }
                
                # Start processing job
                WriteLog $mbo "Start Processing Job"
                switch ($mbo.FailClass)
                {
                    "MaxSize"
                    {
                        remove-Moverequest $mbo.UserGuid -confirm:$false -ErrorAction SilentlyContinue
                        $mbo.Status = "Delay"
                    }
                    "SetSearchCriteria"
                    {
                        $moveStat = Get-MoveRequestStatistics -Identity $mbo.userguid -IncludeReport
                        $count = 0
                        foreach($baditem in $moveStat.report.baditems)
                        {
                            if($baditem.failure.toString() -match "Unable to SetSearchCriteria")
                            {
                                $count++
                            }
                        }
                        if($count -le $maxBadItems)
                        {
                            Set-MoveRequest $mbo.UserGuid -BadItemLimit $count -Priority High                        
                            $mbo.Status = "Delay"
                        }
                        else
                        {
                            $null = $script:removelist.Add($mbo)
                            return                        
                        }                        
                    }
                    "TargetLocked"
                    {
                        remove-Moverequest $mbo.UserGuid -confirm:$false -ErrorAction SilentlyContinue
                        $mbo.Status = "Delay"
                    }
                    "Resubmit"
                    {
                        remove-Moverequest $mbo.UserGuid -confirm:$false -ErrorAction SilentlyContinue
                        $mbo.Status = "Delay"
                    }                   
                    "SignatureError"
                    {
                        remove-Moverequest $mbo.UserGuid -confirm:$false -ErrorAction SilentlyContinue
                        $mbo.Status = "Delay"
                    }
                    "IncreaseBIL"
                    {
                        $moveStat = Get-MoveRequestStatistics -Identity $mbo.userguid -IncludeReport
                        $count = 0
                        foreach($baditem in $moveStat.report.baditems)
                        {
                            if($baditem.failure.toString() -match "StoreEc: 0x471" -or $baditem.failure.toString() -match "Lid: 40691")
                            {
                                $count++
                            }
                        }
                        Set-MoveRequest $mbo.UserGuid -BadItemLimit $count -Priority High
                        $mbo.Status = "Delay"
                    }                    
                    "BadItem"
                    {
                        if($mbo.RequestStyle -eq "IntraOrg")
                        {
                            TSNotify ("[Requires Investigation] Bad items found in internal moves, store team should have a look at it'r`n`r`n{0}" -f ($mbo|fl|Out-String)) "MailAlert" $MRSLogEntries.TSWarning
                        }
                        
                        #else, it is onboarding/offboarding moves, we cannot do anything but tenant admin need increase the bad item limit                        
                        
                        $null = $script:removelist.Add($mbo)
                        return
                    }
                    "PrimaryruleCorrupted"
                    {
                       if($mbo.RequestStyle -eq "IntraOrg")
                        {
                            TSNotify ("[Requires Investigation] Primary mail box has corrupted rules, pls triage whether to skip them'r`n`r`n{0}" -f ($mbo|fl|Out-String)) "MailAlert" $MRSLogEntries.TSWarning
                        }
                        
                        #else, it is onboarding/offboarding moves, we cannot do anything
                        
                        $null = $script:removelist.Add($mbo)
                        return                    
                    }
                    "badOOFResubmit"
                    {                        
                        if($RunningCA -and $allowKeyChange)
                        {
                            Add-PSSnapin *central*
                            $servers = @()
                            $primary = Get-Mailbox $mbo.UserGuid -ErrorAction SilentlyContinue
                            if($primary -ne $null)
                            {
                                $servers += $primary.ServerName
                            }                         
                            $archive = Get-Mailbox $mbo.UserGuid -Archive -ErrorAction SilentlyContinue
                            if($archive -ne $null)
                            {
                                $servers += $archive.ServerName
                            }
                            
                            foreach($server in $servers)
                            {
                                $results = invoke-command -ComputerName $server -ScriptBlock {$path = "HKLM:\SYSTEM\CurrentControlSet\Services\MSExchangeIS\ParametersPrivate"; $oofRule = "Ignore Bad Oof Rule"; $reg = Get-itemproperty -path $path -name $oofRule -erroraction silentlycontinue; if($reg -ne $null -and $reg.$oofRule -ne $null){Set-ItemProperty -Path $path -Name $oofRule  -Value 1} else{New-ItemProperty -Path $path -name $oofRule  -value 1 -propertyType dword}}
                                WriteLog $mbo "Changed reg key $server output $results "
                                $results = invoke-command -ComputerName $server -ScriptBlock {$path = "HKLM:\SYSTEM\CurrentControlSet\Services\MSExchangeIS\ParametersPrivate"; $oofRule = "Ignore Bad Oof Rule"; Get-itemproperty -path $path -name $oofRule}
                                WriteLog $mbo "Changed reg key $server result $results "
                            }
                            
                            remove-Moverequest $mbo.UserGuid -confirm:$false -ErrorAction SilentlyContinue
                            $mbo.Status = "Delay"
                        }
                        else
                        {
                            if($mbo.RequestStyle -eq "IntraOrg")
                            {
                                TSNotify ("[Requires Investigation] bad oof rule need be ignored, bug 536294, workaround attached in the bug, or run Troubleshoot-MRS.ps1 as CA script to fix this'r`n`r`n{0}" -f ($mbo|fl|Out-String)) "MailAlert" $MRSLogEntries.TSWarning
                            }
                            
                            #else, it is onboarding/offboarding moves, we cannot do anything
                        
                            $null = $script:removelist.Add($mbo)
                            return                        
                        }

                    }
                    "bug532524"
                    {
                        #need store team to fix the bug                        
                        $null = $script:removelist.Add($mbo)
                        return
                    
                    }
                    "fixMBXResubmit"
                    {
                        $mbx = Get-Mailbox -Identity $mbo.userguid
                        New-MailboxRepairRequest -CorruptionType MessagePtagCn -Mailbox $mbx.Identity
                        remove-Moverequest $mbo.UserGuid -confirm:$false -ErrorAction SilentlyContinue
                        $mbo.Status = "Delay"                    
                    }
                    "skipAchiverule"
                    {
                       remove-Moverequest $mbo.UserGuid -confirm:$false -ErrorAction SilentlyContinue
                       $mbo.Status = "Delay"                                        
                    }
                    "None"
                    {
                        remove-Moverequest $mbo.UserGuid -confirm:$false -ErrorAction SilentlyContinue
                        $mbo.Status = "Delay"
                    }
                    "Retry"
                    {
                        $moverequest = Get-MoveRequest $mbo.userguid
                        if(!$moverequest.batchname.contains("TSResumed") -and !$moverequest.batchname.contains("TSRecreated"))
                        {                            
                            $batch = $moverequest.batchname +"TSResumed"
                            Set-MoveRequest $mbo.UserGuid -batchname $batch -Priority High
                            Resume-MoveRequest $mbo.UserGuid
                            $mbo.TimeStamp = Get-Date
                            $mbo.Status = "Moving"
                        }
                        elseif(!$moverequest.batchname.contains("TSRecreated"))
                        {
                            if($moverequest.requestStyle -eq "IntraOrg")
                            {
                                $batch = $moverequest.batchname +"TSRecreated"
                                remove-Moverequest $mbo.UserGuid -confirm:$false -ErrorAction SilentlyContinue
                                new-Moverequest $mbo.UserGuid -TargetDatabase $mbo.MDB -Priority High -AllowLargeItems -BatchName $batch -SuspendWhenReadyToComplete:$mbo.AutoSuspend 
                                $mbo.Status = "Moving"
                            }
                        }
                    }
                    default
                    {
                        Resume-MoveRequest $mbo.UserGuid
                        $mbo.TimeStamp = Get-Date
                        $mbo.Status = "Moving"
                    }
                }
            }

            "Delay"
            {
                # Wait on timestamp to expire
                if ($mbo.TimeStamp -le (Get-Date))
                {
                    $batch = "MRSTS-{0}" -f $mbo.FailClass
                    switch ($mbo.FailClass)
                    {
                        "SetSearchCriteria"                        
                        {
                            Resume-MoveRequest $mbo.UserGuid
                            $mbo.Status = "Moving"
                        }
                        "IncreaseBIL"                                                
                        {
                            Resume-MoveRequest $mbo.UserGuid
                            $mbo.Status = "Moving"
                        }
                        "TargetLocked"
                        {
                            # new move request without target will move to same site.  Suspend when ready to complete will unlock mailbox.
                            WriteLog $mbo "Start SWRTC Moving Job"
                            new-Moverequest $mbo.UserGuid -SuspendWhenReadyToComplete -Priority High -AllowLargeItems -BatchName $batch
                            $mbo.Status = "Moving"
                        }
                        "SignatureError"
                        {
                            WriteLog $mbo "Recreate the request with DoNotPreserveMailboxSignature flag "
                            new-Moverequest $mbo.UserGuid -TargetDatabase $mbo.MDB -Priority High -AllowLargeItems -BatchName $batch -SuspendWhenReadyToComplete:$mbo.AutoSuspend -DoNotPreserveMailboxSignature
                            $mbo.Status = "Moving"
                        }
                        "skipAchiverule"
                        {
                           WriteLog $mbo "Recreate the request with -skipmoving folderrules"
                           new-Moverequest $mbo.UserGuid -TargetDatabase $mbo.MDB -Priority High -AllowLargeItems -BatchName $batch -SuspendWhenReadyToComplete:$mbo.AutoSuspend -skipmoving folderrules
                           $mbo.Status = "Moving"
                        }
                        default
                        {
                            WriteLog $mbo "Start Moving Job"
                            new-Moverequest $mbo.UserGuid -TargetDatabase $mbo.MDB -Priority High -AllowLargeItems -BatchName $batch -SuspendWhenReadyToComplete:$mbo.AutoSuspend
                            $mbo.Status = "Moving"
                        }
                    }
                    $mbo.TimeStamp = Get-Date
                }
            }

            "TestConnection"
            {
                # Wait on timestamp to expire
                if ($mbo.TimeStamp -le (Get-Date))
                {
                        # test mapi connectivity
                        $tmc = Test-MAPIConnectivity $mbo.UserGuid -ErrorVariable +global:ex -ErrorAction SilentlyContinue
                        if ($tmc.Result.Value -ne "Success")
                        {
                            $mbo.RetryCount++
                            # Sleep RetryCount minutes for a max of 28 minutes
                            $mbo.TimeStamp = (Get-Date).AddMinutes($mbo.RetryCount)
                            # Notify if its been more than 5 times
                            if ($mbo.RetryCount -lt 8) { return }
                            
                            if ($mbo.FailClass -eq "TargetLocked")
                            {
                                $mbo.FailClass = "TargetLockedFailure"
                                $batch = "MRSTS-{0}" -f $mbo.FailClass
                                WriteLog $mbo "Start SWRTC Moving Job"
                                
                                # Failed to unlock the mailbox by removing the fake move when in progress, 
                                # Try to unlock it by new move request without target, and let it finish all the way to complete
                                new-Moverequest $mbo.UserGuid -Priority High -AllowLargeItems -BatchName $batch
                                $mbo.Status = "Moving"
                            }
                            else
                            {
                                # Notify after 7 tries - and fall through to complete
                                TSNotify ("[Requires Investigation] Completed move failed to unlock mailbox`r`nConnectivity Result:`r`n{0}`r`nPrevious Move Request:`r`n`r`n{1}" -f ($tmc|fl|Out-String), ($mbo|fl|Out-String)) "MailAlert" $MRSLogEntries.TSCrimsonWarning
                            }                                
                        }
                        else
                        {
                            #reset, remove request
                            $mbo.Status = "Complete"
                            WriteLog $mbo "Completed Moving Job"
                            Remove-MoveRequest $mbo.UserGuid -confirm:$false -ErrorAction SilentlyContinue
                            $null = $script:removelist.Add($mbo)
                        }
                }
            }

            "Moving" 
            {
                # Wait for move to complete
                $mr = Get-MoveRequest $mbo.UserGuid -ErrorAction SilentlyContinue
                if ($mr -eq $null)
                {
                    # check if mailbox is already on target MDB
                    if ((get-mailbox $mbo.UserGuid -ErrorAction SilentlyContinue).Database -eq $mbo.MDB)
                    {
                        WriteLog $mbo "Request completed - Move Request unexpectedly removed"
                        $mbo.Status = "Complete"
                        $null = $removelist.Add($mbo)
                        return
                    }
                    
                    # wait up to 10 minutes for AD Replication
                    if($mbo.TimeStamp.AddMinutes(10) -ge (Get-Date))
                    {
                        return
                    }
                    
                    # abandon move after $maxRetryCount tries
                    if ($mbo.RetryCount -gt $maxRetryCount)
                    {
                        #reset, remove request
                        WriteLog $mbo "Repair move failed to be created possibly due to AD replication - Giving up after 5 attempts"
                        $null = $script:removelist.Add($mbo)
                        TSNotify ("[Requires Investigation] Repair move failed to be created possibly due to AD replication - Giving up`r`nPrevious Move Request:`r`n`r`n{0}" -f ($mbo|fl|Out-String)) "MailAlert" $MRSLogEntries.TSWarning
                        return
                    }
                    # increase retry count, then move back to delay to queue moves again
                    $mbo.RetryCount++
                    $mbo.Status = "Delay"
                    return
                }

                WriteLog $mbo ("Job move status: " + $mr.Status)
                switch ($mr.Status)
                {
                    "Queued" 
                    {
                        # Check Move Request Statistics
                        $mrs = Get-MoveRequestStatistics $mbo.UserGuid -includereport -ErrorAction SilentlyContinue
                        if ($mrs -eq $null)
                        {
                            # wait up to 15 minutes for request to show up
                            if ($mbo.TimeStamp.AddMinutes(15) -ge (Get-Date)) { return }
                            if ($mbo.RetryCount -gt $maxRetryCount)
                            {
                                $null = $script:removelist.Add($mbo)
                                TSNotify ("[Requires Investigation] Move request failed to show up in queue after 15 minutes - possible AD Orphan`r`nPrevious Move Request:`r`n`r`n{0}`r`n{1}" -f ($mbo|fl|Out-String), ($mr|fl|Out-String)) "MailAlert" $MRSLogEntries.TSWarning
                                return
                            }
                            $mbo.RetryCount++
                            remove-Moverequest $mbo.UserGuid -confirm:$false -ErrorAction SilentlyContinue
                            $mbo.Status = "Delay"
                            return
                        }
                        
                        #shouldn't be queued for longer than 2 hours
                        if($mrs.LastUpdateTimestamp -ne $null -and $mrs.LastUpdateTimestamp.AddHours(2) -le (Get-Date))
                        {
                            WriteLog $mbo "Queued for 2 hours"

                            # abandon move after $maxRetryCount tries
                            if ($mbo.RetryCount -gt $maxRetryCount)
                            {
                                #reset, remove request
                                WriteLog $mbo "Move failed after 5 tries"
                                $null = $script:removelist.Add($mbo)
                                TSNotify ("[Requires Investigation] Repair Move failed after 5 tries`r`nPrevious Move Request:`r`n`r`n{0}`r`n{1}" -f ($mbo|fl|Out-String), ($mrs|fl|Out-String)) "MailAlert" $MRSLogEntries.TSWarning
                                return
                            }
                            $mbo.RetryCount++
                            remove-Moverequest $mbo.UserGuid -confirm:$false -ErrorAction SilentlyContinue
                            $mbo.Status = "Delay"
                        }
                    }
                    "InProgress" 
                    { 
                        $mbo.Moved = $true
                        if ($mbo.FailClass -eq "TargetLocked")
                        {
                            Remove-MoveRequest $mbo.UserGuid -confirm:$false 
                            #reset, remove request, TODO: validate mailbox actually unlocked
                            # test mapi connectivity
                            if ($(Test-MAPIConnectivity $mbo.UserGuid -ErrorVariable +global:ex -ErrorAction SilentlyContinue).Result.Value -eq "Success")
                            {
                                $mbo.Status = "Complete"
                                WriteLog $mbo "Completed Unlocking Mailbox"                                
                                $null = $script:removelist.Add($mbo)
                            }
                            else
                            {
                                $mbo.Status = "TestConnection"
                                $mbo.RetryCount = 0
                                $mbo.TimeStamp = (Get-Date).AddMinutes(1)
                            }
                        }
                        else
                        {
                             $mrs = Get-MoveRequestStatistics $mbo.UserGuid -includereport -ErrorAction SilentlyContinue
                             # Alert if > 2 hours and no update
                             if ($mrs -eq $null -and $mbo.TimeStamp.AddHours(4) -le (Get-Date))
                             {
                                WriteLog $mbo "Abort: Inprogress for 4 hours without move request in queue"
                                $null = $script:removelist.Add($mbo)
                                TSNotify ("[Requires Investigation] Repair move stuck in progress for 4 hours without request in queue - Giving up`r`nPrevious Move Request:`r`n`r`n{0}`r`n{1}" -f ($mbo|fl|Out-String), ($mrs|fl|Out-String)) "MailAlert" $MRSLogEntries.TSWarning
                             }
                             
                             if($mrs -ne $null -and $mrs.LastUpdateTimestamp.AddHours(4) -le (Get-Date))
                             {
                                WriteLog $mbo "Abort: Inprogress for 4 hours without update"
                                $null = $script:removelist.Add($mbo)
                                TSNotify ("[Requires Investigation] Repair move stuck in progress for 4 hours without update - Giving up`r`nPrevious Move Request:`r`n`r`n{0}`r`n{1}" -f ($mbo|fl|Out-String), ($mrs|fl|Out-String)) "MailAlert" $MRSLogEntries.TSWarning
                             }
                         }
                    }
                    "AutoSuspended"
                    {
                        if ($mbo.AutoSuspend)
                        {
                                $mbo.Status = "Complete"
                                WriteLog $mbo "Completed Moving Job"
                                $null = $script:removelist.Add($mbo)
                        }
                        else
                        {
                            # unexpected state - alert that move needs to be manually resumed
                            WriteLog $mbo "Move unexpectedly in AutoSuspended state"
                            $null = $script:removelist.Add($mbo)
                            TSNotify ("[Requires Investigation] Repair Move unexpectedly in AutoSuspended state`r`nPrevious Move Request:`r`n`r`n{0}" -f ($mbo|fl|Out-String)) "MailAlert" $MRSLogEntries.TSWarning
                        }
                    }
                    "Completed" 
                    { 
                        # test mapi connectivity
                        if ($(Test-MAPIConnectivity $mbo.UserGuid -ErrorVariable +global:ex -ErrorAction SilentlyContinue).Result.Value -ne "Success")
                        {
                            # Reset count to 0, and retry connection test
                            $mbo.RetryCount = 0
                            $mbo.Status = "TestConnection"
                            $mbo.TimeStamp = (Get-Date).AddMinutes(1)
                            return
                        }
                        #reset, remove request
                        $mbo.Status = "Complete"
                        WriteLog $mbo "Completed Moving Job"
                        Remove-MoveRequest $mbo.UserGuid -confirm:$false
                        $null = $script:removelist.Add($mbo)
                    }
                    "Failed"
                    {
                        # wait 15 minutes for moves to resume before checking failed status                        
                        if($mbo.TimeStamp.AddMinutes(15) -ge (Get-Date) -and ! $mbo.Moved)
                        {
                            return
                        }
                        # abandon move after $maxRetryCount tries
                        if ($mbo.RetryCount -gt $maxRetryCount)
                        {
                            #reset, remove request
                            WriteLog $mbo "Move failed after maxRetryCount tries"
                            $null = $script:removelist.Add($mbo)
                            TSNotify ("[Requires Investigation] Repair Move failed after maxRetryCount tries`r`nPrevious Move Request:`r`n`r`n{0}" -f ($mbo|fl|Out-String)) "MailAlert" $MRSLogEntries.TSWarning
                        }

                        # classify failure
                        $mbo.FailClass = ClassifyFail($mbo)
                        $mbo.RetryCount++
                        $mbo.Status = "Pending"
                        $mbo.Moved = $false
                        switch ($mbo.FailClass)
                        {
                            "None"
                            {
                                WriteLog $mbo "Unclassified Failure - Aborting"
                                $null = $script:removelist.Add($mbo)
                            }
                            default
                            {
                                WriteLog $mbo ("Move {0} - Wait 15 minutes" -f $mbo.FailClass)
                                if($RunningCA)
                                {
                                    $mbo.TimeStamp = (Get-Date).AddMinutes(1)
                                }
                                else
                                {  
                                    $mbo.TimeStamp = (Get-Date).AddMinutes(15)
                                }
                            }
                        }
                    }
                    default
                    { 
                        #reset, remove request
                        WriteLog $mbo "Failed Moving Job"
                        $null = $script:removelist.Add($mbo)
                        TSNotify ("[Requires Investigation] Repair Move unknown FailClass $($mbo.FailClass)`r`nPrevious Move Request:`r`n`r`n{0}" -f ($mbo|fl|Out-String)) "Author" $MRSLogEntries.TSCrimsonWarning
                    }
                }
            }
        }
    }

    function ProcessADList
    {
        # Remove completed mailboxes from list
        foreach($entry in $script:removelist)
        {
            $script:ADlist.Remove($entry)
        }
        foreach($entry in $script:ADlist)
        {
            switch ($entry.Status)
            {
                "Init"
                {
                    # Check MoveRequest again
                    $mr = Get-MoveRequest $entry.UserGuid -ErrorAction SilentlyContinue
                    
                    # Ignore entries that no longer exist
                    if ($mr -eq $null) 
                    {
                        $null = $script:removelist.Add($entry)
                        continue;
                    }
                    
                    # Ignore entries that now have entries in queue
                    if ((Get-MoveRequestStatistics $entry.UserGuid -ErrorAction SilentlyContinue) -ne $null)
                    {
                        $null = $script:removelist.Add($entry)
                        continue;
                    }

                    # check if Target and Source is null, remove request - corrupt
                    if ($mr.TargetDatabase -eq $null -and $mr.SourceDatabase -eq $null)
                    {
                        WriteLogEx "Remove-MoveRequest AD Orphan corrupt" $entry.UserGuid $entry.MDB "ADOrphanCorrupt"
                        if (!$test) { Remove-MoveRequest $entry.UserGuid -confirm:$false }
                        $null = $script:removelist.Add($entry)
                        $script:ADOrphan++
                        continue;
                    }

                    # Remove move request if already on Target
                    $mb = get-mailbox $entry.UserGuid -ErrorAction SilentlyContinue
                    if ($mb -ne $null -and $mr.TargetDatabase -eq $mb.Database)
                    {
                        WriteLogEx "Remove-MoveRequest AD Orphan" $entry.UserGuid $entry.MDB "ADOrphanSync"
                        if (!$test) { Remove-MoveRequest $entry.UserGuid -confirm:$false }
                        $null = $script:removelist.Add($entry)
                        $script:ADOrphan++
                        continue;
                    }

                    # Requeue move request not on target
                    WriteLogEx "Requeue MoveRequest AD Orphan" $entry.UserGuid $entry.MDB "ADOrphanSync"
                    $entry.TimeStamp = (Get-Date).AddMinutes(1)
                    $entry.Status = "CreateMR"
                    if (!$test) { Remove-MoveRequest $entry.UserGuid -confirm:$false }
                    $script:ADOrphan++
                    continue;
                }
                "CreateMR"
                {
                    # Wait for timestamp to expire before checking
                    if ($entry.TimeStamp -ge (Get-Date)) { continue }

                    # Re-queue move request
                    if (!$test) { new-MoveRequest $entry.UserGuid -TargetDatabase $entry.MDB -BatchName "MRSTS-ADOrphan"}
                    $null = $script:removelist.Add($entry)
                    continue;
                }
            }
        }
    }

    function ProcessList
    {
            Write-Verbose "Enter ProcessList - $script:total"
            
            # Remove completed mailboxes from list
            foreach($entry in $script:removelist)
            {
                if ($entry.Status -ne "Complete") { $script:errorcount++ }
                $script:list.Remove($entry)
            }
            $script:removelist.Clear()
            $status.Clear()

            for ($i = 0; $i -lt $script:list.Count; $i++)
            {
                if ($i -ge $BatchSize) { break; }

                $complete = [System.Math]::Round(($i + 1)/$script:list.Count * 100, 1)
                Write-Progress -Activity ("Processing batch item " + ($i + 1) + " out of " + $script:list.Count) -PercentComplete $complete -id 1 -CurrentOperation "$complete% complete" -Status "Please wait..."               

                if ($script:list[$i].Status -eq "Init")
                {
                    # Add additional fields
                    $script:list[$i].Status = "Pending"
                    Add-Member -InputObject $script:list[$i] -MemberType NoteProperty -Name StartTime -value (Get-Date)
                    Add-Member -InputObject $script:list[$i] -MemberType NoteProperty -Name TimeStamp -value (Get-Date)
                    if ((Get-Member -InputObject $script:list[$i] -Name FailClass) -eq $null)
                    {
                        Add-Member -InputObject $script:list[$i] -MemberType NoteProperty -Name FailClass -value (ClassifyFail $script:list[$i])
                    }
                    Add-Member -InputObject $script:list[$i] -MemberType NoteProperty -Name RetryCount -value 0
                    Add-Member -InputObject $script:list[$i] -MemberType NoteProperty -Name Moved $false
                    Add-Member -InputObject $script:list[$i] -MemberType NoteProperty -Name AutoSuspend $false

                    # verify data is not null
                    if ($script:list[$i].UserGuid -eq $null -or $script:list[$i].MDB -eq $null)
                    {
                        Writelog $script:list[$i] "UserGuid and MDB must be non-null"
                        $null = $removelist.Add($script:list[$i])
                        continue
                    }
                    elseif ($script:list[$i].FailClass -eq "TargetLocked")
                    {
                        # Ignore following checks as mailbox is on target and there may be no move request
                    }
                    else
                    {
                        # check if mailbox is already on target MDB
                        if ((get-mailbox $script:list[$i].UserGuid -ErrorAction SilentlyContinue).Database -eq $script:list[$i].MDB)
                        {
                            WriteLog $script:list[$i] "Skipping completed request"
                            $script:list[$i].Status = "Complete"
                            $null = $removelist.Add($script:list[$i])
                            continue
                        }

                        # for reprocessing use $IgnoreNMR flag to force move in list that isn't on target MDB yet
                        if (! $IgnoreNMR) 
                        {
                            # verify a failed move request exists
                            $mr = Get-MoveRequest $script:list[$i].UserGuid -ErrorAction SilentlyContinue
                            if ($mr -eq $null -or $mr.TargetDatabase -ne $script:list[$i].MDB)
                            {
                                WriteLog $script:list[$i] "Unable to locate moverequest"
                                Write-Verbose ("TargetDatabase: {0} != MDB: {1}" -f $mr.TargetDataBase, $script:list[$i].MDB)
                                
                                $null = $removelist.Add($script:list[$i])
                                continue
                            }
                            $script:list[$i].AutoSuspend = $mr.SuspendWhenReadyToComplete
                        }
                    }
                }
                try
                {
                    ProcessMailbox $script:list[$i]
                    $status[$script:list[$i].Status]++
                }
                catch
                {
                    [string]$message =  "Unhandled Exception in ProcessMailbox`r`n"
                    $message +=  "Error text: {0}`r`n" -f ($_ |fl |out-string)
                    $message +=  "Exception: {0}`r`n" -f $_.Exception.ToString()
                    WriteError ($message)
                    $script:list[$i].RetryCount++
                    if ($script:list[$i].RetryCount -gt $maxRetryCount)
                    {
                        $null = $script:removelist.Add($script:list[$i])
                        TSNotify ("[Requires Investigation] Repair Move failed after $maxRetryCount tries`r`nPrevious Move Request:`r`n`r`n{0}" -f ($script:list[$i]|fl|Out-String)) "MailAlert" $MRSLogEntries.TSWarning
                    }
                }
            }

            $completecount = $script:total - $script:list.Count - $script:errorCount
            $complete = [System.Math]::Round(($completecount + $script:errorCount)/$script:total* 100, 1)
            Write-Progress -Activity " " -Status "Batch progress: $script:errorcount Errors and $completecount Completed out of $script:total Total" -percentcomplete $complete  -CurrentOperation "$complete% complete"           
            $status
    }
}

#########################################################
#                                                                                                                           #
#                                                     Script Body                                                     #
#                                                                                                                           #
#########################################################

process
{

       try
    {
        if ($MonitoringContext) 
        {
            Log-Event -Arguments $script:Arguments -EventInfo $MRSLogEntries.TSStarted
        }
        else
        {
            $fswrite = new-object System.IO.FileStream $StatusLogFileName, $fm, $fa, $fs
            $fewrite = new-object System.IO.FileStream $ErrorLogFileName, $fm, $fa, $fs
            WriteLine ("#remove requests - {0}" -f (Get-Date)) "Status"
            WriteLine "Message, UserGuid, TargetMDB, TimeStamp, FailClass, RetryCount" "Status"
        }
        if(!$RunningCA)
        {
            Write-Verbose ("Param Set Name: {0}" -f $PSCmdlet.ParameterSetName)
            $parameterSet = $PSCmdlet.ParameterSetName
        }
        else
        {
            #CA script doesn't support parameterset, so we use this as workaround
            #Unlock is the only scenario that we may not have a move request for it, so we need treate it differently
            #Other situations, we can pass a MoveRequestIdentities to handle all the requests we need fix
            if($UnlockUserGuid -ne $null -and $UnlockUserGuid -ne "")
            {                
                $parameterSet = "Unlock"
            }
            else
            {
                $parameterSet = "MoveRequest"
            }
        }        
        switch ($parameterSet)
        {
            "Database"
            {
                # Check for moves in single database
                ScanMDB $database
            }
            "Server"
            {
                # Check for moves in all databases on a server - Default SCOM mode
                $LogDirectory = Join-Path $exchangeInstallPath "Logging\TroubleshootMRS"
                
                if(-not(Test-Path -Path $LogDirectory  -PathType Container))
                {
                    [Void](New-Item -Path $LogDirectory  -ItemType Container)
                }
                else
                {
                    RemoveOldLogs $LogDirectory 7
                }        
        
                $script:queueStatsLog = "$LogDirectory\TSMRS-QueueStats-{0}.csv" -f (get-date).ToString("yyyyMMdd_hh")
                "DBname, RequestStyle,StatusDetail,Count" | out-file $script:queueStatsLog -Encoding "ASCII"
                $script:crossOrgMsg = ""
                
                $mdbs = get-mailboxdatabase -server $env:ComputerName  | ?{$_.Server -eq $env:ComputerName}                
                foreach ($mdb in $mdbs)
                {
                    ScanMDB $mdb
                }
                                
                $sendAlert = QueueStatMapAnalysis             
                DumpQueueStatMap $sendAlert                
            }
            "MoveRequest"
            {
                # Check moves for forest
                if ($MoveRequestIdentities.count -eq 0)
                {
                    $moves = @(Get-MoveRequest -MoveStatus failed -AccountPartition localforest -resultSize unlimited)
                }
                else
                {
                    $moves = @()
                    foreach ($moveRequestGuid in $MoveRequestIdentities)
                    {
                        $moves += Get-MoveRequest $moveRequestGuid
                    }
                }
                
                if($RunningCA)
                {
                    #In CA mode, we only look at Intraorg movements
                    $moves = @($moves | ?{$_.requestStyle -eq "IntraOrg"})
                }
                $count = 0                
                for($index = 0; $index -lt $moves.count; $index++)
                {
                    $mr = $moves[$index]
                    $count++
                    $complete = [System.Math]::Round($count/$moves.Count* 100, 1)
                    Write-Progress -Activity ("Processing $count out of " + $moves.Count) -PercentComplete $complete -CurrentOperation "$complete% complete" -Status "Please wait..."                    
                    $mrs = Get-MoveRequestStatistics $mr -IncludeReport -ErrorAction SilentlyContinue
                    if ($mrs -ne $null)
                    {
                        ProcessMRS $mrs
                    }
                    else
                    {
                        if ($mr.RequestStyle -eq "IntraOrg")
                        {
                            # Detect AD orphans older than 24 hours
                            $user = [ADSI]("LDAP://{0}" -f $mr.DistinguishedName)
                            $props = @("msDS-ReplAttributeMetaData")
                            $user.RefreshCache($props)
                            foreach($val in $user.Properties.($props[0]))
                            {
                                $x = ([XML]$val).DS_REPL_ATTR_META_DATA
                                if ($x.pszAttributeName -eq "msExchMailboxMoveStatus")
                                {
                                    if ((get-date($x.ftimeLastOriginatingChange)).AddDays(1) -lt (get-date))
                                    {
                                        Write-Verbose ("Possible AD Orphan: {0}, Age: {1}" -f $mr.Guid, (get-date($x.ftimeLastOriginatingChange)))                                        
                                        $null = $script:ADlist.Add((New-Object PSObject -Property @{ Status = "Init"; TimeStamp = (Get-Date); MDB = [String]$mr.TargetDatabase; UserGUID = [String]$mr.Guid; }))
                                    }
                                }
                            }
                        }
                    }
                }
            }
            "Pipeline"
            {
                # Check single move request
                $null = $script:list.Add((New-Object PSObject -Property @{ Status = "Init"; MDB = [String]$mr.TargetDatabase; UserGUID = [String]$mr.Guid; }))
                $script:total = 1
            }
            "Unlock"
            {
                # Check single move request
                $null = $script:list.Add((New-Object PSObject -Property @{ Status = "Init"; MDB = "AutoSelect"; UserGUID = [String]$UnlockUserGuid; FailClass = "TargetLocked"}))
                $script:total = 1
            }
            "ByGuid"
            {
                # Check single move request
                $null = $script:list.Add((New-Object PSObject -Property @{ Status = "Init"; MDB = [String]$MRDatabase; UserGUID = [String]$MRGuid; }))
                $script:total = 1
            }
            "Organization"
            {
                # Check moves for organization
                Get-Moverequest -Organization $organization -resultsize unlimited | %{ $null = $script:list.Add((New-Object PSObject -Property @{ Status = "Init"; MDB = [String]$_.TargetDatabase; UserGUID = [String]$_.Guid; }))}
                $script:total = $script:list.Count;
            }    
            "Organizations"
            {
                # Check moves for multiple organizations - filelist with one org each line
                $orgs = get-content $orgList
                foreach($org in $orgs)
                {
                    Get-Moverequest -Organization $org | %{ $null = $script:list.Add((New-Object PSObject -Property @{ Status = "Init"; MDB = [String]$_.TargetDatabase; UserGUID = [String]$_.Guid; })); $script:total++;}
                    if ($script:list.Count -gt 0) { ProcessList }
                }
            }    
            "Filename"
            {
                # Check moves in file
                Write-Progress -Activity "Importing data from csv..." -Status "Please wait."
                Import-Csv $FileName | Select UserGuid, MDB | %{ Add-Member -InputObject $_ -MemberType NoteProperty -Name Status -value "Init";$null = $script:list.Add($_);}
                $script:total = $script:list.Count
            }
            "Scan"
            {
                # Scan all the dbs for forest
                Write-Progress -Activity "Retrieving all MDBs for the forest..." -PercentComplete 0 -CurrentOperation "0% complete" -Status "Please wait."                
                $dbs = @(Get-MailboxDatabase |sort-Object Name)
                $count = 0
                foreach ($mdb in $dbs)
                {
                    $count++
                    $complete = [System.Math]::Round($count/$dbs.Count* 100, 1)
                    Write-Progress -Activity ("Processing $mdb, $count out of " + $dbs.Count) -PercentComplete $complete -CurrentOperation "$complete% complete" -Status "Please wait..."                    
                    ScanMDB $mdb
                    if ($script:list.Count -gt 0) { ProcessList }
                }
            }
        }
        
        # Wait for list to drain
        while ($script:list.Count -gt 0)
        {
            ProcessList
            # sleep if we still have more work to do
            if (($script:list.Count - $script:removelist.Count) -gt 0)
            {
                Write-Progress -Activity ("Sleeping $sleep seconds...") -id 1 -Status "Please wait..."               
                sleep -seconds $sleep
            }
        }
        # Check any AD Orphans
        while ($script:ADlist.Count -gt 0)
        {
            ProcessADList
            # sleep if we still have more work to do
            if (($script:ADlist.Count - $script:removelist.Count) -gt 0)
            {
                Write-Progress -Activity ("Sleeping $sleep seconds...") -id 1 -Status "Please wait..."               
                sleep -seconds $sleep
            }
        }

        # Log Completion
        if ($MonitoringContext) 
        {
            Log-Event -Arguments $Arguments -EventInfo $MRSLogEntries.TSSuccess
            Add-MonitoringEvent -Id $MRSLogEntries.TSSuccess[0] -Type $EVENT_TYPE_INFORMATION -Message $MRSLogEntries.TSSuccess[1]
        }
    }
    catch
    {
        [string]$message =  "Unhandled Exception in Main`r`n"
        $message +=  "Error text: {0}`r`n" -f ($_ |fl |out-string)
        $message +=  "Exception: {0}`r`n" -f $_.Exception.ToString()
        WriteError ($message)
        
        if ($MonitoringContext)
        {
            Add-MonitoringEvent -Id $MRSLogEntries.TSFailed[0] -Type $EVENT_TYPE_ERROR -Message $message
        }
    }
    finally
    {        
        $message = "$script:total mailboxes submitted`r`n"
        $message += ([string]($script:total - $errorcount - $script:list.Count) + " mailboxes completed`r`n")
        $message += "$script:errorcount errors occured while processing.`r`n"
        $message += "$script:cleanup move requests were cleaned up.`r`n"
        $message += "$script:LastUpdateNull move requests that could not be deserialized or were null.`r`n"
        $message += "$script:processing move requests have been updated in last $ADReplicationDelay hours.`r`n"
        $message += "$script:ADOrphan AD Orphan'd move requests were cleaned up.`r`n"
        $message += "$script:unclassified move requests were ignored and are unclassified (potential junk):`r`n`r`n"
        $message += "IsOld >= {0} days old`r`n" -f $AgedCleanupObjectDelay
        $message += "{0,-10}{1,-10}{2,-12}{3,-18}{4,-22}{5,-10}{6,-9}" -f "Orphan", "Mailbox", "CanConnect", "Type", "State", "IsOld", "Count"
        $message += "{0}`r`n`r`n" -f ($script:skippedList |FT -a -HideTableHeaders| out-string)
        $message += "Total script Duration: {0}`r`n" -f (New-TimeSpan $start $(get-date)).ToString()
        if (!$MonitoringContext) { $message += "Log file Path: $StatusLogFileName`r`n" }
        WriteLine $message "Status"

        if ($script:errorcount -ne 0)
        {
            write-Warning "`r`nErrors occured while processing!  Please check $StatusLogFileName for details."
        }

        if ($fswrite -ne $null) { $fswrite.Close() }
        if ($fewrite -ne $null) { $fewrite.Close() }

        if (($script:total + $script:cleanup + $script:unclassified) -gt 0)
        {
            TSNotify ("[Summary] MRS Troubleshooter completed in mode {0}`r`n{1}" -f $parameterSet, $message) "MailAlert" $MRSLogEntries.TSSummary
        }
        if ($MonitoringContext)
        {
            Write-MonitoringEvents
        }
    }
}
# SIG # Begin signature block
# MIIdqAYJKoZIhvcNAQcCoIIdmTCCHZUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUmJEuiwEa72hXSQ6tanI5oDUd
# U/WgghhkMIIEwzCCA6ugAwIBAgITMwAAAJqamxbCg9rVwgAAAAAAmjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI5
# WhcNMTcwNjMwMTkyMTI5WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OkIxQjctRjY3Ri1GRUMyMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEApkZzIcoArX4o
# w+UTmzOJxzgIkiUmrRH8nxQVgnNiYyXy7kx7X5moPKzmIIBX5ocSdQ/eegetpDxH
# sNeFhKBOl13fmCi+AFExanGCE0d7+8l79hdJSSTOF7ZNeUeETWOP47QlDKScLir2
# qLZ1xxx48MYAqbSO30y5xwb9cCr4jtAhHoOBZQycQKKUriomKVqMSp5bYUycVJ6w
# POqSJ3BeTuMnYuLgNkqc9eH9Wzfez10Bywp1zPze29i0g1TLe4MphlEQI0fBK3HM
# r5bOXHzKmsVcAMGPasrUkqfYr+u+FZu0qB3Ea4R8WHSwNmSP0oIs+Ay5LApWeh/o
# CYepBt8c1QIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFCaaBu+RdPA6CKfbWxTt3QcK
# IC8JMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBAIl6HAYUhsO/7lN8D/8YoxYAbFTD0plm82rFs1Mff9WBX1Hz
# /PouqK/RjREf2rdEo3ACEE2whPaeNVeTg94mrJvjzziyQ4gry+VXS9ZSa1xtMBEC
# 76lRlsHigr9nq5oQIIQUqfL86uiYglJ1fAPe3FEkrW6ZeyG6oSos9WPEATTX5aAM
# SdQK3W4BC7EvaXFT8Y8Rw+XbDQt9LJSGTWcXedgoeuWg7lS8N3LxmovUdzhgU6+D
# ZJwyXr5XLp2l5nvx6Xo0d5EedEyqx0vn3GrheVrJWiDRM5vl9+OjuXrudZhSj9WI
# 4qu3Kqx+ioEpG9FwqQ8Ps2alWrWOvVy891W8+RAwggYHMIID76ADAgECAgphFmg0
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
# bWrJUnMTDXpQzTGCBK4wggSqAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25pbmcg
# UENBIDIwMTECEzMAAABkR4SUhttBGTgAAAAAAGQwCQYFKw4DAhoFAKCBwjAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUU7NWbZJqwmMiGtG7sA6J2KYOeM8wYgYKKwYB
# BAGCNwIBDDFUMFKgKoAoAFQAcgBvAHUAYgBsAGUAcwBoAG8AbwB0AC0ATQBSAFMA
# LgBwAHMAMaEkgCJodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vZXhjaGFuZ2UgMA0G
# CSqGSIb3DQEBAQUABIIBABTaPmSyB+dAQ64JvK9dqDEaUKc3JZ3Df35kOOMHxfxz
# M4KpHN6wwDxJY1JR7rMVNokDGRV/f+BK56Ko5mYQkembtQY9IsLkmDCDVoQj9zsk
# K2htxgTIkrq9WX1ngNnLFlJpltQ7l5bOyTliKkptkCWNa1in0FiY1SUr9hf0sc5C
# RNPZ0ap0lVYKIfgXQoCJwlkk3VPAbpWEOeIQ7j3xlhS30sYoeZO2lVYmDpG2fjbJ
# xPYxCOir3PmtX5ER+QwXupzy01QM5oiUptZbq38z8Pd3tuwLajW7ZeCvLuLnNpEb
# +nrqf1x5nXk1eCa1Lsfvc07oSphuLlxRoQgam8cFe6ehggIoMIICJAYJKoZIhvcN
# AQkGMYICFTCCAhECAQEwgY4wdzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBAhMzAAAA
# mpqbFsKD2tXCAAAAAACaMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
# hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNjA5MDMxODQ0NDRaMCMGCSqGSIb3DQEJ
# BDEWBBTgUxDWnxWlkoZVx34PCGEblI7fwjANBgkqhkiG9w0BAQUFAASCAQAvFGQ4
# hkrGFvyA5ganFqfJc+5ohx4TEF/9OUaenaCeXsKAkfaqNOgCh/fLrjhkElQ1a6k6
# DUmKAMOzQX4DTxiwMsRxBes+6WAVcuKj90x2+oPzsMWJIktB45C5gV4CYypMQAEU
# 6i6m9fYi6JkRQ51CkM3zsDknPQVXXtkfViAFdKeBG5kbh07Zs+6VuXVGRrmjESeS
# 4mZxJel76TEkwVNYh1628rQtvlio61+5CHmxyV11Tku5ddtGOEzuW7pRq0LGMM5O
# VRA8S8uklDJcTIoSn/5P50slkGv59nCd91YsQnlNzwrRCoGwtawFT5+YMZqZCQDE
# ZHjNS1/1LoxqT18Y
# SIG # End signature block
