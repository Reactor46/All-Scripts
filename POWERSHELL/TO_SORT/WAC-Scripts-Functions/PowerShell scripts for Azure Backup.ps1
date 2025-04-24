function Get-WACABAgentStatus {
<#

.SYNOPSIS
Gets agent status

.DESCRIPTION
Gets agent status

.ROLE
Readers

#>

$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
$ErrorActionPreference = "Stop"
Try {
    $azureBackupModuleName = 'MSOnlineBackup'
    $azureBackupModule = Get-Module -ListAvailable -Name $azureBackupModuleName
    if ($azureBackupModule) {
        try {
            Import-Module MSOnlineBackup;
            [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::GetMachineRegistrationStatus($false)
        }
        catch {
            $false
        }
    }
    else {
        $false
    }
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}
}
## [END] Get-WACABAgentStatus ##
function Get-WACABBackupDataDetails {
<#

.SYNOPSIS
Gets the number of backups total storage size and from Azure Backup agent.

.DESCRIPTION
Gets the number of backups total storage size and from Azure Backup agent.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $size = 0
    #  $count = 0

    $storage = Get-OBMachineUsage
    if ($storage) {
        $size = $storage.StorageUsedByMachineInBytes
    }

    $systemstaterp = 0
    $filefolderrp = 0

    $rps = Get-OBAllRecoveryPoints
    foreach ($rp in $rps) {
        if ($rp.DataSources -eq "System State") {
            $systemstaterp += 1;
        }
        else {
            $filefolderrp += 1;
        }
    }

    $props = @{
        storagespace  = $size
        systemstaterp = $systemstaterp
        filefolderrp  = $filefolderrp
    }

    $datadetails = New-Object PSObject
    Add-Member -InputObject $datadetails -MemberType NoteProperty -Name "datadetails" -Value $props
    $datadetails
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABBackupDataDetails ##
function Get-WACABCBDSRPInfo {
<#

.SYNOPSIS
Gets the backup items information for items present in current policy from MAB

.DESCRIPTION
Gets the backup items information for items present in current policy from MAB

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $array = @()
    $DSMap = @{}
    $err = $NULL
    $nextbackuptimeSsb = $NULL
    $nextbackuptimeFiles = $NULL
    $systemstatewriterid = 'DA57A531-E7E7-4346-9A68-B511F551DEB6'
    $systemstateapplicationid = '8C3D00F9-3CE9-4563-B373-19837BC2835E'
    $dscount = 0
    $processedforjobdscount = 0
    <#
 Try
 {
     $err = $NULL
     $task = Get-ScheduledTask | where-Object {$_.TaskName -eq 'Microsoft-OnlineBackup'}
     if ($task -eq $NULL){
         throw "exception"
     }
     $taskinfo = Get-ScheduledTaskInfo -TaskPath $task.TaskPath -TaskName Microsoft-OnlineBackup -ErrorVariable $err
     if ($err){
         throw "exception"
     }
     $nextbackuptimeFiles = $taskinfo.nextruntime
 }
 Catch
 {
 }
 Try
 {
     $err = $NULL
     $task = Get-ScheduledTask | where-Object {$_.TaskName -eq 'Microsoft-OnlineBackup-SystemStateBackup'}
     if ($task -eq $NULL){
         throw "exception"
     }
     $taskinfo = Get-ScheduledTaskInfo -TaskPath $task.TaskPath -TaskName Microsoft-OnlineBackup-SystemStateBackup -ErrorVariable $err
     if ($err -ne $NULL){
         throw "exception"
     }
     $nextbackuptimeSsb = $taskinfo.nextruntime
 }
 Catch
 {
 }
 #>
    $pols = @()
    $jobs = @()
 
    $pol = Get-OBPolicy -ErrorAction SilentlyContinue
    if ($pol) {
        $pols += $pol
    }
    try {
        $spol = Get-OBSystemStatePolicy 
        if ($spol) {
            $pols += $spol
        }
    }
    catch {
 
    }
 
    foreach ($pol in $pols) {
        if ($pol -and $pol.DsList -and $pol.DsList.datasourceid) {
            $dses = $pol.DsList
            for ($i = 0 ; $i -lt $dses.length; $i++) {
                $id = $dses[$i].DataSourceId
                $rpInfo = @(0 .. 9)
                $rpInfo = $rpInfo.ForEach( { $NULL })
                $rpInfo[0] = $dses[$i].DataSourceName
                $rpInfo[1] = $pol.policystate.ToString()
                if ($dses[$i].WriterId -eq $systemstatewriterid -and $dses[$i].ApplicationId -eq $systemstateapplicationid) {
                    $rpInfo[2] = $nextbackuptimeSsb
                }
                else {
                    $rpInfo[2] = $nextbackuptimeFiles
                }
 
                $rpinfo[3] = '-'
                $rpinfo[4] = '-'
                $rpinfo[5] = -1
                $rpinfo[6] = -1
                $rpinfo[7] = '-'
                $rpinfo[8] = '-'
                $rpinfo[9] = '-'
 
                $DSMap[$id] = @{
                    rpinfo    = $rpInfo
                    processed = $false
                }
                $dscount++
            }
        }
    }
 
    $jobs += Get-OBJob -previous 200 #-From ([DateTime]::UtcNow).AddDays(-7) -To ([DateTime]::UtcNow) 
 
    for ($i = $jobs.Count - 1 ; $i -ge 0; $i-- ) {
        if ($processedforjobdscount -eq $dscount) {
            break
        }
        $job = $jobs[$i]
        if ( ($job.jobtype -eq "Backup") -and $job.jobStatus -and $job.jobStatus.datasourcestatus -and $job.jobStatus.datasourcestatus.datasource ) {
            $dses = $job.JobStatus.DatasourceStatus
            for ($j = 0; $j -lt $dses.Length ; $j++) {
                $dsid = $job.JobStatus.DatasourceStatus[$j].Datasource.DataSourceId
                if ($DSMap.ContainsKey($dsid) -and $DSMap[$dsid].processed -eq $false) {
                    $rpInfo = $DSMap[$dsid].rpinfo
                    $rpInfo[3] = $job.JobStatus.starttime
                    if ($job.JobStatus.endtime) {
                        $rpInfo[4] = $job.JobStatus.endtime
                        $rpInfo[9] = ($job.JobStatus.endTime - $job.JobStatus.startTime).ToString()
                    }
                    else {
                        #    $rpInfo[4] =$NULL
                        $rpInfo[9] = ($job.JobStatus.startTime - $job.JobStatus.startTime).ToString()
                    }
                    $rpInfo[5] = $job.JobStatus.datasourcestatus[$j].errorinfo.errorcode
                    $rpInfo[6] = $job.JobStatus.datasourcestatus[$j].errorinfo.DetailedErrorCode
                    $rpInfo[7] = ''
                    $DSMap[$dsid].rpinfo = $rpInfo
                    $DSMap[$dsid].processed = $true
                    $processedforjobdscount++
                }
            }
        }
     
    }
 
    $sources = Get-OBRecoverableSource
    foreach ($ds in $sources) {
        $RecoverableItems = Get-OBRecoverableItem $ds
        $latestPIT = $RecoverableItems[0]
  
        if ($DSMap.ContainsKey($latestPIT.RecoverySourceID)) {
            $rpInfo = $DSMap[$latestPIT.RecoverySourceId].rpinfo
            $rpInfo[8] = $latestPIT.pointintime.ToLocalTime()
            $DSMap[$latestPIT.RecoverySourceId].rpinfo = $rpInfo
        }
    }
 
    foreach ($value in $DSMap.Values) {
        $props = @{
     
            name              = $value.rpinfo[0]
            policystate       = $value.rpinfo[1]
            nextbackuptime    = $value.rpinfo[2]
            starttime         = $value.rpinfo[3]
            endtime           = $value.rpinfo[4]
            errorcode         = $value.rpinfo[5]
            detailederrorcode = $value.rpinfo[6]
            msg               = $value.rpinfo[7]
            latestPIT         = $value.rpinfo[8]
            duration          = $value.rpinfo[9]
            processed         = $value.processed
        }
        $object = new-object psobject -Property $props
        $array += $object
    }
 
    $dsrp = New-Object PSObject
    Add-Member -InputObject $dsrp -MemberType NoteProperty -Name "dsrp" -Value $array
    $dsrp
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABCBDSRPInfo ##
function Get-WACABCustomerDetails {
 <#

.SYNOPSIS
Gets customer details from MAB

.DESCRIPTION
Gets the following items from MAB

       1. ContainerClientId

       2. VaultClientId

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
$ContainerId = $null
$VaultId = $null

Try {
    $ContainerId = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Azure Backup\Config" -Name MachineId -ErrorAction SilentlyContinue
    $ContainerId = ($ContainerId).MachineId
}
Catch {

}
Try {
    $VaultId = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Azure Backup\Config" -Name ResourceId -ErrorAction SilentlyContinue
    $VaultId = ($VaultId).ResourceId
}
Catch {

}
function Get-MsiProperty {
    param([string]$guid, [string]$propertyName, [System.Text.StringBuilder]$stringBuilder)
    [int]$buffer = 0;
    [MsiInterop]::MsiGetProductInfo($guid, $propertyName, $null, [ref]$buffer) | Out-Null;

    $buffer++;

    if ($buffer -gt $stringBuilder.Capacity) {
        $stringBuilder.Capacity = $buffer;
    }

    [MsiInterop]::MsiGetProductInfo($guid, $propertyName, $stringBuilder, [ref]$buffer) | Out-Null;
    $stringBuilder.ToString(0, $buffer);
}


$pinvokeSignature = @'
using System.Runtime.InteropServices;
using System.Text;
public class MsiInterop
{

    [DllImport("msi.dll", CharSet=CharSet.Unicode)]
    public static extern int MsiGetProductInfo(string product, string property, [Out] StringBuilder valueBuf, ref int len);
}
'@

$ErrorActionPreference = "Stop"

Try {
    $VersionNumber = "-"
    Add-Type -TypeDefinition $pinvokeSignature
    $tempStringBuilder = New-Object System.Text.StringBuilder 0;
    #marsAgentProductId
    $guid = "{FFE6D16C-3F87-4192-AF94-DDBEFF165106}"
    $CompleteVersion = Get-MsiProperty $guid "VersionString" $tempStringBuilder;
    if (![string]::IsNullOrEmpty($CompleteVersion)) {
        $VersionNumber = $CompleteVersion.split(".")[2]
    }
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

$props = @{
    ContainerId = $ContainerId
    VaultId = $VaultId
    AgentVersionNumber = $VersionNumber
}

$details = New-Object PSObject
Add-Member -InputObject $details -MemberType NoteProperty -Name "details" -Value $props
$details

}
## [END] Get-WACABCustomerDetails ##
function Get-WACABEnhancedSecurityStatus {
<#

.SYNOPSIS
Gets the enhanced security status on the target.

.DESCRIPTION
Gets the enhanced security status on the target.

.ROLE
Readers

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $status = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::GetEnhancedSecurityStatus()
    $status = $status.TokenState.value__

    if ($status -eq 1) {
        $statusBool = $true
    }
    else {
        $statusBool = $false
    }

    $props = @{
        status = $statusBool
    }

    $statusdetails = New-Object PSObject
    Add-Member -InputObject $statusdetails -MemberType NoteProperty -Name "statusdetails" -Value $props
    $statusdetails
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABEnhancedSecurityStatus ##
function Get-WACABFileFolderPolicyFileSpec {
<#

.SYNOPSIS
Gets the file folder policy file spec.

.DESCRIPTION
Gets the file folder policy file spec.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $fileexisted = $false
    $filespecs = @()

    $pol = Get-OBPolicy

    if ($pol) {
        $fileexisted = $true
        $array = @();
        $specs = Get-OBFileSpec $pol
        foreach ($fs in $specs) {
            $array += $fs.FileSpec
        }
        $filespecs = $array
    }

    $props = @{
        filespecs   = $filespecs
        fileexisted = $fileexisted
    }

    $specdetails = New-Object PSObject
    Add-Member -InputObject $specdetails -MemberType NoteProperty -Name "specdetails" -Value $props
    $specdetails
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABFileFolderPolicyFileSpec ##
function Get-WACABFileFolderPolicyState {
<#

.SYNOPSIS
Get file folder policy state.

.DESCRIPTION
Get file folder policy state.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $fileexisted = $false
    $isPaused = $false

    $pol = Get-OBPolicy

    if ($pol) {
        $fileexisted = $true
        $state = Get-OBPolicyState $pol
        if ($State.ToString() -eq "Paused") {
            $isPaused = $true
        }
        elseif ($pol.State -and $pol.State.ToString() -eq "Valid") {
            $isPaused = $false
        }
    }

    $props = @{
        isPaused    = $isPaused
        fileexisted = $fileexisted
    }

    $statedetails = New-Object PSObject
    Add-Member -InputObject $statedetails -MemberType NoteProperty -Name "statedetails" -Value $props
    $statedetails
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABFileFolderPolicyState ##
function Get-WACABIsBackupJobRunning {
<#

.SYNOPSIS
Fetches if there is a backup job running on the target

.DESCRIPTION
Fetches if there is a backup job running on the target

.ROLE
Readers

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $job = Get-OBJob
    if ($job -eq $NULL -or $job.JobType -ne 'Backup') {
        $false
    }
    else {
        $true
    }
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABIsBackupJobRunning ##
function Get-WACABJobMetrics {
<#

.SYNOPSIS
Gets the metrics of job status from Azure Backup agent.

.DESCRIPTION
Gets the metrics of job status from Azure Backup agent.

.ROLE
Readers

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $ongoing = 0
    $completed = 0
    $failed = 0
    $warning = 0

    $queryjobs = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::QueryJobs()

    if ($queryjobs.RunningJobs) {
        $ongoing = 1
    }

    $finishedjobs = $queryjobs.FinishedJobs
    foreach ($finishedjob in $finishedjobs) {
        #Is this list exclusive? Is there any other state falling into completed, failed or warning?
        #keeping a else separate in case some new states are added to MAB
        if ($finishedjob.JobStatus.JobState -eq "Completed") {
            $completed += 1;
        }
        elseif ($finishedjob.JobStatus.JobState -eq "Aborted") {
            $failed += 1;
        }
        elseif ($finishedjob.JobStatus.JobState -eq "CompletedWithWarning" -or
            $finishedjob.JobStatus.JobState -eq "CompletedWithWaitingForImportJob" -or
            $finishedjob.JobStatus.JobState -eq "CompletedWithWaitingForCopyBlob") {
            $warning += 1;
        }
        else {
            $warning += 1;
        }
    }

    #do we want current job status also?
    $last2jobs = get-objob -Previous 2
    $jobstatus = @()
    foreach ($job in $last2jobs) {
        $type = $job.JobType.ToString()
        $datatype = ''
        if ($job.jobstatus -and $job.JobStatus.DatasourceStatus -and $job.jobstatus.DatasourceStatus.datasource) {
            $datatype = $job.jobstatus.DatasourceStatus.datasource.datasourcename
        }

        $backupjobflag = $job.Jobtype -eq [Microsoft.Internal.CloudBackup.ObjectModel.OMCommon.CBJobType]::Backup
        $status = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::ConvertJobStatusToString($job.JobStatus.JobState, $backupjobflag)
        $time = $job.JobStatus.StartTime

        $props = @{
            type     = $type
            datatype = $datatype
            status   = $status
            time     = $time
        }
        $object = new-object psobject -Property $props
        $jobstatus += $object
    }

    $props = @{
        inProgress     = $ongoing
        success        = $completed
        failed         = $failed
        warning        = $warning
        jobstatusarray = $jobstatus
    }

    $JobMetrics = New-Object PSObject
    Add-Member -InputObject $JobMetrics -MemberType NoteProperty -Name "JobMetrics" -Value $props
    $JobMetrics
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABJobMetrics ##
function Get-WACABJobs {
<#

.SYNOPSIS
Gets the list of last 50 jobs from Azure Backup agent.

.DESCRIPTION
Gets the list of last 50 jobs from Azure Backup agent.

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    #Get-OBJob -Previous 1 -ErrorAction SilentlyContinue | Select-Object jobStatus

    $array = @()
    $jobs = @()
    $currentjob = get-objob
    $queryjobs = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::QueryJobs()
    #$jobs = get-objob -Previous 50
    #$jobs += $currentjob
    $jobs += $queryjobs.FinishedJobs
    $jobs += $queryjobs.RunningJobs
    $errortype = [Microsoft.Internal.EnterpriseStorage.Dls.Utils.Errors.ErrorCode]
    foreach ($job in $jobs) {
        $backupitems = @()
        $jobtype = ''
        $jobstate = ''
        $starttime = ''
        $endtime = ''
        $problem = ''
        $resolution = ''
        $backupitemsstate = @()
        $dsesdetailederrorcode = @()
        $dseserrorcode = @()
        $dsesproblem = @()
        $dsesresolution = @()
        $dsesdatatransferred = @()
        if ($job.jobstatus -and $job.JobStatus.DatasourceStatus -and $job.jobstatus.DatasourceStatus.datasource) {
            $backupitems = @($job.jobstatus.DatasourceStatus.datasource.datasourcename)
        }
        else {
            continue;
        }
        $id = $job.JobId
        $jobtype = $job.JobType.ToString()
        $backupjobflag = $job.Jobtype -eq [Microsoft.Internal.CloudBackup.ObjectModel.OMCommon.CBJobType]::Backup
        $jobstate = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::ConvertJobStatusToString($job.JobStatus.JobState, $backupjobflag)
        $startTime = $job.JobStatus.StartTime
        $endTime = $job.JobStatus.EndTime
        $duration = ($job.JobStatus.EndTime - $job.JobStatus.StartTime).ToString()
        $joberrorcode = $job.JobStatus.ErrorInfo.ErrorCode

        if ($currentjob -and $job -eq $currentjob) {
            $endtime = ''
            $duration = ''
        }

        $dsstates = $job.jobstatus.DatasourceStatus
        foreach ($dsstate in $dsstates) {
        
            $backupitemsstate += [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::ConvertJobStatusToString($dsstate.jobstate, $backupjobflag)
            $dsesdetailederrorcode += $dsstate.ErrorInfo.DetailedErrorCode
            $dseserrorcode += $dsstate.ErrorInfo.ErrorCode
            $errorcode = $dsstate.ErrorInfo.ErrorCode -as $errortype
            $errorinfo = [Microsoft.Internal.EnterpriseStorage.Dls.Utils.Errors.ErrorInfo]::new($errorcode)
            $dsesproblem += $errorinfo.Problem
            $dsesresolution += $errorinfo.Resolution
            $dsesdatatransferred += $dsstate.byteprogress.progress
        }
    
        $props = @{
            backupitems         = $backupitems
            jobtype             = $jobtype
            jobstate            = $jobstate
            starttime           = $starttime
            endtime             = $endtime
            id                  = $id
            detailederrorcode   = $dsesdetailederrorcode
            errorcode           = $dseserrorcode
            problem             = $dsesproblem
            resolution          = $dsesresolution
            duration            = $duration
            backupitemsstate    = $backupitemsstate
            joberrorcode        = $joberrorcode
            dsesdatatransferred = $dsesdatatransferred
        }
        $object = new-object psobject -Property $props
        $array += $object
    }

    $Jobs = New-Object PSObject
    Add-Member -InputObject $Jobs -MemberType NoteProperty -Name "Jobs" -Value $array
    $Jobs
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABJobs ##
function Get-WACABOngoingJob {
<#

.SYNOPSIS
Gets ongoing job

.DESCRIPTION
Gets the ongoing job

.ROLE
Readers

#>

}
## [END] Get-WACABOngoingJob ##
function Get-WACABOngoingJobDetails {
<#

.SYNOPSIS
Gets the list jobs in one last week from Azure Backup agent.

.DESCRIPTION
Gets the list jobs in one last week from Azure Backup agent.

.ROLE
Readers

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    #Get-OBJob -Previous 1 -ErrorAction SilentlyContinue | Select-Object jobStatus

    $isOngoing = $true
    $job = get-objob
    if (!$job)
    {
        $isOngoing = $false
        $job = Get-OBJob -Previous 1 -ErrorAction SilentlyContinue
    }

    $errortype = [Microsoft.Internal.EnterpriseStorage.Dls.Utils.Errors.ErrorCode]

    $backupitems = @()
    $jobtype = ''
    $jobstate = ''
    $starttime = ''
    $endtime = ''
    $problem = ''
    $resolution = ''
    $backupitemsstate = @()
    $dsesdetailederrorcode = @()
    $dseserrorcode = @()
    $dsesproblem = @()
    $dsesresolution = @()
    $dsesdatatransferred = @()
    if ($job.jobstatus -and $job.JobStatus.DatasourceStatus -and $job.jobstatus.DatasourceStatus.datasource) {
        $backupitems = @($job.jobstatus.DatasourceStatus.datasource.datasourcename)
    }
    else {
        continue;
    }
    $id = $job.JobId
    $jobtype = $job.JobType.ToString()
    $backupjobflag = $job.Jobtype -eq [Microsoft.Internal.CloudBackup.ObjectModel.OMCommon.CBJobType]::Backup
    $jobstate = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::ConvertJobStatusToString($job.JobStatus.JobState, $backupjobflag)
    $startTime = $job.JobStatus.StartTime
    $endTime = $job.JobStatus.EndTime
    $duration = ($job.JobStatus.EndTime - $job.JobStatus.StartTime).ToString()
    $joberrorcode = $job.JobStatus.ErrorInfo.ErrorCode

    if ($isOngoing) {
        $endtime = ''
        $duration = ''
    }

    $dsstates = $job.jobstatus.DatasourceStatus
    foreach ($dsstate in $dsstates) {

        $backupitemsstate += [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::ConvertJobStatusToString($dsstate.jobstate, $backupjobflag)
        $dsesdetailederrorcode += $dsstate.ErrorInfo.DetailedErrorCode
        $dseserrorcode += $dsstate.ErrorInfo.ErrorCode
        $errorcode = $dsstate.ErrorInfo.ErrorCode -as $errortype
        $errorinfo = [Microsoft.Internal.EnterpriseStorage.Dls.Utils.Errors.ErrorInfo]::new($errorcode)
        $dsesproblem += $errorinfo.Problem
        $dsesresolution += $errorinfo.Resolution
        $dsesdatatransferred += $dsstate.byteprogress.progress
    }

    $props = @{
        backupitems         = $backupitems
        jobtype             = $jobtype
        jobstate            = $jobstate
        starttime           = $starttime
        endtime             = $endtime
        id                  = $id
        detailederrorcode   = $dsesdetailederrorcode
        errorcode           = $dseserrorcode
        problem             = $dsesproblem
        resolution          = $dsesresolution
        duration            = $duration
        backupitemsstate    = $backupitemsstate
        joberrorcode        = $joberrorcode
        dsesdatatransferred = $dsesdatatransferred
    }

    $JobDetails = New-Object PSObject
    Add-Member -InputObject $JobDetails -MemberType NoteProperty -Name "JobDetails" -Value $props
    $JobDetails
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABOngoingJobDetails ##
function Get-WACABOverview {
<#

.SYNOPSIS
Gets overview info from MAB

.DESCRIPTION
Gets the following items from MAB

       1. Registration status of MAB Agent

       2. Vault name

       3. Subscription ID

       4. Update available(Y/N)

       5. State of last backup

       6. Latest RP

       7. Oldest RP

       8. Next Scheduled Backup

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $registrationstatus = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::GetMachineRegistrationStatus(0)
    $updateavailable = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::GetAgentUpdateInfo().showagentupdatepopup
    $subscriptionid = $NULL
    $vault = $NULL
    $resourceGroup = $NULL
    $vaultKey = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Azure Backup\Config" -Name ServiceResourceName -ErrorAction SilentlyContinue
    $subscriptionidKey = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Azure Backup\Config" -Name SubscriptionId -ErrorAction SilentlyContinue
    $rgkey = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Azure Backup\Config" -Name ResourceGroupName -ErrorAction SilentlyContinue

    if ($vaultKey) {
        $vault = $vaultKey.ServiceResourceName
    }
    if ($subscriptionidKey) {
        $subscriptionid = $subscriptionidKey.SubscriptionId
    }
    if ($rgkey) {
        $resourceGroup = $rgkey.ResourceGroupName
    }


    $lastbackuperrorcode = 0
    $lastbackupdetailederrorcode = 0
    $latestrp = $NULL
    $oldestrp = $NULL
    $nextscheduledbackup = $NULL

    $jobs = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::QueryJobs()
    if ($jobs -and $jobs.lastbackupjob -and $jobs.lastbackupjob.jobstatus -and $jobs.lastbackupjob.jobstatus.errorinfo) {
        $lastbackuperrorcode = $jobs.lastbackupjob.jobstatus.errorinfo.errorcode
        $lastbackupdetailederrorcode = $jobs.lastbackupjob.jobstatus.errorinfo.detailederrorcode
    }

    $rpinfo = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::GetPolicyAndRPInfoForMachine()
    if ($rpinfo -and $rpinfo.RecoveryPointsInfo -and $rpinfo.RecoveryPointsInfo.latestcopy) {
        $latestrp = $rpinfo.RecoveryPointsInfo.latestcopy
    }


    if ($rpinfo -and $rpinfo.RecoveryPointsInfo -and $rpinfo.RecoveryPointsInfo.oldestcopy) {
        $oldestrp = $rpinfo.RecoveryPointsInfo.oldestcopy
    }

    $nextscheduledbackup = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::GetNextRunTimeOfScheduledTask()
    $dateTimeNow = Get-Date
    $dateTimeNow = $dateTimeNow.AddDays(-1)
    # Next scheduled time cannot be in the past
    if ($nextscheduledbackup -lt $dateTimeNow) {
        $nextscheduledbackup = ''
    }

    $props = @{
        registrationstatus          = $registrationstatus
        vault                       = $vault
        subscriptionid              = $subscriptionid
        resourcegroup               = $resourceGroup
        updateavailable             = $updateavailable
        lastbackuperrorcode         = $lastbackuperrorcode
        lastbackupdetailederrorcode = $lastbackupdetailederrorcode
        latestrp                    = $latestrp
        oldestrp                    = $oldestrp
        nextscheduledbackup         = $nextscheduledbackup
    }

    $overview = New-Object PSObject
    Add-Member -InputObject $overview -MemberType NoteProperty -Name "overview" -Value $props
    $overview
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABOverview ##
function Get-WACABPolicies {
<#

.SYNOPSIS
Gets the policy details from MAB

.DESCRIPTION
Gets the policy details from MAB

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $array = @()
    $policyTypeFileFolder = 0
    $policyTypeSystemState = 1
    function flattenScheduleRetention ($pol, $policytype) {
        $backupdays = @()
        $backuptime = @()
        $backupweeklyfrequency = 0
        $dailyretention = 0
        $weeklyretention = 0
        $monthlyretention = 0
        $yearlyretention = 0
        $dslist = ''
 
        if ($pol -and $pol.dslist) {
            $dslist = $pol.dslist.datasourcename
        }
 
        if ($pol -and $pol.backupschedule -and $pol.backupschedule) {
            $backupdays += $pol.backupschedule.schedulerundays
        }
 
        if ($pol -and $pol.backupschedule -and $pol.backupschedule.scheduleruntimes) {
            $backuptime += $pol.backupschedule.scheduleruntimes
        }
 
        if ($pol -and $pol.backupschedule -and $pol.backupschedule.scheduleweeklyfrequency) {
            $backupweeklyfrequency = $pol.backupschedule.scheduleweeklyfrequency
        }
 
        if ($pol -and $pol.RetentionPolicy) {
            $dailyretention = $pol.RetentionPolicy.RetentionDays
        }
 
        if ($pol -and $pol.RetentionPolicy -and $pol.RetentionPolicy.WeeklyLTRSchedule) {
            $weeklyretention = $pol.RetentionPolicy.WeeklyLTRSchedule.RetentionRange / 7
        }
 
        if ($pol -and $pol.RetentionPolicy -and $pol.RetentionPolicy.MonthlyLTRSchedule) {
            $monthlyretention = $pol.RetentionPolicy.MonthlyLTRSchedule.RetentionRange / 31
        }
     
     
        if ($pol -and $pol.RetentionPolicy -and $pol.RetentionPolicy.YearlyLTRSchedule) {
            $yearlyretention = $pol.RetentionPolicy.YearlyLTRSchedule.RetentionRange / 366
        }
 
        $dsArray = @()
        $dsArray += $dslist
        $props = @{
            policyguid            = $pol.PolicyName
            policystate           = $pol.PolicyState.ToString()
            dslist                = $dsArray
            backupdays            = $backupdays
            backuptime            = $backuptime
            backupweeklyfrequency = $backupweeklyfrequency
            dailyretention        = $dailyretention
            weeklyretention       = $weeklyretention
            monthlyretention      = $monthlyretention
            yearlyretention       = $yearlyretention
            policytype            = $policytype
        }
        $object = new-object psobject -Property $props
        return $object;
    }
 
    $fpol = Get-OBPolicy -ErrorAction SilentlyContinue
    if ($fpol) {
        $array += flattenScheduleRetention $fpol $policyTypeFileFolder
    }
    try {
        $spol = Get-OBSystemStatePolicy 
        if ($spol) {
            $array += flattenScheduleRetention $spol $policyTypeSystemState
        }
    }
    catch {
 
    }
 
    $policies = New-Object PSObject
    Add-Member -InputObject $policies -MemberType NoteProperty -Name "policies" -Value $array
    $policies
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABPolicies ##
function Get-WACABPolicyType {
<#

.SYNOPSIS
Fetches the currently present policy on the target

.DESCRIPTION
Fetches the currently present policy on the target

.ROLE
Readers

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $fileexists = $false
    $ssbexists = $false

    $pol = Get-OBPolicy
    $ssbpol = Get-OBSystemStatePolicy

    if ($pol) {
        $fileexists = $true
    }
    if ($ssbpol) {
        $ssbexists = $true
    }

    $props = @{
        fileexists = $fileexists
        ssbexists  = $ssbexists
    }

    $policydetails = New-Object PSObject
    Add-Member -InputObject $policydetails -MemberType NoteProperty -Name "policydetails" -Value $props
    $policydetails
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABPolicyType ##
function Get-WACABPreSetupStatus {
<#########################################################################################################
 # File: GetPreSetupStatus.ps1
 #
 # .DESCRIPTION
 #
 #  Fetches the current status of the MARS agent on target
 #
 #  The supported Operating Systems are Windows Server 2008 R2, Window Server 2012, 
 #
 #  Windows Server 2012R2, Windows Server 2016.
 #
 #  Copyright (c) Microsoft Corp 2017.
 #
 #########################################################################################################>

<#

.SYNOPSIS
Gets pre setup status

.DESCRIPTION
Gets pre setup status

.ROLE
Readers

#>

<#
export enum CBPreSetupStatus {
    CannotConnectToTarget = 0,  // Connection failure
    DPMInstalled = 1,   // DPM/Venus/LaJolla installed
    DPMRAInstalled = 2,  // DPM_RA installed on target machine
    MARSAgentNotInstalled = 3,  // Agent not installed on target server
    MARSAgentNotRegisterted = 4,    // Agent is not registered
    MARSAgentInstalledAndRegistered = 5  // MARS agent is ready to use
    CurrentUserNotAdmin = 6  // Either honolulu server is not running in admin context or current user is not admin
}
#>

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
$ErrorActionPreference = "Stop"
Try {
    $dpm2012RAProductId = '34ACE441-5C52-40CD-A8E6-3521F76F92DA'
    $dpm2016RAProductId = '14DD5B44-17CE-4E89-8BEB-2E6536B81B35'
    $marsAgentProductId = 'FFE6D16C-3F87-4192-AF94-DDBEFF165106'
    $azureBackupModuleName = 'MSOnlineBackup'
    $dpmModuleName = 'DataProtectionManager'
    $dpmModule = Get-Module -ListAvailable -Name $dpmModuleName
    $azureBackupModule = Get-Module -ListAvailable -Name $azureBackupModuleName
    $installedProductList1 = @()
    $installedProductList2 = @()
    $installedProductList3 = @()
    $isAdmin = $false;

    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (!$isAdmin) {
        6
        return
    }

    if ($dpmModule) {
        1
        return
    }
    if ($azureBackupModule) {
        Import-Module $azureBackupModuleName
        try {
            $registrationstatus = [Microsoft.Internal.CloudBackup.Client.Common.CBClientCommon]::GetMachineRegistrationStatus(0)
            if ($registrationstatus -eq $true) {
                5
            }
            else {
                4
            }
        }
        catch {
            7
        }
        return
    }
    $installedProductList1 = Get-Item -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall -ErrorAction SilentlyContinue
    $installedProductList2 = Get-Item -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall -ErrorAction SilentlyContinue
    $installedProductList3 = Get-Item -Path Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall -ErrorAction SilentlyContinue
    $isDPM2012RAInstalled = $false
    $isDPM2016RAInstalled = $false
    $isMARSAgentInstalled = $false
    foreach ($productId in $installedProductList1.GetSubKeyNames()) {
        if ($productId -contains $dpm2012RAProductId) {
            $isDPM2012RAInstalled = $true
            break
        }
        elseif ($productId -contains $dpm2016RAProductId) {
            $isDPM2016RAInstalled = $true
            break
        }
    }
    if (!$isDPM2012RAInstalled -and !$isDPM2016RAInstalled -and $isMARSAgentInstalled) {
        foreach ($productId in $installedProductList2.GetSubKeyNames()) {
            if ($productId -contains $dpm2012RAProductId) {
                $isDPM2012RAInstalled = $true
                break
            }
            elseif ($productId -contains $dpm2016RAProductId) {
                $isDPM2016RAInstalled = $true
                break
            }
        }
    }
    if (!$isDPM2012RAInstalled -and !$isDPM2016RAInstalled -and $isMARSAgentInstalled) {
        foreach ($productId in $installedProductList3.GetSubKeyNames()) {
            if ($productId -contains $dpm2012RAProductId) {
                $isDPM2012RAInstalled = $true
                break
            }
            elseif ($productId -contains $dpm2016RAProductId) {
                $isDPM2016RAInstalled = $true
                break
            }
        }
    }
    if ($isDPM2012RAInstalled -or $isDPM2016RAInstalled) {
        2
        return
    }
    if (!$azureBackupModule) {
        3
        return
    }
    else {
        0
    }
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABPreSetupStatus ##
function Get-WACABProtectableItems {
<#

.SYNOPSIS
Gets the local file system root entities of the machine.

.DESCRIPTION
Gets the local file system root entities of the machine.

.ROLE
Readers

#>

function Get-FileSystemRoot
{
    $volumes = Get-Volumes;

    return $volumes |
        Microsoft.PowerShell.Utility\Select-Object @{Name="DisplayLabel"; Expression={if ($_.FileSystemLabel) { $_.FileSystemLabel + " (" + $_.DriveLetter + ":)"} else { "(" + $_.DriveLetter + ":)" }}},
        @{Name="Path"; Expression={$_.Path +":\"}},
        @{Name="Name"; Expression={$_.DriveLetter +":\"}},
        @{Name="Size"; Expression={($_.Size - $_.SizeRemaining)}}
        #  @{Name="Caption"; Expression={$_.DriveLetter +":\"}},
}

############################################################################################################################

# Helper functions.

############################################################################################################################

<# 
.Synopsis
    Name: Get-VolumePathToPartition
    Description: Gets the list of partitions (that have volumes) in hashtable where key is volume path.

.Returns
    The list of partitions (that have volumes) in hashtable where key is volume path.
#>
function Get-VolumePathToPartition
{
    $volumePaths = @{}

    foreach($partition in Get-Partition)
    {
        foreach($volumePath in @($partition.AccessPaths))
        {
            if($volumePath -and (-not $volumePaths.Contains($volumePath)))
            {
                $volumePaths.Add($volumePath, $partition)
            }
        }
    }
    
    $volumePaths
}

<# 
.Synopsis
    Name: Get-DiskIdToDisk
    Description: Gets the list of all the disks in hashtable where key is:
                 "Disk.Path" in case of WS2016 and above.
                 OR
                 "Disk.ObjectId" in case of WS2012 and WS2012R2.

.Returns
    The list of partitions (that have volumes) in hashtable where key is volume path.
#>
function Get-DiskIdToDisk
{    
    $diskIds = @{}

    $isDownlevel = [Environment]::OSVersion.Version.Major -lt 10;

    # In downlevel Operating systems. MSFT_Partition.DiskId is equal to MSFT_Disk.ObjectId
    # However, In WS2016 and above,   MSFT_Partition.DiskId is equal to MSFT_Disk.Path

    foreach($disk in Get-Disk)
    {
        if($isDownlevel)
        {
            $diskId = $disk.ObjectId
        }
        else
        {
            $diskId = $disk.Path
        }

        if(-not $diskIds.Contains($diskId))
        {
            $diskIds.Add($diskId, $disk)
        }
    }

    return $diskIds
}

<# 
.Synopsis
    Name: Get-VolumeDownlevelOS
    Description: Gets the list of all applicable volumes from WS2012 and Ws2012R2 Operating Systems.
                 
.Returns
    The list of all applicable volumes
#>
function Get-VolumeDownlevelOS
{
    $volumes = @()
    $partitionsMapping = Get-VolumePathToPartition
    $disksMapping =  Get-DiskIdToDisk
    
    foreach($volume in (Get-WmiObject -Class MSFT_Volume -Namespace root/Microsoft/Windows/Storage))
    {
       $partition = $partitionsMapping.Get_Item($volume.Path)

       # Check if this volume is associated with a partition.
       if($partition)
       {
            # If this volume is associated with a partition, then get the disk to which this partition belongs.
            $disk = $disksMapping.Get_Item($partition.DiskId)

            # If the disk is a clustered disk then simply ignore this volume.
            if($disk -and $disk.IsClustered) {continue}
       }
  
       $volumes += $volume
    }


    return $volumes
}

<# 
.Synopsis
    Name: Get-VolumeWs2016AndAboveOS
    Description: Gets the list of all applicable volumes from WS2016 and above Operating System.
                 
.Returns
    The list of all applicable volumes
#>
function Get-VolumeWs2016AndAboveOS
{
    $volumes = @()
    
    $applicableVolumePaths = @{}

    $subSystem = Get-CimInstance -ClassName MSFT_StorageSubSystem -Namespace root/Microsoft/Windows/Storage| Where-Object { $_.FriendlyName -like "Win*" }

    foreach($volume in @($subSystem | Get-CimAssociatedInstance -ResultClassName MSFT_Volume))
    {
        if(-not $applicableVolumePaths.Contains($volume.Path))
        {
            $applicableVolumePaths.Add($volume.Path, $null)
        }
    }

    foreach($volume in (Get-WmiObject -Class MSFT_Volume -Namespace root/Microsoft/Windows/Storage))
    {
        if(-not $applicableVolumePaths.Contains($volume.Path)) { continue }

        $volumes += $volume
    }

    return $volumes
}

<#
.Synopsis
    Name: Get-Volumes
    Description: Gets the local volumes of the machine.

.Returns
    The local volumes.
#>
function Get-Volumes
{
    Remove-Module Storage -ErrorAction Ignore; # Remove the Storage module to prevent it from automatically localizing

    $isDownlevel = [Environment]::OSVersion.Version.Major -lt 10;
    if ($isDownlevel)
    {
        $volumes = Get-VolumeDownlevelOS
    }
    else
    {
        $volumes = Get-VolumeWs2016AndAboveOS
    }

    return $volumes | Where-Object { [byte]$_.DriveLetter -ne 0 -and $_.DriveLetter -ne $null -and $_.Size -gt 0 -and $_.FileSystem -eq 'NTFS'};
}

Get-FileSystemRoot;

}
## [END] Get-WACABProtectableItems ##
function Get-WACABRPInfo {
<#

.SYNOPSIS
Gets the recovery points information from MAB

.DESCRIPTION
Gets the recovery points information from MAB

.ROLE
Readers

#>

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $array = @()

    $rps = Get-OBAllRecoveryPoints
    foreach ($rp in $rps) {
        $time = $rp.BackupTime
        $rpinfo = $rp.DataSources
    
        $props = @{
            time   = $time
            rpinfo = $rpinfo
        }
        $object = new-object psobject -Property $props
    
        $array += $object
    }

    $Rps = New-Object PSObject
    Add-Member -InputObject $Rps -MemberType NoteProperty -Name "Rps" -Value $array
    $Rps
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABRPInfo ##
function Get-WACABRPMetrics {
<#

.SYNOPSIS
Gets the metrics of RP status from Azure Backup agent.

.DESCRIPTION
Gets the metrics of RP status from Azure Backup agent.

.ROLE
Readers

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $systemstaterp = 0
    $filefolderrp = 0


    $rps = Get-OBAllRecoveryPoints
    foreach ($rp in $rps) {
        if ($rp.DataSources -eq "System State") {
            $systemstaterp += 1;
        }
        else {
            $filefolderrp += 1;
        }
    }

    $props = @{
        systemstaterp = $systemstaterp
        filefolderrp  = $filefolderrp
    }

    $RpMetrics = New-Object PSObject
    Add-Member -InputObject $RpMetrics -MemberType NoteProperty -Name "RpMetrics" -Value $props
    $RpMetrics
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABRPMetrics ##
function Get-WACABRecoverableItems {
<#

.SYNOPSIS
Gets the list of sources and PITS which can be recovered.

.DESCRIPTION
Gets the list of sources and PITS which can be recovered.

.ROLE
Readers

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $sources = Get-OBRecoverableSource
    $sourceNames = @()
    $itemsForAllSources = @()
    $itemTimesForAllSources = @()
    foreach ($source in $sources) {
        $items = @()
        $itemTimes = @()
        $sourceNames += $source.RecoverySourceName
        $itemsPerSource = Get-OBRecoverableItem $source
        foreach ($itemPerSource in $itemsPerSource) {
            $items += $itemPerSource
            $itemTimes += $itemPerSource.RecoveryPointLocalTime
        }
        $itemsForAllSources += , $items
        $itemTimesForAllSources += , $itemTimes
    }

    $props = @{
        sourceNames            = $sourceNames
        itemsForAllSources     = $itemsForAllSources
        itemTimesForAllSources = $itemTimesForAllSources
    }

    $recoverableitems = New-Object PSObject
    Add-Member -InputObject $recoverableitems -MemberType NoteProperty -Name "recoverableitems" -Value $props
    $recoverableitems
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABRecoverableItems ##
function Get-WACABSystemStatePolicyState {
<#

.SYNOPSIS
Fetches the system state policy state

.DESCRIPTION
Fetches the system state policy state

.ROLE
Readers

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $ssbexisted = $false
    $isPaused = $false

    $pol = Get-OBSystemStatePolicy

    if ($pol) {
        $ssbexisted = $true
        $state = Get-OBPolicyState $pol
        if ($State.ToString() -eq "Paused") {
            $isPaused = $true
        }
        elseif ($pol.State -and $pol.State.ToString() -eq "Valid") {
            $isPaused = $false
        }
    }

    $props = @{
        isPaused   = $isPaused
        ssbexisted = $ssbexisted
    }

    $statedetails = New-Object PSObject
    Add-Member -InputObject $statedetails -MemberType NoteProperty -Name "statedetails" -Value $props
    $statedetails
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Get-WACABSystemStatePolicyState ##
function New-WACABCert {
 <#

.SYNOPSIS
Generates the certificate required to create a vault cred file and register the agent on target

.DESCRIPTION
Generates the certificate required to create a vault cred file and register the agent on target

.ROLE
Readers

#>

 Set-StrictMode -Version 5.0
 $Subject = "CN=Windows Azure Tools"
 $NotBefore = [DateTime]::Now.AddDays(-1)
 $NotAfter = $NotBefore.AddDays(2)
 $ProviderName = "Microsoft Enhanced Cryptographic Provider v1.0"
 $AlgorithmName = "RSA"
 $KeyLength = 2048
 $KeySpec = "Exchange"
 $PathLength = -1
 $SignatureAlgorithm = "SHA1"
 $FriendlyName = "AzureBackupVaultCredCert"
 $StoreLocation = "LocalMachine"
 $Exportable = $true
 $EnhancedKeyUsage = '1.3.6.1.5.5.7.3.2'
 $KeyUsage = $null
 $SubjectAlternativeName = $null
 $CustomExtension = $null
 $SerialNumber = $null
 $AllowSMIME = $false

 $ErrorActionPreference = "Stop"
 if ([Environment]::OSVersion.Version.Major -lt 6) {
                 $NotSupported = New-Object NotSupportedException -ArgumentList "Windows XP and Windows Server 2003 are not supported!"
                 throw $NotSupported
 }
 $ExtensionsToAdd = @()

 #region constants
 # contexts
 New-Variable -Name UserContext -Value 0x1 -Option ReadOnly -Force
 New-Variable -Name MachineContext -Value 0x2 -Option ReadOnly -Force
 # encoding
 New-Variable -Name Base64Header -Value 0x0 -Option ReadOnly -Force
 New-Variable -Name Base64 -Value 0x1 -Option ReadOnly -Force
 New-Variable -Name Binary -Value 0x3 -Option ReadOnly -Force
 New-Variable -Name Base64RequestHeader -Value 0x4 -Option ReadOnly -Force
 # SANs
 New-Variable -Name OtherName -Value 0x1 -Option ReadOnly -Force
 New-Variable -Name RFC822Name -Value 0x2 -Option ReadOnly -Force
 New-Variable -Name DNSName -Value 0x3 -Option ReadOnly -Force
 New-Variable -Name DirectoryName -Value 0x5 -Option ReadOnly -Force
 New-Variable -Name URL -Value 0x7 -Option ReadOnly -Force
 New-Variable -Name IPAddress -Value 0x8 -Option ReadOnly -Force
 New-Variable -Name RegisteredID -Value 0x9 -Option ReadOnly -Force
 New-Variable -Name Guid -Value 0xa -Option ReadOnly -Force
 New-Variable -Name UPN -Value 0xb -Option ReadOnly -Force
 # installation options
 New-Variable -Name AllowNone -Value 0x0 -Option ReadOnly -Force
 New-Variable -Name AllowNoOutstandingRequest -Value 0x1 -Option ReadOnly -Force
 New-Variable -Name AllowUntrustedCertificate -Value 0x2 -Option ReadOnly -Force
 New-Variable -Name AllowUntrustedRoot -Value 0x4 -Option ReadOnly -Force
 # PFX export options
 New-Variable -Name PFXExportEEOnly -Value 0x0 -Option ReadOnly -Force
 New-Variable -Name PFXExportChainNoRoot -Value 0x1 -Option ReadOnly -Force
 New-Variable -Name PFXExportChainWithRoot -Value 0x2 -Option ReadOnly -Force
 #endregion

 #region Subject processing
 # http://msdn.microsoft.com/en-us/library/aa377051(VS.85).aspx
 $SubjectDN = New-Object -ComObject X509Enrollment.CX500DistinguishedName
 $SubjectDN.Encode($Subject, 0x0)
 #endregion

 #region Extensions

 #region Enhanced Key Usages processing
 if ($EnhancedKeyUsage) {
                 $OIDs = New-Object -ComObject X509Enrollment.CObjectIDs
                 $EnhancedKeyUsage | ForEach-Object {
                                 $OID = New-Object -ComObject X509Enrollment.CObjectID
                                 $OID.InitializeFromValue("1.3.6.1.5.5.7.3.2")
                                 # http://msdn.microsoft.com/en-us/library/aa376785(VS.85).aspx
                                 $OIDs.Add($OID)
                 }
                 # http://msdn.microsoft.com/en-us/library/aa378132(VS.85).aspx
                 $EKU = New-Object -ComObject X509Enrollment.CX509ExtensionEnhancedKeyUsage
                 $EKU.InitializeEncode($OIDs)
                 $ExtensionsToAdd += "EKU"
 }
 #endregion

 #region Key Usages processing
 if ($KeyUsage -ne $null) {
                 $KU = New-Object -ComObject X509Enrollment.CX509ExtensionKeyUsage
                 $KU.InitializeEncode([int]$KeyUsage)
                 $KU.Critical = $true
                 $ExtensionsToAdd += "KU"
 }
 #endregion

 #region Basic Constraints processing
 if ($PSBoundParameters.Keys.Contains("IsCA")) {
                 # http://msdn.microsoft.com/en-us/library/aa378108(v=vs.85).aspx
                 $BasicConstraints = New-Object -ComObject X509Enrollment.CX509ExtensionBasicConstraints
                 if (!$IsCA) {$PathLength = -1}
                 $BasicConstraints.InitializeEncode($IsCA,$PathLength)
                 $BasicConstraints.Critical = $IsCA
                 $ExtensionsToAdd += "BasicConstraints"
 }
 #endregion

 #region SAN processing
 if ($SubjectAlternativeName) {
                 $SAN = New-Object -ComObject X509Enrollment.CX509ExtensionAlternativeNames
                 $Names = New-Object -ComObject X509Enrollment.CAlternativeNames
                 foreach ($altname in $SubjectAlternativeName) {
                                 $Name = New-Object -ComObject X509Enrollment.CAlternativeName
                                 if ($altname.Contains("@")) {
                                                 $Name.InitializeFromString($RFC822Name,$altname)
                                 } else {
                                                 try {
                                                                 $Bytes = [Net.IPAddress]::Parse($altname).GetAddressBytes()
                                                                 $Name.InitializeFromRawData($IPAddress,$Base64,[Convert]::ToBase64String($Bytes))
                                                 } catch {
                                                                 try {
                                                                                 $Bytes = [Guid]::Parse($altname).ToByteArray()
                                                                                 $Name.InitializeFromRawData($Guid,$Base64,[Convert]::ToBase64String($Bytes))
                                                                 } catch {
                                                                                 try {
                                                                                                 $Bytes = ([Security.Cryptography.X509Certificates.X500DistinguishedName]$altname).RawData
                                                                                                 $Name.InitializeFromRawData($DirectoryName,$Base64,[Convert]::ToBase64String($Bytes))
                                                                                 } catch {$Name.InitializeFromString($DNSName,$altname)}
                                                                 }
                                                 }
                                 }
                                 $Names.Add($Name)
                 }
                 $SAN.InitializeEncode($Names)
                 $ExtensionsToAdd += "SAN"
 }
 #endregion

 #region Custom Extensions
 if ($CustomExtension) {
                 $count = 0
                 foreach ($ext in $CustomExtension) {
                                 # http://msdn.microsoft.com/en-us/library/aa378077(v=vs.85).aspx
                                 $Extension = New-Object -ComObject X509Enrollment.CX509Extension
                                 $EOID = New-Object -ComObject X509Enrollment.CObjectId
                                 $EOID.InitializeFromValue($ext.Oid.Value)
                                 $EValue = [Convert]::ToBase64String($ext.RawData)
                                 $Extension.Initialize($EOID,$Base64,$EValue)
                                 $Extension.Critical = $ext.Critical
                                 New-Variable -Name ("ext" + $count) -Value $Extension
                                 $ExtensionsToAdd += ("ext" + $count)
                                 $count++
                 }
 }
 #endregion

 #endregion

 #region Private Key
 # http://msdn.microsoft.com/en-us/library/aa378921(VS.85).aspx
 $PrivateKey = New-Object -ComObject X509Enrollment.CX509PrivateKey
 $PrivateKey.ProviderName = $ProviderName
 $AlgID = New-Object -ComObject X509Enrollment.CObjectId
 $AlgID.InitializeFromValue(([Security.Cryptography.Oid]$AlgorithmName).Value)
 $PrivateKey.Algorithm = $AlgID
 # http://msdn.microsoft.com/en-us/library/aa379409(VS.85).aspx
 $PrivateKey.KeySpec = switch ($KeySpec) {"Exchange" {1}; "Signature" {2}}
 $PrivateKey.Length = $KeyLength
 # key will be stored in current user certificate store
 $PrivateKey.MachineContext = if ($StoreLocation -eq "LocalMachine") {$true} else {$false}
 $PrivateKey.ExportPolicy = if ($Exportable) {1} else {0}
 $PrivateKey.Create()
 #endregion

 # http://msdn.microsoft.com/en-us/library/aa377124(VS.85).aspx
 $Cert = New-Object -ComObject X509Enrollment.CX509CertificateRequestCertificate
 if ($PrivateKey.MachineContext) {
                 $Cert.InitializeFromPrivateKey($MachineContext,$PrivateKey,"")
 } else {
                 $Cert.InitializeFromPrivateKey($UserContext,$PrivateKey,"")
 }
 $Cert.Subject = $SubjectDN
 $Cert.Issuer = $Cert.Subject
 $Cert.NotBefore = $NotBefore
 $Cert.NotAfter = $NotAfter
 foreach ($item in $ExtensionsToAdd) {$Cert.X509Extensions.Add((Get-Variable -Name $item -ValueOnly))}
 if (![string]::IsNullOrEmpty($SerialNumber)) {
                 if ($SerialNumber -match "[^0-9a-fA-F]") {throw "Invalid serial number specified."}
                 if ($SerialNumber.Length % 2) {$SerialNumber = "0" + $SerialNumber}
                 $Bytes = $SerialNumber -split "(.{2})" | Where-Object {$_} | ForEach-Object{[Convert]::ToByte($_,16)}
                 $ByteString = [Convert]::ToBase64String($Bytes)
                 $Cert.SerialNumber.InvokeSet($ByteString,1)
 }
 if ($AllowSMIME) {$Cert.SmimeCapabilities = $true}
 $SigOID = New-Object -ComObject X509Enrollment.CObjectId
 $SigOID.InitializeFromValue(([Security.Cryptography.Oid]$SignatureAlgorithm).Value)
 $Cert.SignatureInformation.HashAlgorithm = $SigOID
 # completing certificate request template building
 $Cert.Encode()

 # interface: http://msdn.microsoft.com/en-us/library/aa377809(VS.85).aspx
 $Request = New-Object -ComObject X509Enrollment.CX509enrollment
 $Request.InitializeFromRequest($Cert)
 $Request.CertificateFriendlyName = $FriendlyName
 $endCert = $Request.CreateRequest($Base64)
 $Request.InstallResponse($AllowUntrustedCertificate,$endCert,$Base64,"")
 [Byte[]]$CertBytes = [Convert]::FromBase64String($endCert)
 $certificate = New-Object Security.Cryptography.X509Certificates.X509Certificate2 @(,$CertBytes)

 $path = "cert:\LocalMachine\My\" + $certificate.Thumbprint

 $vaultCert = Get-ChildItem -path $path

 if($vaultCert)
 {
     $certPublicData = [System.Convert]::ToBase64String($vaultCert.RawData)
     $certPrivateData = [System.Convert]::ToBase64String($vaultCert.Export('Pfx'))
 }
 else
 {
     throw "=== Exception Exception Exception ==="
 }
 $props = @{
     privatedata = $certPrivateData
     publicdata = $certPublicData
     thumbprint = $vaultCert.Thumbprint
 }

 $generatecertificate = New-Object PSObject
 Add-Member -InputObject $generatecertificate -MemberType NoteProperty -Name "generatecertificate" -Value $props
 $generatecertificate

}
## [END] New-WACABCert ##
function New-WACABCertificate {
<#

.SYNOPSIS
Generates certificates

.DESCRIPTION
Generates certificates

.ROLE
Administrators

#>

# Works only for >= Windows 10 and Win Server 2016

$certName = "{0}{1}-{2}-vaultcredentials" -f 'prefix', 'subscriptionId', (Get-Date -f "M-d-yyyy")
$startTime = [System.DateTime]::UtcNow
$startTime = $startTime.AddMinutes(-15)
$endTime = [System.DateTime]::UtcNow
$endTime = $endTime.AddDays(1)
# keeping it here for reference
#$cert = New-SelfSignedCertificate -DnsName azurebackup -CertStoreLocation cert:\LocalMachine\My
$cert = New-SelfSignedCertificate -Subject "CN=Windows Azure Tools" -NotBefore $startTime -NotAfter $endTime -Provider "Microsoft Enhanced Cryptographic Provider v1.0" -KeyAlgorithm RSA -KeyExportPolicy Exportable -KeyLength 2048 -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.2") -KeyUsage None
$certPublicData = [System.Convert]::ToBase64String($cert.RawData)
$certPrivateData = [System.Convert]::ToBase64String($cert.Export('Pfx'))
$props = @{
    privatedata = $certPrivateData
    publicdata = $certPublicData
    thumbprint = $cert.Thumbprint
}

$generatecertificate = New-Object PSObject
Add-Member -InputObject $generatecertificate -MemberType NoteProperty -Name "generatecertificate" -Value $props
$generatecertificate

}
## [END] New-WACABCertificate ##
function New-WACABCertificateMakeCert {
<#########################################################################################################
# File: GenerateCertificateMAkeCert.ps1
#
# .DESCRIPTION
#
#  Generates the certificate required to create a vault cred file and register the agent on target
#  Uses Makecert which needs to be present on target machine
#
#  The supported Operating Systems are Windows Server 2008 R2, Window Server 2012, 
#
#  Windows Server 2012R2, Windows Server 2016.
#
#  Copyright (c) Microsoft Corp 2017.
#
#########################################################################################################>

<#

.SYNOPSIS
Generates certificates

.DESCRIPTION
Generates certificates

.ROLE
Administrators

#>
Set-StrictMode -Version 5.0
$certName = "{0}{1}-{2}-vaultcredentials" -f 'prefix', 'subscriptionId', (Get-Date -f "M-d-yyyy")
$certFileName = $certName + '.cer'
$startTime = [System.DateTime]::UtcNow
$startTime = $startTime.AddMinutes(-15)
$endTime = [System.DateTime]::UtcNow
$endTime = $endTime.AddDays(1)
$endTime = $endTime.tostring("MM/dd/yyyy")
$makecertResult = makecert.exe -r -pe -n CN=$certName -ss my -sr localmachine -eku 1.3.6.1.5.5.7.3.2 -len 2048 -e $endTime $certFileName
$certs = Get-ChildItem -Path "cert:\localMachine\my"
$vaultCert = $NULL
foreach ($cert in $certs)
{
    if($cert.SubjectName.Name -match $certName)
    {
        $vaultCert = $cert
        break;
    }
}
if($vaultCert)
{
    $certPublicData = [System.Convert]::ToBase64String($vaultCert.RawData)
    $certPrivateData = [System.Convert]::ToBase64String($vaultCert.Export('Pfx'))
}
else
{
    throw "=== Exception Exception Exception ==="
}
$props = @{
    privatedata = $certPrivateData
    publicdata = $certPublicData
    thumbprint = $vaultCert.Thumbprint
}

$generatecertificate = New-Object PSObject
Add-Member -InputObject $generatecertificate -MemberType NoteProperty -Name "generatecertificate" -Value $props
$generatecertificate

}
## [END] New-WACABCertificateMakeCert ##
function Register-WACABMARSAgent {
<#

.SYNOPSIS
Registers MARS agent

.DESCRIPTION
Registers MARS agent

.ROLE
Administrators

#>
param (
    [Parameter(Mandatory = $true)]
    [String]
    $vaultCredString,
    [Parameter(Mandatory = $true)]
    [String]
    $passphrase
)
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $date = Get-Date
    $vaultcredPath = $env:TEMP + '\honoluluvaultcredential_' + $date.Day + "_" + $date.Month + "_" + $date.Year + "_" + '.vaultcredentials';
    $vaultCredString | Out-File $vaultcredPath
    Start-OBRegistration -VaultCredentials $vaultcredPath -Confirm:$false
    $securePassphrase = ConvertTo-SecureString -String $passphrase -AsPlainText -Force
    Set-OBMachineSetting -EncryptionPassphrase $securePassphrase -SecurityPIN " "
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}


}
## [END] Register-WACABMARSAgent ##
function Remove-WACABAllPolicies {
<#

.SYNOPSIS
Deletes agent status

.DESCRIPTION
Deletes agent status

.ROLE
Administrators

#>
param (
    [Parameter(Mandatory = $false)]
    [string]
    $pin
)

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $fileexisted = $false
    $ssbexisted = $false
    $filedeleted = $false
    $ssbdeleted = $false

    $pol = Get-OBPolicy
    $ssbpol = Get-OBSystemStatePolicy

    if ($pol) {
        $fileexisted = $true
    }
    if ($ssbpol) {
        $ssbexisted = $true
    }

    if ($pol -or $ssbpol) {
        if ($pol) {
            $ans = Remove-OBPolicy -Policy $pol -SecurityPIN $pin -Confirm:$false -DeleteBackup:$true
        }
        else {
            $ans = Remove-OBSystemStatePolicy -Policy $pol -SecurityPIN $pin -Confirm:$false -DeleteBackup:$true
        }

        $policy = get-obpolicy
        if ($policy) {
            $filedeleted = $false
        }
        else {
            $filedeleted = $true
        }

        $ssbpolicy = Get-OBSystemStatePolicy
        if ($ssbpolicy) {
            $ssbdeleted = $false
        }
        else {
            $ssbdeleted = $true
        }
    }

    $props = @{
        filedeleted = $filedeleted
        ssbdeleted  = $ssbdeleted
        fileexisted = $fileexisted
        ssbexisted  = $ssbexisted
    }

    $deletiondetails = New-Object PSObject
    Add-Member -InputObject $deletiondetails -MemberType NoteProperty -Name "deletiondetails" -Value $props
    $deletiondetails

}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Remove-WACABAllPolicies ##
function Remove-WACABSSBPolicy {
<#

.SYNOPSIS
 Deletes SSB policy backup data from Azure Backup agent.

.DESCRIPTION
 Deletes SSB policy backup data from Azure Backup agent.

.ROLE
Administrators

#>
param (
    [Parameter(Mandatory = $false)]
    [string]
    $pin
)

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $ssbexisted = $false
    $ssbdeleted = $false

    $ssbpol = Get-OBSystemStatePolicy

    if ($ssbpol) {
        $ssbexisted = $true
    }

    if ($ssbpol) {
        $ssbans = Remove-OBSystemStatePolicy -Policy $ssbpol -SecurityPIN $pin -Confirm:$false -DeleteBackup:$true
        if ($ssbans) {
            $ssbdeleted = $true
        }
        else {
            $ssbdeleted = $false
        }
    }

    $props = @{
        ssbdeleted = $ssbdeleted
        ssbexisted = $ssbexisted
    }

    $deletiondetails = New-Object PSObject
    Add-Member -InputObject $deletiondetails -MemberType NoteProperty -Name "deletiondetails" -Value $props
    $deletiondetails

}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Remove-WACABSSBPolicy ##
function Resume-WACABFileFolderPolicy {
<#

.SYNOPSIS
To resume the FileFolder policy

.DESCRIPTION
To resume the FileFolder policy

.ROLE
Administrators

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $fileexisted = $false
    $fileresumed = $false

    $pol = Get-OBPolicy

    if ($pol) {
        $fileexisted = $true
    }

    if ($pol) {
        $ans1 = Set-OBPolicyState $pol -Confirm:$false -State Valid -ErrorAction SilentlyContinue
        $ans2 = Set-OBPolicy $pol -Confirm:$false -ErrorAction SilentlyContinue
        if ($ans1 -and $ans2) {
            $fileresumed = $true
        }
        else {
            $fileresumed = $false
        }
    }

    $props = @{
        fileresumed = $fileresumed
        fileexisted = $fileexisted
    }

    $resumedetails = New-Object PSObject
    Add-Member -InputObject $resumedetails -MemberType NoteProperty -Name "resumedetails" -Value $props
    $resumedetails
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Resume-WACABFileFolderPolicy ##
function Resume-WACABSystemStatePolicy {
<#

.SYNOPSIS
To resume the system state policy

.DESCRIPTION
To resume the system state policy

.ROLE
Administrators

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $ssbexisted = $false
    $ssbresumed = $false

    $pol = Get-OBSystemStatePolicy

    if ($pol) {
        $ssbexisted = $true
    }

    if ($pol) {
        $ans1 = Set-OBPolicyState $pol -Confirm:$false -State Valid -ErrorAction SilentlyContinue
        $ans2 = Set-OBSystemStatePolicy $pol -Confirm:$false -ErrorAction SilentlyContinue
        if ($ans1 -and $ans2) {
            $ssbresumed = $true
        }
        else {
            $ssbresumed = $false
        }
    }

    $props = @{
        ssbresumed = $ssbresumed
        ssbexisted = $ssbexisted
    }

    $resumedetails = New-Object PSObject
    Add-Member -InputObject $resumedetails -MemberType NoteProperty -Name "resumedetails" -Value $props
    $resumedetails
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Resume-WACABSystemStatePolicy ##
function Set-WACABFileFolderPolicy {
<#

.SYNOPSIS
 Modify file folder policy

.DESCRIPTION
 Modify file folder policy

.ROLE
Administrators

#>
param (
    [Parameter(Mandatory = $true)]
    [string[]]
    $filePath,
    [Parameter(Mandatory = $true)]
    [string[]]
    $daysOfWeek,
    [Parameter(Mandatory = $true)]
    [string[]]
    $timesOfDay,
    [Parameter(Mandatory = $true)]
    [int]
    $weeklyFrequency,

    [Parameter(Mandatory = $false)]
    [int]
    $retentionDays,

    [Parameter(Mandatory = $false)]
    [Boolean]
    $retentionWeeklyPolicy,
    [Parameter(Mandatory = $false)]
    [int]
    $retentionWeeks,

    [Parameter(Mandatory = $false)]
    [Boolean]
    $retentionMonthlyPolicy,
    [Parameter(Mandatory = $false)]
    [int]
    $retentionMonths,

    [Parameter(Mandatory = $false)]
    [Boolean]
    $retentionYearlyPolicy,
    [Parameter(Mandatory = $false)]
    [int]
    $retentionYears
)
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $timesOfDaySchedule = @()
    foreach ($time in $timesOfDay) {
        $timesOfDaySchedule += ([TimeSpan]$time)
    }
    $daysOfWeekSchedule = @()
    foreach ($day in $daysOfWeek) {
        $daysOfWeekSchedule += ([System.DayOfWeek]$day)
    }

    $schedule = New-OBSchedule -DaysOfWeek $daysOfWeekSchedule -TimesOfDay $timesOfDaySchedule -WeeklyFrequency $weeklyFrequency
    if ($daysOfWeekSchedule.Count -eq 7) {
        if ($retentionWeeklyPolicy -and $retentionMonthlyPolicy -and $retentionYearlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionDays $retentionDays -RetentionWeeklyPolicy:$true -WeekDaysOfWeek $daysOfWeekSchedule -WeekTimesOfDay $timesOfDaySchedule -RetentionWeeks $retentionWeeks -RetentionMonthlyPolicy:$true -MonthDaysOfWeek $daysOfWeekSchedule -MonthTimesOfDay $timesOfDaySchedule -RetentionMonths $retentionMonths -RetentionYearlyPolicy:$true -YearDaysOfWeek $daysOfWeekSchedule -YearTimesOfDay $timesOfDaySchedule -RetentionYears $retentionYears
        }
        elseif ($retentionWeeklyPolicy -and $retentionMonthlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionDays $retentionDays -RetentionWeeklyPolicy:$true -WeekDaysOfWeek $daysOfWeekSchedule -WeekTimesOfDay $timesOfDaySchedule -RetentionWeeks $retentionWeeks -RetentionMonthlyPolicy:$true -MonthDaysOfWeek $daysOfWeekSchedule -MonthTimesOfDay $timesOfDaySchedule -RetentionMonths $retentionMonths
        }
        elseif ($retentionWeeklyPolicy -and $retentionYearlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionDays $retentionDays -RetentionWeeklyPolicy:$true -WeekDaysOfWeek $daysOfWeekSchedule -WeekTimesOfDay $timesOfDaySchedule -RetentionWeeks $retentionWeeks -RetentionYearlyPolicy:$true -YearDaysOfWeek $daysOfWeekSchedule -YearTimesOfDay $timesOfDaySchedule -RetentionYears $retentionYears
        }
        elseif ($retentionYearlyPolicy -and $retentionMonthlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionDays $retentionDays -RetentionMonthlyPolicy:$true -MonthDaysOfWeek $daysOfWeekSchedule -MonthTimesOfDay $timesOfDaySchedule -RetentionMonths $retentionMonths -RetentionYearlyPolicy:$true -YearDaysOfWeek $daysOfWeekSchedule -YearTimesOfDay $timesOfDaySchedule -RetentionYears $retentionYears
        }
        elseif ($retentionWeeklyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionDays $retentionDays -RetentionWeeklyPolicy:$true -WeekDaysOfWeek $daysOfWeekSchedule -WeekTimesOfDay $timesOfDaySchedule -RetentionWeeks $retentionWeeks
        }
        elseif ($retentionMonthlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionDays $retentionDays -RetentionMonthlyPolicy:$true -MonthDaysOfWeek $daysOfWeekSchedule -MonthTimesOfDay $timesOfDaySchedule -RetentionMonths $retentionMonths
        }
        elseif ($retentionYearlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionDays $retentionDays -RetentionYearlyPolicy:$true -YearDaysOfWeek $daysOfWeekSchedule -YearTimesOfDay $timesOfDaySchedule -RetentionYears $retentionYears
        }
        else {
            $retention = New-OBRetentionPolicy -RetentionDays $retentionDays
        }
    }
    else {
        if ($retentionWeeklyPolicy -and $retentionMonthlyPolicy -and $retentionYearlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionWeeklyPolicy:$true -WeekDaysOfWeek $daysOfWeekSchedule -WeekTimesOfDay $timesOfDaySchedule -RetentionWeeks $retentionWeeks -RetentionMonthlyPolicy:$true -MonthDaysOfWeek $daysOfWeekSchedule -MonthTimesOfDay $timesOfDaySchedule -RetentionMonths $retentionMonths -RetentionYearlyPolicy:$true -YearDaysOfWeek $daysOfWeekSchedule -YearTimesOfDay $timesOfDaySchedule -RetentionYears $retentionYears
        }
        elseif ($retentionWeeklyPolicy -and $retentionMonthlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionWeeklyPolicy:$true -WeekDaysOfWeek $daysOfWeekSchedule -WeekTimesOfDay $timesOfDaySchedule -RetentionWeeks $retentionWeeks -RetentionMonthlyPolicy:$true -MonthDaysOfWeek $daysOfWeekSchedule -MonthTimesOfDay $timesOfDaySchedule -RetentionMonths $retentionMonths
        }
        elseif ($retentionWeeklyPolicy -and $retentionYearlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionWeeklyPolicy:$true -WeekDaysOfWeek $daysOfWeekSchedule -WeekTimesOfDay $timesOfDaySchedule -RetentionWeeks $retentionWeeks -RetentionYearlyPolicy:$true -YearDaysOfWeek $daysOfWeekSchedule -YearTimesOfDay $timesOfDaySchedule -RetentionYears $retentionYears
        }
        elseif ($retentionYearlyPolicy -and $retentionMonthlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionMonthlyPolicy:$true -MonthDaysOfWeek $daysOfWeekSchedule -MonthTimesOfDay $timesOfDaySchedule -RetentionMonths $retentionMonths -RetentionYearlyPolicy:$true -YearDaysOfWeek $daysOfWeekSchedule -YearTimesOfDay $timesOfDaySchedule -RetentionYears $retentionYears
        }
        elseif ($retentionWeeklyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionWeeklyPolicy:$true -WeekDaysOfWeek $daysOfWeekSchedule -WeekTimesOfDay $timesOfDaySchedule -RetentionWeeks $retentionWeeks
        }
        elseif ($retentionMonthlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionMonthlyPolicy:$true -MonthDaysOfWeek $daysOfWeekSchedule -MonthTimesOfDay $timesOfDaySchedule -RetentionMonths $retentionMonths
        }
        elseif ($retentionYearlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionYearlyPolicy:$true -YearDaysOfWeek $daysOfWeekSchedule -YearTimesOfDay $timesOfDaySchedule -RetentionYears $retentionYears
        }
    }


    $oldPolicy = Get-OBPolicy
    if ($oldPolicy) {
        $ospec = Get-OBFileSpec $oldPolicy

        $p = Remove-OBFileSpec -FileSpec $ospec -Policy $oldPolicy -Confirm:$false

        $fileSpec = New-OBFileSpec -FileSpec $filePath

        Add-OBFileSpec -Policy $p -FileSpec $fileSpec -Confirm:$false
        Set-OBSchedule -Policy $p -Schedule $schedule -Confirm:$false
        Set-OBRetentionPolicy -Policy $p -RetentionPolicy $retention -Confirm:$false
        Set-OBPolicy -Policy $p -Confirm:$false
        $p
    }
    else {
        $policy = New-OBPolicy
        $fileSpec = New-OBFileSpec -FileSpec $filePath
        Add-OBFileSpec -Policy $policy -FileSpec $fileSpec
        Set-OBSchedule -Policy $policy -Schedule $schedule
        Set-OBRetentionPolicy -Policy $policy -RetentionPolicy $retention
        Set-OBPolicy -Policy $policy -Confirm:$false
    }
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Set-WACABFileFolderPolicy ##
function Set-WACABMARSAgent {
<#

.SYNOPSIS
Sets MARS agent

.DESCRIPTION
Sets MARS agent

.ROLE
Administrators

#>
Set-StrictMode -Version 5.0
$ErrorActionPreference = "Stop"
Try {
    $agentPath = $env:TEMP + '\MARSAgentInstaller.exe'
    Invoke-WebRequest -Uri 'https://aka.ms/azurebackup_agent' -OutFile $agentPath
    & $agentPath /q | out-null

    $env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
    $azureBackupModuleName = 'MSOnlineBackup'
    $azureBackupModule = Get-Module -ListAvailable -Name $azureBackupModuleName
    if ($azureBackupModule) {
        $true
    }
    else {
        $false
    }
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}
}
## [END] Set-WACABMARSAgent ##
function Set-WACABSystemStatePolicy {
<#

.SYNOPSIS
Modify system state policy

.DESCRIPTION
Modify system state policy

.ROLE
Administrators

#>
param (
    [Parameter(Mandatory = $true)]
    [string[]]
    $daysOfWeek,
    [Parameter(Mandatory = $true)]
    [string[]]
    $timesOfDay,
    [Parameter(Mandatory = $true)]
    [int]
    $weeklyFrequency,

    [Parameter(Mandatory = $false)]
    [int]
    $retentionDays,

    [Parameter(Mandatory = $false)]
    [Boolean]
    $retentionWeeklyPolicy,
    [Parameter(Mandatory = $false)]
    [int]
    $retentionWeeks,

    [Parameter(Mandatory = $false)]
    [Boolean]
    $retentionMonthlyPolicy,
    [Parameter(Mandatory = $false)]
    [int]
    $retentionMonths
)
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $oldPolicy = Get-OBSystemStatePolicy
    if ($oldPolicy) {
        return
    }
    $policy = New-OBPolicy
    $policy = Add-OBSystemState -Policy $policy

    $timesOfDaySchedule = @()
    foreach ($time in $timesOfDay) {
        $timesOfDaySchedule += ([TimeSpan]$time)
    }
    $daysOfWeekSchedule = @()
    foreach ($day in $daysOfWeek) {
        $daysOfWeekSchedule += ([System.DayOfWeek]$day)
    }

    $schedule = New-OBSchedule -DaysOfWeek $daysOfWeekSchedule -TimesOfDay $timesOfDaySchedule -WeeklyFrequency $weeklyFrequency
    if ($daysOfWeekSchedule.Count -eq 7) {
        if ($retentionWeeklyPolicy -and $retentionMonthlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionDays $retentionDays -RetentionWeeklyPolicy:$true -WeekDaysOfWeek $daysOfWeekSchedule -WeekTimesOfDay $timesOfDaySchedule -RetentionWeeks $retentionWeeks -RetentionMonthlyPolicy:$true -MonthDaysOfWeek $daysOfWeekSchedule -MonthTimesOfDay $timesOfDaySchedule -RetentionMonths $retentionMonths
        }
        elseif ($retentionWeeklyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionDays $retentionDays -RetentionWeeklyPolicy:$true -WeekDaysOfWeek $daysOfWeekSchedule -WeekTimesOfDay $timesOfDaySchedule -RetentionWeeks $retentionWeeks
        }
        elseif ($retentionMonthlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionDays $retentionDays -RetentionMonthlyPolicy:$true -MonthDaysOfWeek $daysOfWeekSchedule -MonthTimesOfDay $timesOfDaySchedule -RetentionMonths $retentionMonths
        }
        else {
            $retention = New-OBRetentionPolicy -RetentionDays $retentionDays
        }
    }
    else {
        if ($retentionWeeklyPolicy -and $retentionMonthlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionWeeklyPolicy:$true -WeekDaysOfWeek $daysOfWeekSchedule -WeekTimesOfDay $timesOfDaySchedule -RetentionWeeks $retentionWeeks -RetentionMonthlyPolicy:$true -MonthDaysOfWeek $daysOfWeekSchedule -MonthTimesOfDay $timesOfDaySchedule -RetentionMonths $retentionMonths
        }
        elseif ($retentionWeeklyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionWeeklyPolicy:$true -WeekDaysOfWeek $daysOfWeekSchedule -WeekTimesOfDay $timesOfDaySchedule -RetentionWeeks $retentionWeeks
        }
        elseif ($retentionMonthlyPolicy) {
            $retention = New-OBRetentionPolicy -RetentionMonthlyPolicy:$true -MonthDaysOfWeek $daysOfWeekSchedule -MonthTimesOfDay $timesOfDaySchedule -RetentionMonths $retentionMonths
        }
    }
    Set-OBSchedule -Policy $policy -Schedule $schedule
    Set-OBRetentionPolicy -Policy $policy -RetentionPolicy $retention
    Set-OBSystemStatePolicy -Policy $policy -Confirm:$false
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Set-WACABSystemStatePolicy ##
function Start-WACABFileFolderBackup {
<#

.SYNOPSIS
Starts file folder backup

.DESCRIPTION
Starts file folder backup

.ROLE
Administrators

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    Get-OBPolicy | Start-OBBackup
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}
}
## [END] Start-WACABFileFolderBackup ##
function Start-WACABRecoveryMount {
<#

.SYNOPSIS
 Start the recovery job.

.DESCRIPTION
 Start the recovery job.

.ROLE
Administrators

#>
param (
    [Parameter(Mandatory = $true)]
    [int]
    $sourcePosition,
    [Parameter(Mandatory = $true)]
    [int]
    $itemPosition
)


Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {

    $job = Get-OBJob
    $driveLetter = ''
    $timeout = new-timespan -Minutes 180
    $preempt = new-timespan -Minutes 29

    $recoveryStarted = $false
    $diskmounted = $false

    if ($job) {
        $recoveryStarted = $false
    }
    else {
        $sources = Get-OBRecoverableSource
        $items = Get-OBRecoverableItem $sources[$sourcePosition]
        $recoveryJob = Start-OBRecoveryMount $items[$itemPosition] -Async
        $stopWatch = [diagnostics.stopwatch]::StartNew()
        $recoveryStarted = $true
        do {
            Start-sleep -seconds 10
            $job = Get-OBJob
            if ($job -and $job.jobstatus -and $job.jobstatus.datasourcestatus[0]) {
                $driveLetter = $job.jobstatus.datasourcestatus[0].driveletter
            }
            if ([char]::IsLetter($driveLetter)) {
                $diskmounted = $true
                break
            }
            if ($stopWatch.elapsed -gt $timeout -or
                (($stopWatch.elapsed -gt $preempt -and $job.JobStatus.DatasourceStatus -and $Job.JobStatus.DatasourceStatus.length -gt 0) -and $job.JobStatus.DatasourceStatus[0].ByteProgress.Progress -lt 64 * 1024)) {
                break
            }
        } while ($true)
    }

    $props = @{
        recoveryStarted = $recoveryStarted
        diskmounted     = $diskmounted
    }

    $recoverystatus = New-Object PSObject
    Add-Member -InputObject $recoverystatus -MemberType NoteProperty -Name "recoverystatus" -Value $props
    $recoverystatus
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Start-WACABRecoveryMount ##
function Start-WACABSystemStateBackup {
<#

.SYNOPSIS
Starts system state backup

.DESCRIPTION
Starts system state backup

.ROLE
Administrators

#>

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    Start-OBSystemStateBackup
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}
}
## [END] Start-WACABSystemStateBackup ##
function Stop-WACABBackupJob {
<#

.SYNOPSIS
Stops backup job

.DESCRIPTION
Stops backup job

.ROLE
Administrators

#>

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $job = Get-OBJob
    if ($job -ne $NULL -and $job.JobType -eq "Backup") {
        Stop-OBJob -Job $job -Confirm:$false
        return $true
    }
    else {
        return $false;
    }
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Stop-WACABBackupJob ##
function Stop-WACABRecoveryJob {
<#

.SYNOPSIS
Stops the currently ongoing backup job on the target

.DESCRIPTION
Stops the currently ongoing backup job on the target

.ROLE
Administrators

#>

Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $job = Get-OBJob
    if ($job -ne $NULL -and $job.JobType.toString() -eq "Recovery") {
        Stop-OBJob -Job $job -Confirm:$false
        return $true
    }
    else {
        return $false;
    }
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Stop-WACABRecoveryJob ##
function Suspend-WACABFileFolderPolicy {
<#

.SYNOPSIS
To pause FileFolder policy

.DESCRIPTION
To pause FileFolder policy

.ROLE
Administrators

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $fileexisted = $false
    $filepaused = $false

    $pol = Get-OBPolicy

    if ($pol) {
        $fileexisted = $true
    }

    if ($pol) {
        $ans1 = Set-OBPolicyState $pol -Confirm:$false -State Paused -ErrorAction SilentlyContinue
        $ans2 = Set-OBPolicy $pol -Confirm:$false -ErrorAction SilentlyContinue
        if ($ans1 -and $ans2) {
            $filepaused = $true
        }
        else {
            $filepaused = $false
        }
    }

    $props = @{
        filepaused  = $filepaused
        fileexisted = $fileexisted
    }

    $pausedetails = New-Object PSObject
    Add-Member -InputObject $pausedetails -MemberType NoteProperty -Name "pausedetails" -Value $props
    $pausedetails
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Suspend-WACABFileFolderPolicy ##
function Suspend-WACABSystemStatePolicy {
<#

.SYNOPSIS
To pause the system state policy

.DESCRIPTION
To pause the system state policy

.ROLE
Administrators

#>
Set-StrictMode -Version 5.0
$env:PSModulePath = (Get-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Session Manager\Environment' -Name PSModulePath).PSModulePath
Import-Module MSOnlineBackup
$ErrorActionPreference = "Stop"
Try {
    $ssbexisted = $false
    $ssbpaused = $false

    $pol = Get-OBSystemStatePolicy

    if ($pol) {
        $ssbexisted = $true
    }

    if ($pol) {
        $ans1 = Set-OBPolicyState $pol -Confirm:$false -State Paused -ErrorAction SilentlyContinue
        $ans2 = Set-OBSystemStatePolicy $pol -Confirm:$false -ErrorAction SilentlyContinue
        if ($ans1 -and $ans2) {
            $ssbpaused = $true
        }
        else {
            $ssbpaused = $false
        }
    }

    $props = @{
        ssbpaused  = $ssbpaused
        ssbexisted = $ssbexisted
    }

    $pausedetails = New-Object PSObject
    Add-Member -InputObject $pausedetails -MemberType NoteProperty -Name "pausedetails" -Value $props
    $pausedetails
}
Catch {
    if ($error[0].ErrorDetails) {
        throw $error[0].ErrorDetails
    }
    throw $error[0]
}

}
## [END] Suspend-WACABSystemStatePolicy ##
function Get-WACABCimWin32LogicalDisk {
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
## [END] Get-WACABCimWin32LogicalDisk ##
function Get-WACABCimWin32NetworkAdapter {
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
## [END] Get-WACABCimWin32NetworkAdapter ##
function Get-WACABCimWin32PhysicalMemory {
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
## [END] Get-WACABCimWin32PhysicalMemory ##
function Get-WACABCimWin32Processor {
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
## [END] Get-WACABCimWin32Processor ##
function Get-WACABClusterInventory {
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
    return $null -ne (Get-StorageSubSystem clus* | Get-StorageHealthSetting -Name "System.PerformanceHistory.Path")
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
## [END] Get-WACABClusterInventory ##
function Get-WACABClusterNodes {
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
## [END] Get-WACABClusterNodes ##
function Get-WACABDecryptedDataFromNode {
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
## [END] Get-WACABDecryptedDataFromNode ##
function Get-WACABEncryptionJWKOnNode {
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
## [END] Get-WACABEncryptionJWKOnNode ##
function Get-WACABServerInventory {
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
Checks if a system lockdown policy is enforced on this computer.

.DESCRIPTION
By checking the system lockdown policy, we can infer that PowerShell is in ConstrainedLanguage mode (WDAC).
Note: $ExecutionContext.SessionState.LanguageMode should not be used within a script context as in the case of allowed scripts (by the WDAC policy),
being executed locally, the language mode will always be FullLanguage and does NOT reflect the default system lockdown policy/language mode.

#>
function isSystemLockdownPolicyEnforced() {
  return [System.Management.Automation.Security.SystemPolicy]::GetSystemLockdownPolicy() -eq [System.Management.Automation.Security.SystemEnforcementMode]::Enforce
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
$isSystemLockdownPolicyEnforced = isSystemLockdownPolicyEnforced
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
$result | Add-Member -MemberType NoteProperty -Name 'IsSystemLockdownPolicyEnforced' -Value $isSystemLockdownPolicyEnforced
$result | Add-Member -MemberType NoteProperty -Name 'IsHciServer' -Value $isHciServer

$result

}
## [END] Get-WACABServerInventory ##

# SIG # Begin signature block
# MIIoOQYJKoZIhvcNAQcCoIIoKjCCKCYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAf6gHwdeE/eD8E
# WvttWh1YNo76Cey8bRilra9kjPqiw6CCDYUwggYDMIID66ADAgECAhMzAAADTU6R
# phoosHiPAAAAAANNMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMwMzE2MTg0MzI4WhcNMjQwMzE0MTg0MzI4WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDUKPcKGVa6cboGQU03ONbUKyl4WpH6Q2Xo9cP3RhXTOa6C6THltd2RfnjlUQG+
# Mwoy93iGmGKEMF/jyO2XdiwMP427j90C/PMY/d5vY31sx+udtbif7GCJ7jJ1vLzd
# j28zV4r0FGG6yEv+tUNelTIsFmmSb0FUiJtU4r5sfCThvg8dI/F9Hh6xMZoVti+k
# bVla+hlG8bf4s00VTw4uAZhjGTFCYFRytKJ3/mteg2qnwvHDOgV7QSdV5dWdd0+x
# zcuG0qgd3oCCAjH8ZmjmowkHUe4dUmbcZfXsgWlOfc6DG7JS+DeJak1DvabamYqH
# g1AUeZ0+skpkwrKwXTFwBRltAgMBAAGjggGCMIIBfjAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUId2Img2Sp05U6XI04jli2KohL+8w
# VAYDVR0RBE0wS6RJMEcxLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJh
# dGlvbnMgTGltaXRlZDEWMBQGA1UEBRMNMjMwMDEyKzUwMDUxNzAfBgNVHSMEGDAW
# gBRIbmTlUAXTgqoXNzcitW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIw
# MTEtMDctMDguY3JsMGEGCCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDEx
# XzIwMTEtMDctMDguY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIB
# ACMET8WuzLrDwexuTUZe9v2xrW8WGUPRQVmyJ1b/BzKYBZ5aU4Qvh5LzZe9jOExD
# YUlKb/Y73lqIIfUcEO/6W3b+7t1P9m9M1xPrZv5cfnSCguooPDq4rQe/iCdNDwHT
# 6XYW6yetxTJMOo4tUDbSS0YiZr7Mab2wkjgNFa0jRFheS9daTS1oJ/z5bNlGinxq
# 2v8azSP/GcH/t8eTrHQfcax3WbPELoGHIbryrSUaOCphsnCNUqUN5FbEMlat5MuY
# 94rGMJnq1IEd6S8ngK6C8E9SWpGEO3NDa0NlAViorpGfI0NYIbdynyOB846aWAjN
# fgThIcdzdWFvAl/6ktWXLETn8u/lYQyWGmul3yz+w06puIPD9p4KPiWBkCesKDHv
# XLrT3BbLZ8dKqSOV8DtzLFAfc9qAsNiG8EoathluJBsbyFbpebadKlErFidAX8KE
# usk8htHqiSkNxydamL/tKfx3V/vDAoQE59ysv4r3pE+zdyfMairvkFNNw7cPn1kH
# Gcww9dFSY2QwAxhMzmoM0G+M+YvBnBu5wjfxNrMRilRbxM6Cj9hKFh0YTwba6M7z
# ntHHpX3d+nabjFm/TnMRROOgIXJzYbzKKaO2g1kWeyG2QtvIR147zlrbQD4X10Ab
# rRg9CpwW7xYxywezj+iNAc+QmFzR94dzJkEPUSCJPsTFMIIHejCCBWKgAwIBAgIK
# YQ6Q0gAAAAAAAzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEw
# OTA5WjB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYD
# VQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+la
# UKq4BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc
# 6Whe0t+bU7IKLMOv2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4D
# dato88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+
# lD3v++MrWhAfTVYoonpy4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nk
# kDstrjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6
# A4aN91/w0FK/jJSHvMAhdCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmd
# X4jiJV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL
# 5zmhD+kjSbwYuER8ReTBw3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zd
# sGbiwZeBe+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3
# T8HhhUSJxAlMxdSlQy90lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS
# 4NaIjAsCAwEAAaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRI
# bmTlUAXTgqoXNzcitW2oynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTAL
# BgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBD
# uRQFTuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jv
# c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEF
# BQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1h
# cnljcHMuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkA
# YwB5AF8AcwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn
# 8oalmOBUeRou09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7
# v0epo/Np22O/IjWll11lhJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0b
# pdS1HXeUOeLpZMlEPXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/
# KmtYSWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvy
# CInWH8MyGOLwxS3OW560STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBp
# mLJZiWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJi
# hsMdYzaXht/a8/jyFqGaJ+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYb
# BL7fQccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbS
# oqKfenoi+kiVH6v7RyOA9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sL
# gOppO6/8MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtX
# cVZOSEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGgowghoGAgEBMIGVMH4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01p
# Y3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAANNTpGmGiiweI8AAAAA
# A00wDQYJYIZIAWUDBAIBBQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIGyc
# 0WkFy2lW6LBwgwX17bTdmjJ5rJYVgHqu7Jtm8pZEMEIGCisGAQQBgjcCAQwxNDAy
# oBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20wDQYJKoZIhvcNAQEBBQAEggEABjFVZtNsGyYiNBJ6NCrSsAdEGDIOstZy2Z/G
# 6gnmLRgUFMYDCX8QwrAhnkorbyERrYXUt2VrYQPv1hZLrGHXG4liYI9fb8e7Sn2B
# 5sOjqfQmCNepopinye5D1wc8704z9ryYMmhhm0iXy4lgNNTRNPdvjHnBkntvQiBh
# +bA4pgB6ey0ebJ2gmioD0+hw+NIWq1NomrYbZuAj3nj07czdkjecmxAE5lKZCux/
# p1ZRENDAlGO4X/XgpInfP7yioNpLGcU/8QRC4Jk5pc7WcdLFp2uk97LwMZP/sa5D
# 80hGPtGI4mm6soOGSWkEMwBTh6B0LG3pKosbT5g9LRvNbPlcAKGCF5QwgheQBgor
# BgEEAYI3AwMBMYIXgDCCF3wGCSqGSIb3DQEHAqCCF20wghdpAgEDMQ8wDQYJYIZI
# AWUDBAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYBBAGE
# WQoDATAxMA0GCWCGSAFlAwQCAQUABCDWUbzsEnGzZCtAvMTngl8pwgplxa3J7DVV
# 6MxwKaGmvgIGZSi09tNlGBMyMDIzMTEwMzA2MTc0MS43NjVaMASAAgH0oIHRpIHO
# MIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQL
# ExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxk
# IFRTUyBFU046MzMwMy0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1l
# LVN0YW1wIFNlcnZpY2WgghHqMIIHIDCCBQigAwIBAgITMwAAAcyGpdw369lhLQAB
# AAABzDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAx
# MDAeFw0yMzA1MjUxOTEyMDFaFw0yNDAyMDExOTEyMDFaMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046MzMwMy0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Uw
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDMsSIF8e9NmEc+83NVZGgW
# WZi/wBYt8zhxAfSGM7xw7K7CbA/1A4GhovPvkIY873tnyzdyZe+6YHXx+Rd618lQ
# Dmmm5X4euiYG53Ld7WIK+Dd+hyi0H97D6HM4ZzGqovmwB0fZ3lh+phJLoPT+9yrT
# LFzkkKw2Vcb7wXMBziD0MVVYbmwRlRaypTntl39IENCEijW9j6MElTyXP2zrc0Ot
# hQN5RrMTY5iZja3MyHCFmYMGinmHftsaG3Ydi8Ga8BQjdtoTm5dVhnqs2qKNEOqZ
# Son28R4Xff0tlJL5UHyI3bywH/+zQeJu8qnsSCi8VFPOsZEb6cZzhXHaAiSGtdKA
# bQRaAIhExbIUpeJypC7l+wqKC3BO9ADGupB9ZgUFbSv5ECFjMDzbfm8M5zz2A4xY
# NPQXqZv0wGWL+jTvb7kFYiDPPe+zRyBbzmrSpObB7XqjqzUFNKlwp+Mx15k1F7FM
# s5EM2uG68IQsdAGBkZbSDmuGmjPbZ7dtim+XHuh3NS6JmXYPS7rikpCbUsMZMn5e
# WxiWFIk6f00skR4RLWmh0N6Oq+KYI1fA59LzGiAbOrcxgvQkRo3OD4o1JW9z1TNM
# wEbkzPrXMo8rrGsuGoyYWcsm9xhd0GXIRHHC64nzbI3e0G5jqEsWQc4uaQeSRyr7
# 0KRijzVyWjjYfsEtvVMlJwIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFIKmHGRdPIdL
# RXtsR5XRSyM3+2kMMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1GelyMF8G
# A1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNybDBs
# BggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUy
# MDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
# AwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQB5GUMo9XviUl3g
# 72u8oQTorIKDoAdgWZ4LQ9+dAEQCmaetsThkxbNm15seu7GmwpZdhMQN8TNddGki
# 5s5Ie+aA2VEo9vZz31llusHBXAVrQtpufQqtIA+2nnusfaYviitr6p5kVT609LIT
# OYgdKRWEpfx/4yT5R9yMeKxoxkk8tyGiGPZK40ST5Z14OPdJfVbkYeCvlLQclsX1
# +WBZNx/XZvazJmXjvYjTuG0QbZpxw4ZO3ZoffQYxZYRzn0z41U7MDFlXo2ihfasd
# bHuua6kpHxJ9AIoUevh3mzvUxYp0u0z3wYDPpLuo+M2VYh8XOCUB0u75xG3S5+98
# TKmFbqZYgpgr6P+YKeao2YpB1izs850YSzuwaX7kRxAURlmN/j5Hv4wabnOfZb36
# mDqJp4IeGmwPtwI8tEPsuRAmyreejyhkZV7dfgJ4N83QBhpHVZlB4FmlJR8yF3aB
# 15QW6tw4CaH+PMIDud6GeOJO4cQE+lTc6rIJmN4cfi2TTG7e49TvhCXfBS2pzOyb
# 9YemSm0krk8jJh6zgeGqztk7zewfE+3shQRc74sXLY58pvVoznfgfGvy1llbq4Oe
# y96KouwiuhDtxuKlTnW7pw7xaNPhIMsOxW8dpSp915FtKfOqKR/dfJOsbHDSJY/i
# iJz4mWKAGoydeLM6zLmohRCPWk/Q5jCCB3EwggVZoAMCAQICEzMAAAAVxedrngKb
# SZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRpZmlj
# YXRlIEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIy
# NVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UE
# AxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qlsTnXI
# yjVX9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLAEBjo
# YH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrEqv1y
# aa8dq6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v
# 3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1oO5pG
# ve2krnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viS
# kR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2TPYr
# bqgSUei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlM
# jgK8QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJNmSL
# W6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AF
# emzFER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+auIu
# rQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3FQIE
# FgQUKqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl0mWn
# G1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUHAgEW
# M2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5
# Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMAdQBi
# AEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV
# 9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3Js
# Lm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAx
# MC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2
# LTIzLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL/Klv
# 6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZn
# OlNN3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5tggz1
# bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfgQJY4
# rPf5KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU
# 6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCrdTDF
# NLB62FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/
# HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2tVdU
# CbFpAUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKi
# excdFYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9JZTm
# dHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZq
# ELQdVTNYs6FwZvKhggNNMIICNQIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJp
# Y2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjMzMDMtMDVF
# MC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMK
# AQEwBwYFKw4DAhoDFQBOTuZ3uYfiihS4zRToxisDt9mJpKCBgzCBgKR+MHwxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
# c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA6O7hlDAi
# GA8yMDIzMTEwMzAzMDM0OFoYDzIwMjMxMTA0MDMwMzQ4WjB0MDoGCisGAQQBhFkK
# BAExLDAqMAoCBQDo7uGUAgEAMAcCAQACAisYMAcCAQACAhNwMAoCBQDo8DMUAgEA
# MDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAI
# AgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBADyls6r8d8ADXxQ+9+flx3xnt8sO
# /gHKYWVsupuCnclSUPQMXAHqzazKwPiGGwreUuFh+fBbLrzYV+Fbh49wmjMVnXGY
# Hi6HhOVbF04GgfbwIUcMDPLXKYwub4Og7OCG1Wn+RixWzEZxc5HVkrbsloqI1/py
# iOob6ULr1aytYbLIxgKO52HByZ1Cm9LUrKn3BARYXAjWufu2h975tfKjNo4mzpRa
# 3lFG3sqB4cgJaVWKUIdjsbY9tt7PXRVw7F9wrh/ff2GDQn+H+NoxL7HoP81sjNHG
# 9/HUjnw9ZP+xb7tWVrXfNNy4iQ5FYSCrizPKQYJiO1IA69nJ6pfGXVQuJ9QxggQN
# MIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAcyG
# pdw369lhLQABAAABzDANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0G
# CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCDhwGkWbbRar4NNOSMGRunqfrfO
# +ck8uWiIJTL5DpvHLDCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0EINbuZQHC
# /sMCE+cgKVSkwKpDICfnefFZYgDbF4HrcFjmMIGYMIGApH4wfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAHMhqXcN+vZYS0AAQAAAcwwIgQg+/p+vmlH
# yNB64i80gK3EwMqYIFuJi2Jotch6b+jCy/gwDQYJKoZIhvcNAQELBQAEggIAvEBW
# QxWwSpXeIkyJavtq00txKNd1BV6XdMNqh2GeoU7NfgSCAUXvD6/DpBICWLUtTnFg
# X6Lv65ZelpOwsB5NdYqzvRsBPAivptuebP/RBRGvY94ZMI7OZY7M4vwf7ZjMOKdk
# 8ieT7aNfjy7OPa29AoL0wPI7CdRPdu+9EytcXVe616lsubzXs/JOIPSEMr3aAq7G
# kG6+KIPV5l9aqCFn1RuIYlpjyxt7k3SqVYLoKsIyOt2Pd9qCabusxpyFUqUFWpKZ
# XdeRH/uXSPssj+uYFreLlYeiSwXLZA7cZB2xdJc4RXmeFq4ovKk05rFy1OomMzEA
# YM0rsbiiEo1mrMto3BSI7SJkpSrOWWAJg+Dl3Q3wB5kA/2uuZ3aAMDOL6Rr1EQfj
# VzjV+jI5gyXlZvrYwozStrWDYJ4d4QOa01y7T/FTs6O46rq0qS/D77clwfMuMdhm
# UCuQtjTVwSRB+0tLn98KijmS+ocQnbnx4cENQhEEzFyEUL3rgDZXKW+udn3Q40XG
# a50n3l6QDFUvVSTP2cQw+E0zhZh72plo6VbuqE9Rd3DQTBQyExs2AvBKq0zbZHz7
# 48E/fke/qTLUlCYwNWyc3TdQYyvAom73sMZxOATQxDwVY0t7KNaYnsDlcO9+5CTn
# pNt9Gw0084ZUZnyTkL9AS4WCTrj3Fu3duSOMS40=
# SIG # End signature block
