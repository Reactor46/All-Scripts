function Get-VmwOrphan{
<#
.SYNOPSIS
Find orphaned files on a datastore
.DESCRIPTION
This function will scan the complete content of a datastore.
It will then verify all registered VMs and Templates on that
datastore, and compare those files with the datastore list.
Files that are not present in a VM or Template are considered
orphaned
.NOTES
Author:  Luc Dekens
.PARAMETER Datastore
The datastore that needs to be scanned
.EXAMPLE
PS> Get-VmwOrphan -Datastore DS1
.EXAMPLE
PS> Get-Datastore -Name DS* | Get-VmwOrphan
#>
[CmdletBinding()]
param(
[parameter(Mandatory=$false,ValueFromPipeline=$true)]
[PSObject[]]$Datastore
)
Begin{
$flags = New-Object VMware.Vim.FileQueryFlags
$flags.FileOwner = $true
$flags.FileSize = $true
$flags.FileType = $true
$flags.Modification = $true
$qFloppy = New-Object VMware.Vim.FloppyImageFileQuery
$qFolder = New-Object VMware.Vim.FolderFileQuery
$qISO = New-Object VMware.Vim.IsoImageFileQuery
$qConfig = New-Object VMware.Vim.VmConfigFileQuery
$qConfig.Details = New-Object VMware.Vim.VmConfigFileQueryFlags
$qConfig.Details.ConfigVersion = $true
$qTemplate = New-Object VMware.Vim.TemplateConfigFileQuery
$qTemplate.Details = New-Object VMware.Vim.VmConfigFileQueryFlags
$qTemplate.Details.ConfigVersion = $true
$qDisk = New-Object VMware.Vim.VmDiskFileQuery
$qDisk.Details = New-Object VMware.Vim.VmDiskFileQueryFlags
$qDisk.Details.CapacityKB = $true
$qDisk.Details.DiskExtents = $true
$qDisk.Details.DiskType = $true
$qDisk.Details.HardwareVersion = $true
$qDisk.Details.Thin = $true
$qLog = New-Object VMware.Vim.VmLogFileQuery
$qRAM = New-Object VMware.Vim.VmNvramFileQuery
$qSnap = New-Object VMware.Vim.VmSnapshotFileQuery
$searchSpec = New-Object VMware.Vim.HostDatastoreBrowserSearchSpec
$searchSpec.details = $flags
$searchSpec.Query = $qFloppy,$qFolder,$qISO,$qConfig,$qTemplate,$qDisk,$qLog,$qRAM,$qSnap
$searchSpec.sortFoldersFirst = $true
}
Process{
foreach($ds in $Datastore){
if($ds.GetType().Name -eq "String"){
$ds = Get-Datastore -Name $ds
}
# Only shared VMFS datastore
if($ds.Type -eq "VMFS" -and $ds.ExtensionData.Summary.MultipleHostAccess -and $ds.Accessible){
Write-Verbose -Message "$(Get-Date)`t$((Get-PSCallStack)[0].Command)`tLooking at $($ds.Name)"
# Define file DB
$fileTab = @{}
# Get datastore files
$dsBrowser = Get-View -Id $ds.ExtensionData.browser
$rootPath = "[" + $ds.Name + "]"
$searchResult = $dsBrowser.SearchDatastoreSubFolders($rootPath, $searchSpec) | Sort-Object -Property {$_.FolderPath.Length}
foreach($folder in $searchResult){
foreach ($file in $folder.File){
$key = "$($folder.FolderPath)$(if($folder.FolderPath[-1] -eq ']'){' '})$($file.Path)"
$fileTab.Add($key,$file)
$folderKey = "$($folder.FolderPath.TrimEnd('/'))"
if($fileTab.ContainsKey($folderKey)){
$fileTab.Remove($folderKey)
}
}
}
# Get VM inventory
Get-VM -Datastore $ds | %{
$_.ExtensionData.LayoutEx.File | %{
if($fileTab.ContainsKey($_.Name)){
$fileTab.Remove($_.Name)
}
}
}
# Get Template inventory
Get-Template | where {$_.DatastoreIdList -contains $ds.Id} | %{
$_.ExtensionData.LayoutEx.File | %{
if($fileTab.ContainsKey($_.Name)){
$fileTab.Remove($_.Name)
}
}
}
# Remove system files & folders from list
$systemFiles = $fileTab.Keys | where{$_ -match "] \.|vmkdump"}
$systemFiles | %{
$fileTab.Remove($_)
}
# Organise remaining files
if($fileTab.Count){
$fileTab.GetEnumerator() | %{
$obj = [ordered]@{
Name = $_.Value.Path
Folder = $_.Name
Size = $_.Value.FileSize
CapacityKB = $_.Value.CapacityKb
Modification = $_.Value.Modification
Owner = $_.Value.Owner
Thin = $_.Value.Thin
Extents = $_.Value.DiskExtents -join ','
DiskType = $_.Value.DiskType
HWVersion = $_.Value.HardwareVersion
}
New-Object PSObject -Property $obj
}
Write-Verbose -Message "$(Get-Date)`t$((Get-PSCallStack)[0].Command)`tFound orphaned files on $($ds.Name)!"
}
else{
Write-Verbose -Message "$(Get-Date)`t$((Get-PSCallStack)[0].Command)`tNo orphaned files found on $($ds.Name)."
}
}
}
}
}


function Get-Vmalerts{
$VMs = Get-View -ViewType VirtualMachine -Property Name,OverallStatus,TriggeredAlarmstate
$FaultyVMs = $VMs | Where-Object {$_.OverallStatus -ne "Green"}
$report = @()
if ($FaultyVMs -ne $null) {
    foreach ($FaultyVM in $FaultyVMs) {
            foreach ($TriggeredAlarm in $FaultyVM.TriggeredAlarmstate) {
                $alarmID = $TriggeredAlarm.Alarm.ToString()
                $timesince =  new-timespan -end ($TriggeredAlarm.time)
                $duration = "$($timesince.days)d$($timesince.hours)h$($timesince.minutes)m$($timesince.seconds)s"
                $object = New-Object PSObject
                Add-Member -InputObject $object NoteProperty entity $FaultyVM.Name
                Add-Member -InputObject $object NoteProperty triggeredalarms ("$(Get-AlarmDefinition -Id $alarmID)")
                Add-Member -InputObject $object NoteProperty triggered $duration
                $report += $object 
            }
        }
    }
    $report | Where-Object {$_.TriggeredAlarms -ne ""}
}


function Send-VmalertstoOpBot{
$limit = (Get-date).Addminutes(-5)
$VMs = Get-View -ViewType VirtualMachine -Property Name,OverallStatus,TriggeredAlarmstate
$FaultyVMs = $VMs | Where-Object {$_.OverallStatus -ne "Green"}
 
if ($FaultyVMs -ne $null) {
    foreach ($FaultyVM in $FaultyVMs) {
            foreach ($TriggeredAlarm in $FaultyVM.TriggeredAlarmstate) {
                  if ($limit -le $TriggeredAlarmState.Time) {    
                        $alarmID = $TriggeredAlarm.Alarm.ToString()
                        $object = New-Object PSObject
                        Add-Member -InputObject $object NoteProperty entity $FaultyVM.Name
                        Add-Member -InputObject $object NoteProperty triggeredalarms ("$(Get-AlarmDefinition -Id $alarmID)")
                        $object | ConvertTo-Json | Invoke-WebRequest -uri "http://127.0.0.1:8080/opbot/vmcritical" -Method POST -ContentType "application/json" | out-null
                  }
            }
        }
    }
}

function Send-FilestoOpBot{
  [CmdletBinding()]
  Param(
   [Parameter(Mandatory=$True,Position=1)]
   [string]$filePath,
   [Parameter(Mandatory=$False)]
   [string]$channels = "#opbotuploads"
  )
  curl -F file="@$($filePath)" -F channels=$channels -F token="$($env:OPBOT_SLACK_TOKEN)" https://slack.com/api/files.upload | out-null
}

function Get-BadVms{
  get-view -ViewType VirtualMachine -Filter @{'RunTime.ConnectionState'='disconnected|inaccessible|invalid|orphaned'} | select name
}

function Get-RedHosts{
  get-view -ViewType HostSystem -Filter @{'overallStatus'='red'} | select name
}

function Get-DisconnectHosts{
  posh Get-VMHost -state Disconnected
}

function Get-RedCluster{
 get-view -ViewType ComputeResource -Filter @{'overallStatus'='red'} | select name
}

function Get-ClusterCompliance{
 $HPDetails = @()
 Foreach ($VMHost in Get-VMHost) {
 $HostProfile = $VMHost | Get-VMHostProfile
 if ($VMHost | Get-VMHostProfile) {
  $HP = $VMHost | Test-VMHostProfileCompliance
 If ($HP.ExtensionData.ComplianceStatus -eq "nonCompliant") {
 Foreach ($issue in ($HP.IncomplianceElementList)) {
 $Details = "" | Select VMHost, Compliance, HostProfile, IncomplianceDescription
 $Details.VMHost = $VMHost.Name
 $Details.Compliance = $HP.ExtensionData.ComplianceStatus
 $Details.HostProfile = $HP.VMHostProfile
 $Details.IncomplianceDescription = $Issue.Description
 $HPDetails += $Details
 }
 } Else {
 $Details = "" | Select VMHost, Compliance, HostProfile, IncomplianceDescription
 $Details.VMHost = $VMHost.Name
 $Details.Compliance = "Compliant"
 $Details.HostProfile = $HostProfile.Name
 $Details.IncomplianceDescription = ""
 $HPDetails += $Details
 }
 } Else {
 $Details = "" | Select VMHost, Compliance, HostProfile, IncomplianceDescription
 $Details.VMHost = $VMHost.Name
 $Details.Compliance = "No profile attached"
 $Details.HostProfile = ""
 $Details.IncomplianceDescription = ""
 $HPDetails += $Details
 }
 }
$HPDetails
}


function Send-VmHostsalertstoOpBot{
$limit = (Get-date).AddMinutes(-5)
$VMHosts = Get-View -ViewType HostSystem -Property Name,OverallStatus,TriggeredAlarmstate
$FaultyVMHosts = $VMHosts | Where-Object {$_.OverallStatus -ne "Green"}
 
if ($FaultyVMHosts -ne $null) {
    foreach ($FaultyVMHost in $FaultyVMHosts) {
            foreach ($TriggeredAlarm in $FaultyVMHost.TriggeredAlarmstate) {
              if ($limit -le $TriggeredAlarmState.Time) {     
                $alarmID = $TriggeredAlarm.Alarm.ToString()
                $object = New-Object PSObject
                Add-Member -InputObject $object NoteProperty entity $FaultyVMHost.Name
                Add-Member -InputObject $object NoteProperty triggeredalarms ("$(Get-AlarmDefinition -Id $alarmID)")
                $object | ConvertTo-Json | Invoke-WebRequest -uri "http://127.0.0.1:8080/opbot/vmhostcritical" -Method POST -ContentType "application/json"
               }
            }
        }
    }
}

function Get-VIEventPlus {
<#   
.SYNOPSIS  Returns vSphere events    
.DESCRIPTION The function will return vSphere events. With
    the available parameters, the execution time can be
   improved, compered to the original Get-VIEvent cmdlet. 
.NOTES  Author:  Luc Dekens   
.PARAMETER Entity
   When specified the function returns events for the
   specific vSphere entity. By default events for all
   vSphere entities are returned. 
.PARAMETER EventType
   This parameter limits the returned events to those
   specified on this parameter. 
.PARAMETER Start
   The start date of the events to retrieve 
.PARAMETER Finish
   The end date of the events to retrieve. 
.PARAMETER Recurse
   A switch indicating if the events for the children of
   the Entity will also be returned 
.PARAMETER User
   The list of usernames for which events will be returned 
.PARAMETER System
   A switch that allows the selection of all system events. 
.PARAMETER ScheduledTask
   The name of a scheduled task for which the events
   will be returned 
.PARAMETER FullMessage
   A switch indicating if the full message shall be compiled.
   This switch can improve the execution speed if the full
   message is not needed.   
.EXAMPLE
   PS> Get-VIEventPlus -Entity $vm
.EXAMPLE
   PS> Get-VIEventPlus -Entity $cluster -Recurse:$true
#>

  param(
    [VMware.VimAutomation.ViCore.Types.V1.Inventory.InventoryItem[]]$Entity,
    [string[]]$EventType,
    [DateTime]$Start,
    [DateTime]$Finish = (Get-Date),
    [switch]$Recurse,
    [string[]]$User,
    [Switch]$System,
    [string]$ScheduledTask,
    [switch]$FullMessage = $false
  )

  process {
    $eventnumber = 100
    $events = @()
    $eventMgr = Get-View EventManager
    $eventFilter = New-Object VMware.Vim.EventFilterSpec
    $eventFilter.disableFullMessage = ! $FullMessage
    $eventFilter.entity = New-Object VMware.Vim.EventFilterSpecByEntity
    $eventFilter.entity.recursion = &{if($Recurse){"all"}else{"self"}}
    $eventFilter.eventTypeId = $EventType
    if($Start -or $Finish){
      $eventFilter.time = New-Object VMware.Vim.EventFilterSpecByTime
    if($Start){
        $eventFilter.time.beginTime = $Start
    }
    if($Finish){
        $eventFilter.time.endTime = $Finish
    }
    }
  if($User -or $System){
    $eventFilter.UserName = New-Object VMware.Vim.EventFilterSpecByUsername
    if($User){
      $eventFilter.UserName.userList = $User
    }
    if($System){
      $eventFilter.UserName.systemUser = $System
    }
  }
  if($ScheduledTask){
    $si = Get-View ServiceInstance
    $schTskMgr = Get-View $si.Content.ScheduledTaskManager
    $eventFilter.ScheduledTask = Get-View $schTskMgr.ScheduledTask |
      where {$_.Info.Name -match $ScheduledTask} |
      Select -First 1 |
      Select -ExpandProperty MoRef
  }
  if(!$Entity){
    $Entity = @(Get-Folder -Name Datacenters)
  }
  $entity | %{
      $eventFilter.entity.entity = $_.ExtensionData.MoRef
      $eventCollector = Get-View ($eventMgr.CreateCollectorForEvents($eventFilter))
      $eventsBuffer = $eventCollector.ReadNextEvents($eventnumber)
      while($eventsBuffer){
        $events += $eventsBuffer
        $eventsBuffer = $eventCollector.ReadNextEvents($eventnumber)
      }
      $eventCollector.DestroyCollector()
    }
    $events
  }
}

function Get-MotionHistory {
<#   
.SYNOPSIS  Returns the vMotion/svMotion history    
.DESCRIPTION The function will return information on all
   the vMotions and svMotions that occurred over a specific
    interval for a defined number of virtual machines 
.NOTES  Author:  Luc Dekens   
.PARAMETER Entity
   The vSphere entity. This can be one more virtual machines,
   or it can be a vSphere container. If the parameter is a
    container, the function will return the history for all the
   virtual machines in that container. 
.PARAMETER Days
   An integer that indicates over how many days in the past
   the function should report on. 
.PARAMETER Hours
   An integer that indicates over how many hours in the past
   the function should report on. 
.PARAMETER Minutes
   An integer that indicates over how many minutes in the past
   the function should report on. 
.PARAMETER Sort
   An switch that indicates if the results should be returned
   in chronological order. 
.EXAMPLE
   PS> Get-MotionHistory -Entity $vm -Days 1
.EXAMPLE
   PS> Get-MotionHistory -Entity $cluster -Sort:$false
.EXAMPLE
   PS> Get-Datacenter -Name $dcName |
   >> Get-MotionHistory -Days 7 -Sort:$false
#>

  param(
    [CmdletBinding(DefaultParameterSetName="Days")]
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [VMware.VimAutomation.ViCore.Types.V1.Inventory.InventoryItem[]]$Entity,
    [Parameter(ParameterSetName='Days')]
    [int]$Days = 1,
    [Parameter(ParameterSetName='Hours')]
    [int]$Hours,
    [Parameter(ParameterSetName='Minutes')]
    [int]$Minutes,
    [switch]$Recurse = $false,
    [switch]$Sort = $true
  )

  begin{
    $history = @()
    switch($psCmdlet.ParameterSetName){
      'Days' {
        $start = (Get-Date).AddDays(- $Days)
      }
      'Hours' {
        $start = (Get-Date).AddHours(- $Hours)
      }
      'Minutes' {
        $start = (Get-Date).AddMinutes(- $Minutes)
      }
    }
    $eventTypes = "DrsVmMigratedEvent","VmMigratedEvent"
  }

  process{
    $history += Get-VIEventPlus -Entity $entity -Start $start -EventType $eventTypes -Recurse:$Recurse |
    Select CreatedTime,
    @{N="Type";E={
        if($_.SourceDatastore.Name -eq $_.Ds.Name){"vMotion"}else{"svMotion"}}},
    @{N="UserName";E={if($_.UserName){$_.UserName}else{"System"}}},
    @{N="VM";E={$_.VM.Name}},
    @{N="SrcVMHost";E={$_.SourceHost.Name.Split('.')[0]}},
    @{N="TgtVMHost";E={if($_.Host.Name -ne $_.SourceHost.Name){$_.Host.Name.Split('.')[0]}}},
    @{N="SrcDatastore";E={$_.SourceDatastore.Name}},
    @{N="TgtDatastore";E={if($_.Ds.Name -ne $_.SourceDatastore.Name){$_.Ds.Name}}}
  }

  end{
    if($Sort){
      $history | Sort-Object -Property CreatedTime
    }
    else{
      $history
    }
  }
}




#####################################                                                                
  ## http://kunaludapi.blogspot.com                                                                
  ## Version: 2                                                               
  ## Date: 16 Dec 2015                                                              
  ## Script tested on below platform                                                                
  ## 1) Powershell v4                                                               
  ## 2) Powercli v5.5                                                                
  ## 3) Vsphere 5.5                                                               
  ####################################                                                               
 
 
 function Get-DatastoreInventory {                                                               
   $HostDatastoreInfo = Get-VMHost | Get-ScsiLun -LunType disk                                                                
   $DatastoreInfo = Get-Datastore 
   foreach ($Hostdatastore in $HostDatastoreInfo) {                                                                
    $Datastore = $DatastoreInfo | Where-Object {$_.extensiondata.info.vmfs.extent.Diskname -match $Hostdatastore.CanonicalName}                                                               
    $LunPath = $Hostdatastore | Get-ScsiLunPath                                                              
    if ($Datastore.ExtensionData.vm) {                                                               
     $VMsOnDatastore = $(Get-view $Datastore.ExtensionData.vm).name -join ","                                                               
    } #if                                                               
    else {$VMsOnDatastore = "No VMs"}                                                               
                                                                 
   #Work on not assigned Luns error at silentlyContinue                                                               
    if ($Datastore.Name -eq $null) {                                                              
     $DatastoreName = "Not mapped"                                                              
     $FileSystemVersion = "Not mapped"                                                              
    }                                                              
    else {                                                              
     $DatastoreName = $Datastore.Name -join ","                                                              
     $FileSystemVersion = $Datastore[0].FileSystemVersion                                                               
    }                                                              
                                                                   
    $DatastoreFreeSpace = $Datastore.FreeSpaceGB -join ", "                                                               
    $DatastoreCapacityGB = $Datastore.CapacityGB -join ", "                                                               
    $DatastoreDatacenter = $Datastore.Datacenter -join ", "                                                               
                                                               
    $State = $LunPath.State -join ", "                                                              
    $Preferred = $LunPath.Preferred -join ", "                                                              
    $Paths = ($LunPath.ExtensionData.Transport | foreach {($_.Address -split ":")[0]}) -Join ", "                                                              
    $IsWorkingPath = $LunPath.ExtensionData.IsWorkingPath -Join ", "                                                              
                                                                  
    $Obj = New-Object PSObject                                                               
    $Obj | Add-Member -Name VMhost -MemberType NoteProperty -Value $hostdatastore.VMHost                                                               
    $Obj | Add-Member -Name DatastoreName -MemberType NoteProperty -Value $DatastoreName                                                                
    $Obj | Add-Member -Name FreeSpaceGB -MemberType NoteProperty -Value $DatastoreFreeSpace                                                               
    $Obj | Add-Member -Name CapacityGB -MemberType NoteProperty -Value $DatastoreCapacityGB                                                               
    $Obj | Add-Member -Name FileSystemVersion -MemberType NoteProperty -Value $FileSystemVersion                                                               
    $Obj | Add-Member -Name RuntimeName -MemberType NoteProperty -Value $hostdatastore.RuntimeName                                                               
    $Obj | Add-Member -Name CanonicalName -MemberType NoteProperty -Value $hostdatastore.CanonicalName                                                               
    $Obj | Add-Member -Name MultipathPolicy -MemberType NoteProperty -Value $hostdatastore.MultipathPolicy                                                               
    $Obj | Add-Member -Name Vendor -MemberType NoteProperty -Value $hostdatastore.Vendor                                                               
    $Obj | Add-Member -Name DatastoreDatacenter -MemberType NoteProperty -Value $DatastoreDatacenter                                                               
    $Obj | Add-Member -Name VMsOnDataStore -MemberType NoteProperty -Value $VMsOnDatastore                                                               
    $Obj | Add-Member -Name NumberOfPaths -MemberType NoteProperty -Value $LunPath.Count                                                              
    $Obj | Add-Member -Name Paths -MemberType NoteProperty -Value $Paths                                                              
    $Obj | Add-Member -Name State -MemberType NoteProperty -Value $State                                                              
    $Obj | Add-Member -Name Preferred -MemberType NoteProperty -Value $Preferred                                                              
    $Obj | Add-Member -Name IsWorkingPath -MemberType NoteProperty -Value $IsWorkingPath                                                                                                                            
    $Obj 
   }                                                               
 } 

 function Send-DatastoreInventory {                                                               
   $HostDatastoreInfo = Get-VMHost | Get-ScsiLun -LunType disk                                                                
   $DatastoreInfo = Get-Datastore 
   $report = @()
   foreach ($Hostdatastore in $HostDatastoreInfo) {                                                                
    $Datastore = $DatastoreInfo | Where-Object {$_.extensiondata.info.vmfs.extent.Diskname -match $Hostdatastore.CanonicalName}                                                               
    $LunPath = $Hostdatastore | Get-ScsiLunPath                                                              
    if ($Datastore.ExtensionData.vm) {                                                               
     $VMsOnDatastore = $(Get-view $Datastore.ExtensionData.vm).name -join ","                                                               
    } #if                                                               
    else {$VMsOnDatastore = "No VMs"}                                                               
                                                                 
   #Work on not assigned Luns error at silentlyContinue                                                               
    if ($Datastore.Name -eq $null) {                                                              
     $DatastoreName = "Not mapped"                                                              
     $FileSystemVersion = "Not mapped"                                                              
    }                                                              
    else {                                                              
     $DatastoreName = $Datastore.Name -join ","                                                              
     $FileSystemVersion = $Datastore[0].FileSystemVersion                                                               
    }                                                              
                                                                   
    $DatastoreFreeSpace = $Datastore.FreeSpaceGB -join ", "                                                               
    $DatastoreCapacityGB = $Datastore.CapacityGB -join ", "                                                               
    $DatastoreDatacenter = $Datastore.Datacenter -join ", "                                                               
                                                               
    $State = $LunPath.State -join ", "                                                              
    $Preferred = $LunPath.Preferred -join ", "                                                              
    $Paths = ($LunPath.ExtensionData.Transport | foreach {($_.Address -split ":")[0]}) -Join ", "                                                              
    $IsWorkingPath = $LunPath.ExtensionData.IsWorkingPath -Join ", "                                                              
                                                                  
    $Obj = New-Object PSObject                                                               
    $Obj | Add-Member -Name VMhost -MemberType NoteProperty -Value $hostdatastore.VMHost                                                               
    $Obj | Add-Member -Name DatastoreName -MemberType NoteProperty -Value $DatastoreName                                                                
    $Obj | Add-Member -Name FreeSpaceGB -MemberType NoteProperty -Value $DatastoreFreeSpace                                                               
    $Obj | Add-Member -Name CapacityGB -MemberType NoteProperty -Value $DatastoreCapacityGB                                                               
    $Obj | Add-Member -Name FileSystemVersion -MemberType NoteProperty -Value $FileSystemVersion                                                               
    $Obj | Add-Member -Name RuntimeName -MemberType NoteProperty -Value $hostdatastore.RuntimeName                                                               
    $Obj | Add-Member -Name CanonicalName -MemberType NoteProperty -Value $hostdatastore.CanonicalName                                                               
    $Obj | Add-Member -Name MultipathPolicy -MemberType NoteProperty -Value $hostdatastore.MultipathPolicy                                                               
    $Obj | Add-Member -Name Vendor -MemberType NoteProperty -Value $hostdatastore.Vendor                                                               
    $Obj | Add-Member -Name DatastoreDatacenter -MemberType NoteProperty -Value $DatastoreDatacenter                                                               
    $Obj | Add-Member -Name VMsOnDataStore -MemberType NoteProperty -Value $VMsOnDatastore                                                               
    $Obj | Add-Member -Name NumberOfPaths -MemberType NoteProperty -Value $LunPath.Count                                                              
    $Obj | Add-Member -Name Paths -MemberType NoteProperty -Value $Paths                                                              
    $Obj | Add-Member -Name State -MemberType NoteProperty -Value $State                                                              
    $Obj | Add-Member -Name Preferred -MemberType NoteProperty -Value $Preferred                                                              
    $Obj | Add-Member -Name IsWorkingPath -MemberType NoteProperty -Value $IsWorkingPath                                                                                                                            
    $report += $Obj 
   }                                                               
   $report | Export-Csv -NoTypeInformation /tmp/DatastoreInfoHostwise.csv
   Send-FilestoOpBot -filePath /tmp/DatastoreInfoHostwise.csv -channels  "#poshtalk"
 } 
