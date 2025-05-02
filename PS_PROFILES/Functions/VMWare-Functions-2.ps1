## Begin Enable-MemHotAdd
Function Enable-MemHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
## End Enable-MemHotAdd
## Begin Disable-MemHotAdd
Function Disable-MemHotAdd($vm){
$vmview = Get-VM $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="mem.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
## End Disable-MemHotAdd
## Begin Enable-vCpuHotAdd
Function Enable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="true"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
## End Enable-vCpuHotAdd
## Begin Disable-vCpuHotAdd
Function Disable-vCpuHotAdd($vm){
$vmview = Get-vm $vm | Get-View
$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
$extra = New-Object VMware.Vim.optionvalue
$extra.Key="vcpu.hotadd"
$extra.Value="false"
$vmConfigSpec.extraconfig += $extra
$vmview.ReconfigVM($vmConfigSpec)
}
## End Disable-vCpuHotAdd
## Begin Get-VMOSList
Function Get-VMOSList {
    [cmdletbinding()]
    param($vCenter)
    
    Connect-VIServer $vCenter  | Out-Null
    
    [array]$osNameObject       = $null
    $vmHosts                   = Get-VMHost
    $i = 0
    
    foreach ($h in $vmHosts) {
        
        Write-Progress -Activity "Going through each host in $vCenter..." -Status "Current Host: $h" -PercentComplete ($i/$vmHosts.Count*100)
        $osName = ($h | Get-VM | Get-View).Summary.Config.GuestFullName
        [array]$guestOSList += $osName
        Write-Verbose "Found OS: $osName"
        
        $i++    
 
    
    }
    
    $names = $guestOSList | Select-Object -Unique
    
    $i = 0
    
    foreach ($n in $names) { 
    
        Write-Progress -Activity "Going through VM OS Types in $vCenter..." -Status "Current Name: $n" -PercentComplete ($i/$names.Count*100)
        $vmTotal = ($guestOSList | ?{$_ -eq $n}).Count
        
        $osNameProperty  = @{'Name'=$n} 
        $osNameProperty += @{'Total VMs'=$vmTotal}
        $osNameProperty += @{'vCenter'=$vcenter}
        
        $osnO             = New-Object PSObject -Property $osNameProperty
        $osNameObject     += $osnO
        
        $i++
    
    }    
    Disconnect-VIserver -force -confirm:$false
        
    Return $osNameObject
}
## End Get-VMOSList
## Begin Disconnect-ViSession
Function Disconnect-ViSession {
<#
.SYNOPSIS
Disconnects a connected vCenter Session.

.DESCRIPTION
Disconnects a open connected vCenter Session.

.PARAMETER  SessionList
A session or a list of sessions to disconnect.

.EXAMPLE
PS C:\> Get-VISession | Where { $_.IdleMinutes -gt 5 } | Disconnect-ViSession

.EXAMPLE
PS C:\> Get-VISession | Where { $_.Username -eq “User19” } | Disconnect-ViSession
#>
[CmdletBinding()]
Param (
[Parameter(ValueFromPipeline=$true)]
$SessionList
)
Process {
$SessionMgr = Get-View $DefaultViserver.ExtensionData.Client.ServiceContent.SessionManager
$SessionList | Foreach {
Write “Disconnecting Session for $($_.Username) which has been active since $($_.LoginTime)”
$SessionMgr.TerminateSession($_.Key)
}

}

}
## End Disconnect-ViSession
## Begin Get-ViSession
Function Get-ViSession {
<#
.SYNOPSIS
Lists vCenter Sessions.

.DESCRIPTION
Lists all connected vCenter Sessions.

.EXAMPLE
PS C:\> Get-VISession

.EXAMPLE
PS C:\> Get-VISession | Where { $_.IdleMinutes -gt 5 }
#>
$SessionMgr = Get-View $DefaultViserver.ExtensionData.Client.ServiceContent.SessionManager
$AllSessions = @()
$SessionMgr.SessionList | Foreach {
$Session = New-Object -TypeName PSObject -Property @{
Key = $_.Key
UserName = $_.UserName
FullName = $_.FullName
LoginTime = ($_.LoginTime).ToLocalTime()
LastActiveTime = ($_.LastActiveTime).ToLocalTime()

}
## End Get-ViSession
If ($_.Key -eq $SessionMgr.CurrentSession.Key) {
$Session | Add-Member -MemberType NoteProperty -Name Status -Value “Current Session”
} Else {
## End Get-ViSession
$Session | Add-Member -MemberType NoteProperty -Name Status -Value “Idle”
}
## End Get-ViSession
$Session | Add-Member -MemberType NoteProperty -Name IdleMinutes -Value ([Math]::Round(((Get-Date) – ($_.LastActiveTime).ToLocalTime()).TotalMinutes))
$AllSessions += $Session
}
## End Get-ViSession
$AllSessions
}
## End Get-ViSession
## Begin GoGo-VSphere
Function GoGo-VSphere {

Connect-VIServer -Server 10.20.1.9
}
## End GoGo-VSphere