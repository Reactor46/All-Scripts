function Get-VMHostInfoSummary {
Connect-VIServer -Server lasvcenter01
$VMHostInfo = Get-VMHost | ForEach-Object {
$GuestCount = (@($_ | Get-VM)).count
$HostView = $_ | Get-View
$CPUInfo = $HostView.hardware.cpuinfo
$_ | Add-Member -MemberType NoteProperty -Name GuestCount -Value $GuestCount
$_ | Add-Member -MemberType NoteProperty -Name NumCPUSocket -Value $CPUInfo.NumCpuPackages
$_ | Add-Member -MemberType NoteProperty -Name NumCpuCores -Value $CPUInfo.NumCpuCores
$_ | Add-Member -MemberType NoteProperty -Name NumCpuThreads -Value $CPUInfo.NumCpuThreads
$_ | Add-Member -MemberType NoteProperty -Name TotalGB -Value ([Math]::Round(($HostView.summary.hardware.memorysize/1GB)))
$_
} | Select-Object Name, Parent, Version, Build, Model,NumCPUSocket, TotalGB, GuestCount, NumCpuCores, NumCpuThreads
#Send-HTMLEmail -InputObject $VMHostInfo -Subject "VMHost Info"

}#Get-VMHostInfoSummary
$VMHostInfo

