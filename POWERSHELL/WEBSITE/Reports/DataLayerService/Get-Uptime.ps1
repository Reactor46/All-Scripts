#---------Get-Uptime.ps1------------#
param([string] $serverList)
foreach($server in (get-content $serverList)) {
  $lastBootWmi = get-wmiobject -query "SELECT LastBootUpTime FROM Win32_OperatingSystem" -computerName $server
  $lastBootDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($lastBootWmi.LastBootUpTime)
  $uptime = new-timespan -start $lastBootDate -end (get-date)
  New-Object PsObject -property @{Server = $server; Days = $uptime.Days; Hours = $uptime.Hours; Minutes = $uptime.Minutes; Seconds = $uptime.Seconds}
}