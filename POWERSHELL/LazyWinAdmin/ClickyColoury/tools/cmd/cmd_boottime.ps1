# api: multitool
# version: 2.0
# title: Boot time
# description: Determine last PC startup
# type: psexec
# category: cmd
# img: date.png
# hidden: 0
# key: 3|boot
# src: https://github.com/lazywinadmin/LazyWinAdmin_GUI/blob/master/LazyWinAdmin/LazyWinAdmin.ps1#L159
#
# Now queries via WMI Win32_OperatingSystem ➔ Lastbootuptime
#
# CMD approach:
#  ❏ systeminfo /s $machine
#

Param($machine = (Read-Host "Computer"))


#-- new WMI
$wmi = Get-WmiObject -class Win32_OperatingSystem -computer $machine
$uptime = $wmi.ConvertToDateTime($wmi.Lastbootuptime)
Write-Host "Last boot time: $uptime"

