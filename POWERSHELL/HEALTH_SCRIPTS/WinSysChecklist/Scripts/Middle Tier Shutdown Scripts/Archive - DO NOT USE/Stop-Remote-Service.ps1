#SETTINGS
$service = “Windows Time”
$computer = “MIS-UTIL”

$state = (gwmi -Query “select * from win32_service where DisplayName=’$service’” -computer $computer).State
$trigger = 0

while ($state -eq “Running”) {
write-host “`r$service is $state” -nonewline -foregroundcolor green
sleep -milliseconds 100
if ($trigger -lt 1) {
(gwmi -query “select * from win32_service where DisplayName=’$service’” -computer $computer).stopservice() | out-null
$trigger = 1
}
$state = (gwmi -query “select * from win32_service where DisplayName=’$service’” -computer $computer).State
}

while ($state -eq “Stop Pending”) {
write-host “`r$service is $state” -nonewline -foregroundcolor yellow
sleep -milliseconds 100
$state = (gwmi -query “select * from win32_service where DisplayName=’$service’” -computer $computer).State
}

if ($state -eq “Stopped”) {
write-host “`r ” -nonewline
write-host “`r$service is $state`n” -nonewline -foregroundcolor red
}

if ($state -eq “Running”) {
write-host “`r ” -nonewline
write-host “`r$service is $state (Failed to Stop)`n” -nonewline -foregroundcolor Green
}
