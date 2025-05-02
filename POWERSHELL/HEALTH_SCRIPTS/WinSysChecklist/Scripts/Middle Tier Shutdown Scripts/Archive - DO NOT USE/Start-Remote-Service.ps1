#SETTINGS
$service = “Your Service”
$computer = “Your Computer”

$state = (gwmi -query “select * from win32_service where DisplayName=’$service’” -computer $computer).State
$trigger = 0

while ($state -eq “Stopped”) {
write-host “`r$service is $state” -nonewline -foregroundcolor red
sleep -milliseconds 100
if ($trigger -lt 1) {
(gwmi -query “select * from win32_service where DisplayName=’$service’” -computer $computer).startservice() | out-null
$trigger = 1
}
$state = (gwmi -query “select * from win32_service where DisplayName=’$service’” -computer $computer).State
}

while ($state -eq “Start Pending”) {
write-host “`r$service is $state” -nonewline -foregroundcolor yellow
sleep -milliseconds 100
$state = (gwmi -query “select * from win32_service where DisplayName=’$service’” -computer $computer).State
}

#SETTINGS
$service = “Servers Alive”
$computer = “Vesuvius”

$state = (gwmi -query “select * from win32_service where DisplayName=’$service’” -computer $computer).State
$trigger = 0

while ($state -eq “Stopped”) {
write-host “`r$service is $state” -nonewline -foregroundcolor red
sleep -milliseconds 100
if ($trigger -lt 1) {
(gwmi -query “select * from win32_service where DisplayName=’$service’” -computer $computer).startservice() | out-null
$trigger = 1
}
$state = (gwmi -query “select * from win32_service where DisplayName=’$service’” -computer $computer).State
}

while ($state -eq “Start Pending”) {
write-host “`r$service is $state” -nonewline -foregroundcolor yellow
sleep -milliseconds 100
$state = (gwmi -query “select * from win32_service where DisplayName=’$service’” -computer $computer).State
}

if ($state -eq “Running”) {
write-host “`r ” -nonewline
write-host “`r$service is $state`n” -nonewline -foregroundcolor green
}

if ($state -eq “Stopped”) {
write-host “`r ” -nonewline
write-host “`r$service is $state (Failed to Start)`n” -nonewline -foregroundcolor red
}
