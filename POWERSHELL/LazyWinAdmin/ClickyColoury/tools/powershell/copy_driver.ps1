# api: multitool
# version: 0.3.7
# title: Copy driver
# description: from X:\Treiber\ to user \Temp
# type: window
# depends: wpf
# category: powershell
# img: copy
# hidden: 0
# status: untested
# param: machine, driver
#
# ServiceDesk Driver Collection
#  ❏ from X:\drivers
#  ❏ coypied to user \\$machine\C$\Temp\
#


#-- vars
Param(
    $machine = (Read-Host "Computer"),
    $driver = (Read-Host "Driver"),
    $cache_fn = "data/combobox.driver.txt",
    $driver_d = "X:\Drivers",
    $CRLF = "`r`n"
)

#-- update list
if ($driver -match "^-*update-*(list)?$") {
    Write-Host -f Green "❏ updating $cache_fn"
    $r = "update-list"
    ForEach ($fn in GCI $driver_d) {
        $r += "$CRLF$fn"
    }
    $r | Out-File $cache_fn -Encoding UTF8
}

#-- else copy
elseif (Test-Connection -Quiet $machine) {
    md "\\$machine\c$\Temp\$driver"
    robocopy /E /V /B  "$driver_d\$driver" "\\$machine\c$\Temp\$driver"

    Write-Host -f Green "-- Close me window --"
    Start-Sleep -seconds 20
}
