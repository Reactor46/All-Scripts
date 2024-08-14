REM: This bat file will reset all the monitors which are defined.
REM: Fill in the display name of the monitor to reset.
REM: If the monitor has a % in it make sure to double it.


Powershell -Command "& C:\scripts\Maintenance\ResetAllMonitorsOfASpecificType.ps1 'Logical Disk Fragmentation Level'"

