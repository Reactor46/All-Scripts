@echo on

REM Delete IPC$ - Deletes open connections to the folder

net use \\192.168.220.76\iemdata /d /y

sleep 10

REM Connect to the folder

net use \\192.168.220.76\iemdata\DFTNextGen\process OPTlog6876! /USER:task.admin@roc

REM Copy the file(s)

Xcopy \\192.168.220.76\iemdata\DFTNextGen\process\*.* \\optfile01.cloud.local\NextGenRoot\Varian /y /z /c

Xcopy \\192.168.220.76\iemdata\DFTNextGen\process\*.* \\192.168.220.76\iemdata\DFTNextGen\Archive /y /z /c

REM Move file(s) to Varian archive

Move /Y \\192.168.220.76\iemdata\DFTNextGen\process\*.* \\optfile01.cloud.local\NextGenRoot\Varian\Archive

exit