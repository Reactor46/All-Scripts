@ECHO ON

SET Version=Unknown

wmic os get osarchitecture | FINDSTR /IL "32" > NUL
IF %ERRORLEVEL% EQU 0 SET Version="32"

wmic os get osarchitecture | FINDSTR /IL "64" > NUL
IF %ERRORLEVEL% EQU 0 SET Version="64"

ECHO The OS architecture of Windows found is %VERSION% bit

IF %VERSION% == "32" GOTO 32
IF %VERSION% == "64" GOTO 64
::If no versions are found go to UNKNOWN
GOTO UNKNOWN

:32
ECHO Execute script for 32 bit OS
copy "\\lasfs03\software\Current Versions\ESET\ESET 6.2.2\ESET-EPP-6.5.2107.0_x86_en_US.exe" C:\Windows\Temp
cd c:\windows\temp\
start /wait ESET-EPP-6.5.2107.0_x86_en_US.exe --silent --accepteula
GOTO FINISH

:64
ECHO Execute script for 64 bit OS
copy "\\lasfs03\software\Current Versions\ESET\ESET 6.x\ESET-EPP-6.5.2107.0_x64_en_US.exe" C:\Windows\Temp
cd c:\windows\temp\
start /wait ESET-EPP-6.5.2107.0_x64_en_US.exe --silent --accepteula
GOTO FINISH

:FINISH
ECHO Script executed successfully
GOTO END

:UNKNOWN
ECHO OS Architecture Unknown

:END