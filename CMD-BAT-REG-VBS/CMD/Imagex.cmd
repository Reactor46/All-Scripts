@ECHO OFF
Title  Restore Factory recovery image from HDD

:INIT
set IMAGEX=x:\sources\recovery\tools\imagex.exe
set DISKPART_SCRIPT=x:\sources\recovery\tools\DiskPart_script.txt
set SETAUTOFAILOVER=x:\sources\recovery\tools\setAutoFailoverRE.cmd
set DISKPART=x:\windows\system32\diskpart.exe
set Index=1

:MAIN
REM Assign R letter to Recovery partition and format regular partition

ECHO Sel Dis 0 > %DISKPART_SCRIPT%
ECHO Sel Par 1 >> %DISKPART_SCRIPT%
ECHO Remove >> %DISKPART_SCRIPT%
ECHO Assign Letter=R >> %DISKPART_SCRIPT%
ECHO Sel Par 2 >> %DISKPART_SCRIPT%
ECHO Del Par >> %DISKPART_SCRIPT%
ECHO Create Par Pri >> %DISKPART_SCRIPT%
ECHO Sel Par 2 >> %DISKPART_SCRIPT%
ECHO Format fs=NTFS Label="Windows" quick >> %DISKPART_SCRIPT%
ECHO Assign letter=W >> %DISKPART_SCRIPT%
%diskpart% /s %DISKPART_SCRIPT% > null
del %DISKPART_SCRIPT%
set RECOVERY_DRIVE=R

cls
REM Ensure recovery image is present before continuing

if not exist %RECOVERY_DRIVE%:\Recovery\WindowsRE\install.wim goto :EOF
Title Apply Recovery Image to OS partition
%IMAGEX% /apply %RECOVERY_DRIVE%:\Recovery\WindowsRE\install.wim %Index% W:

REM Enable WinRE failover and hide recovery partition
cmd /c %SETAUTOFAILOVER%
