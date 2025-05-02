REM @ECHO OFF

ECHO This is: ics_confirm

SET ics_ftp=%1
SET ics_user=%2
SET ics_pass=%3

IF [%ics_pass%]==[] (EXIT 900)

SET whereTo=%4
SET myWorkDir=%5

SET remotePath=%whereTo%

SET checkFtpExportSuccess="^226.*Transfer.*complete"
SET checkFtpTransfFailed="^226.*Quota.*exceeded"
SET ftpCmdFile=store_archs.ftp
SET ics_confirmOut=ics_confirm.out
SET ics_confirm=ics_confirm.list

IF [%whereTo%]==[] SET remotePath=.
IF [%whereTo%]==[""] SET remotePath=.
IF NOT [%myWorkDir%]==[] (CD %myWorkDir%)

%~d6

ECHO %ics_user%> %ftpCmdFile%
ECHO %ics_pass%>> %ftpCmdFile%
ECHO CD %remotePath% >> %ftpCmdFile%
ECHO dir -a %ics_confirm% >> %ftpCmdFile%
ECHO quit >> %ftpCmdFile%

FTP -s:%ftpCmdFile% %ics_ftp% > %ics_confirmOut%


FINDSTR /R /I /C:%checkFtpTransfFailed% %ics_confirmOut%
IF %ERRORLEVEL% EQU 0 (EXIT 510)

FINDSTR /R /I /C:%checkFtpExportSuccess% %ics_confirmOut%
IF %ERRORLEVEL% NEQ 0 (EXIT 500)

EXIT 0
