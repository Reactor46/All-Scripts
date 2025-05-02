REM @ECHO OFF

ECHO This is: ics_put

SET ics_ftp=%1
SET ics_user=%2
SET ics_pass=%3
SET ics_what=%4

IF [%ics_what%]==[] (EXIT 900)

SET whereTo=%5
SET remoteName=%6
SET myWorkDir=%7

SET remotePath=%whereTo%

SET checkFtpExportSuccess="^226.*Transfer.*complete"
SET checkFtpTransfFailed="^226.*Quota.*exceeded"
SET ftpCmdFile=store_archs.ftp
SET ics_putOut=ics_put.out

IF [%whereTo%]==[] SET remotePath=.
IF [%whereTo%]==[""] SET remotePath=.
IF [%remoteName%]==[] SET remoteName=%ics_what%
IF [%remoteName%]==[""] SET remoteName=%ics_what%
IF NOT [%myWorkDir%]==[] (CD %myWorkDir%)

%~d6

ECHO %ics_user%> %ftpCmdFile%
ECHO %ics_pass%>> %ftpCmdFile%
ECHO binary >> %ftpCmdFile%
COPY store_archs.ftp+ics_exportMakePath %ftpCmdFile% /A /Y
ECHO put %ics_what% %remotePath%/%remoteName% >> %ftpCmdFile%
ECHO quit >> %ftpCmdFile%

FTP -s:%ftpCmdFile% %ics_ftp% > %ics_putOut%


FINDSTR /R /I /C:%checkFtpTransfFailed% %ics_putOut%
IF %ERRORLEVEL% EQU 0 (EXIT 510)

FINDSTR /R /I /C:%checkFtpExportSuccess% %ics_putOut%
IF %ERRORLEVEL% NEQ 0 (EXIT 500)

EXIT 0
