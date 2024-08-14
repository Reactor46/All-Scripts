REM @ECHO OFF

ECHO This is: ics_get

SET ics_ftp=%1
SET ics_user=%2
SET ics_pass=%3
SET ics_what=%4

IF [%ics_what%]==[] (EXIT 900)

SET whereFrom=%5
SET localName=%6
SET myWorkDir=%7

SET remoteLoc=%whereFrom%/%ics_what%

SET checkFtpExistents="Password required"
REM SET checkFtpLogingSuccess="^230.*login.*accepted"
SET checkFtpExportSuccess="^226.*Transfer.*complete"
REM SET checkFtpConnEnd="^221.*goodbye"

SET ftpCmdFile=store_archs.ftp
SET ics_getOut=ics_get.out

IF [%whereFrom%]==[] SET remoteLoc=%ics_what%
IF [%whereFrom%]==[""] SET remoteLoc=%ics_what%
IF [%localName%]==[] SET localName=
IF [%localName%]==[""] SET localName=
IF NOT [%myWorkDir%]==[] (CD %myWorkDir%)

%~d6

ECHO %ics_user%> %ftpCmdFile%
ECHO %ics_pass%>> %ftpCmdFile%
ECHO binary >> %ftpCmdFile%
ECHO get %remoteLoc% %localName% >> %ftpCmdFile%
ECHO quit >> %ftpCmdFile%

FTP -s:%ftpCmdFile% %ics_ftp% > %ics_getOut%

FINDSTR /R /I /C:%checkFtpExistents% %ics_getOut%
IF %ERRORLEVEL% NEQ 0 (EXIT 200)

REM FINDSTR /R /I /C:%checkFtpLogingSuccess% %ics_getOut%
REM IF %ERRORLEVEL% NEQ 0 (EXIT 210)

FINDSTR /R /I /C:%checkFtpExportSuccess% %ics_getOut%
IF %ERRORLEVEL% NEQ 0 (
DEL /F /Q %localName%
EXIT 220
)

EXIT 0
