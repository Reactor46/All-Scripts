REM @ECHO OFF

ECHO This is: ics_rename

SET ics_ftp=%1
SET ics_user=%2
SET ics_pass=%3
SET ics_what=%4
SET ics_ts=%5

IF [%ics_ts%]==[] (EXIT 900)

SET whereTo=%6
SET myWorkDir=%7

SET remoteLoc=%whereTo%/%ics_what%
SET ftpCmdFile=store_archs.ftp

IF [%whereTo%]==[] SET remoteLoc=%ics_what%
IF [%whereTo%]==[""] SET remoteLoc=%ics_what%
IF NOT [%myWorkDir%]==[] (CD %myWorkDir%)

%~d6

ECHO %ics_user%> %ftpCmdFile%
ECHO %ics_pass%>> %ftpCmdFile%
ECHO rename %remoteLoc% %remoteLoc%_%ics_ts% >> %ftpCmdFile%
ECHO quit >> %ftpCmdFile%

FTP -s:%ftpCmdFile% %ics_ftp%

EXIT 0
