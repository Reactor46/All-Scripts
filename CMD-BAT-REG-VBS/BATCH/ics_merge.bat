REM @ECHO OFF

ECHO This is: ics_merge

SET ics_mergeTo=%1
SET ics_mergeFrom=%2

IF [%ics_mergeFrom%]==[] (EXIT 900)

copy %ics_mergeTo%+%ics_mergeFrom% %ics_mergeTo% /A /Y

EXIT 0
