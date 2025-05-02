REM @ECHO OFF

ECHO This is: ics_checkState

SET ics_stateOld=%1
SET ics_stateCurr=%2

IF [%ics_stateCurr%]==[] (EXIT 900)

SET myWorkDir=%3

IF [%myWorkDir%]==[] SET myWorkDir="%CD%"

%~d1

IF NOT [%ics_stateOld%]==[""] FC /L %ics_stateOld% %myWorkDir%\%ics_stateCurr%

EXIT %ERRORLEVEL%
