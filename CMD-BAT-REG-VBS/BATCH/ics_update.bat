REM @ECHO OFF

ECHO This is: ics_synchronize

SET ics_dirSrcRoot=%1
SET ics_toPack=%2
SET ics_packAppPath=%3

IF [%ics_packAppPath%]==[] (EXIT 900)

SET ics_dirSrcWork=%4
SET ics_pass=%5
SET ics_myWorkDir=%6
SET ics_archType=%7
SET ics_noRecursionFlag=%8

SET toExport=%ics_dirSrcWork%
SET recursionOptn=-r
SET passOptn=-p%ics_pass%
SET exitCode=0

IF [%ics_dirSrcWork%]==[""] (SET toExport=)
IF [%ics_pass%]==[] (SET passOptn=)
IF [%ics_pass%]==[""] (SET passOptn=)
IF [%ics_myWorkDir%]==[] SET ics_myWorkDir="%CD%"
IF [%ics_myWorkDir%]==[""] SET ics_myWorkDir="%CD%"
IF [%ics_archType%]==[""] (SET ics_archType=)
IF [%ics_noRecursionFlag%]==[norecursion] (SET recursionOptn=)

SET excludeOptn=-x@%ics_myWorkDir%\ics_packExclude
IF NOT EXIST %ics_myWorkDir%\ics_packExclude (SET excludeOptn=)

SET packCmd=%ics_packAppPath%\7zip\7z.exe u -y -ssw %recursionOptn% %ics_archType% %passOptn%

%~d1

CD %ics_dirSrcRoot%

%packCmd% %excludeOptn% %ics_myWorkDir%\%ics_toPack% %toExport% > %ics_myWorkDir%\ics_synchronize.out

IF %ERRORLEVEL% EQU 1 (
SET exitCode=0
) ELSE (SET exitCode=%ERRORLEVEL%)

TYPE %ics_myWorkDir%\ics_synchronize.out | FINDSTR /C:"cannot access" > %ics_myWorkDir%\ics_synchronize.warr

IF %exitCode% NEQ 0 (EXIT 400)

EXIT 0
