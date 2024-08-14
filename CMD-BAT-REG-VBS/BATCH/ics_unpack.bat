REM @ECHO OFF

ECHO This is: ics_unpack

SET ics_toUnpack=%1
SET ics_dirDestRoot=%2
SET ics_unpackAppPath=%3

IF [%ics_dirDestRoot%]==[] (EXIT 900)

SET unpackForceFresh=%4
SET dirDestWork=%5
SET pass=%6
SET myWorkDir=%7

SET dirDestTmp=ics_unpack_tmp

IF [%dirDestWork%]==[""] SET dirDestWork=
IF [%myWorkDir%] NEQ [] (CD %myWorkDir%)

SET unpackCmd=%ics_unpackAppPath%\7zip\7z.exe x -y -o%dirDestTmp% -p%pass%

IF [%pass%]==[] (SET unpackCmd=%ics_unpackAppPath%\7zip\7z.exe x -y -o%dirDestTmp% )
IF [%pass%]==[""] (SET unpackCmd=%ics_unpackAppPath%\7zip\7z.exe x -y -o%dirDestTmp% )

%~d6

IF EXIST %myWorkDir%\ics_unpackIgnore (
   %unpackCmd% -xr@%myWorkDir%\ics_unpackIgnore %ics_toUnpack%
) ELSE (
   %unpackCmd% %ics_toUnpack%
)

IF %ERRORLEVEL% EQU 2 (EXIT 320)
IF %ERRORLEVEL% EQU 7 (EXIT 300)
IF %ERRORLEVEL% EQU 8 (EXIT 320)

REM IF %ERRORLEVEL% EQU 2 GOTO lblUnpack
REM IF %ERRORLEVEL% NEQ 0 (EXIT 300)

:lblUnpack
SET ERRORLEVEL=0
IF [%unpackForceFresh%] EQU [true] (RMDIR /S /Q %ics_dirDestRoot%\%dirDestWork%)

IF NOT EXIST %ics_dirDestRoot% (MD %ics_dirDestRoot%)
IF %ERRORLEVEL% NEQ 0 (EXIT 330)
IF NOT EXIST %ics_dirDestRoot% (EXIT 330)

xcopy /H /S /E  /Y /C /Q  /R /K %dirDestTmp%\*.* %ics_dirDestRoot%
IF %ERRORLEVEL% NEQ 0 (EXIT 310)

REM DEL /Q /F %ics_toUnpack%

EXIT 0
