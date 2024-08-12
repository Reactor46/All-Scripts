@set COPY_jdb_TO="C:\Program Files\Symantec\Symantec Endpoint Protection Manager\data\inbox\content\incoming"
@set RAPIDRELEASE=0
@set jdbTEMP=%temp%

@rem ==============================================================================================
@rem Set RAPIDRELEASE=1 to download repidrelease definitions, RAPIDRELEASE=0 for fully QA'd definitions (standard).
@rem Change COPY_jdb_TO= to point to the SAV CE server directory (or where you want the jdb file copied)
@rem   you can also run the script directly from the SAV folder and it will copy the definitions there.
@rem jdbTEMP is the temp folder the script will use while downloading definitions, set to %temp% to use system default
@rem ==============================================================================================
@rem  Script for downloading virus definition updates for
@rem  Symantec Antivirus Corporate Edition version 8.x and 9.x
@rem  This unsupported utility is provided for your convenience only.
@rem  Symantec Technical Support cannot provide support for the creation,
@rem  use, or troubleshooting of Windows scripts.

@rem ==============================================================================================
@echo off


rem   ========= check that OS is win2k or better ============
if not "%OS%" == "Windows_NT" goto BADOS
if "%APPDATA%" == "" goto BADOS

rem   ========= make sure to be in script directory ============
if exist rtvscan.exe set COPY_jdb_TO=%CD%
for %%i in (%0) do @%%~di
for %%i in (%0) do @cd %%~pi
if exist rtvscan.exe set COPY_jdb_TO=%CD%

rem   =========== get name/size of last file from "jdbdown.lastfile" ============
if not exist jdbdown.lastfile goto NOLAST
for /f "tokens=1" %%f in (jdbdown.lastfile) do set lastfile=%%f
for /f "tokens=2" %%f in (jdbdown.lastfile) do set lastsize=%%f
:NOLAST

rem   ========= jump to temp dir ============
if not exist "%jdbTEMP%\jdbtmp" md "%jdbTEMP%\jdbtmp"
if exist "%jdbTEMP%\jdbtmp\*.jdb" del "%jdbTEMP%\jdbtmp\*.jdb"
pushd "%jdbTEMP%\jdbtmp"

rem   =========== make ftp script for checking jdb directory on ftp ===========
echo open ftp.symantec.com> check.txt
echo anonymous>> check.txt
echo email@address.com>> check.txt
set jdbfolder=jdb
if "%RAPIDRELEASE%" == "1" set jdbfolder=rapidrelease
echo cd AVDEFS/symantec_antivirus_corp/%jdbfolder%>> check.txt
echo dir *.jdb chk.lst>> check.txt
echo bye>> check.txt

rem   =========== get filename and size from ftp ============
if exist chk.lst del chk.lst
ftp -s:check.txt
if not exist chk.lst goto ERROR
for /f "tokens=9" %%f in (chk.lst) do set jdbfile=%%f
for /f "tokens=5" %%f in (chk.lst) do set jdbsize=%%f
if "%jdbfile%" == "" goto ERROR
if "%jdbsize%" == "" goto ERROR

rem   =========== compare ftp name/size to local ============
if not "%jdbfile%" == "%lastfile%" goto DOWNLOAD
if not "%jdbsize%" == "%lastsize%" goto DOWNLOAD
popd
echo.
echo Already downloaded latest %jdbfolder% file: %jdbfile% - size %jdbsize%
echo %date% %time%  Already downloaded latest %jdbfolder% file: %jdbfile% - size %jdbsize% >> jdbdown.log
goto END

:DOWNLOAD
rem   ========= make ftp script for downloading new jdb file =========
echo open ftp.symantec.com> down.txt
echo anonymous>> down.txt
echo email@address.com>> down.txt
echo cd AVDEFS/symantec_antivirus_corp/%jdbfolder%>> down.txt
echo bin>> down.txt
echo hash>> down.txt
echo get %jdbfile%>> down.txt
echo bye>> down.txt

rem   ============= download new file =================
ftp -s:down.txt
for %%i in (%jdbfile%) do @set newsize=%%~zi
if not "%newsize%" == "%jdbsize%" goto ERROR
copy %jdbfile% E:\Downloads
move %jdbfile% %COPY_jdb_TO%
if exist %jdbfile% goto ERRORMOVE
popd
echo.
echo %jdbfile% %jdbsize% > jdbdown.lastfile
echo Downloaded new %jdbfolder% file: %jdbfile% - size %jdbsize%
echo %date% %time%  Downloaded new %jdbfolder% file: %jdbfile% - size %jdbsize% >> jdbdown.log
goto END


:ERROR
popd
echo.
echo ERROR: problem downloading %jdbfolder% definition file. jdbfile=%jdbfile% jdbsize=%jdbsize% newsize=%newsize% (lastfile=%lastfile% lastsize=%lastsize%).
echo %date% %time%  ERROR: problem downloading %jdbfolder% definition file. jdbfile=%jdbfile% jdbsize=%jdbsize% newsize=%newsize% (lastfile=%lastfile% lastsize=%lastsize%). >> jdbdown.log
type "%jdbTEMP%\jdbtmp\chk.lst" >> jdbdown.log
echo.  >> jdbdown.log
goto END

:ERRORMOVE
popd
echo.
echo ERROR: problem moving definition file to SAV folder. COPY_jdb_TO=%COPY_jdb_TO%  newsize=%newsize% (lastfile=%lastfile% lastsize=%lastsize%).
echo %date% %time%  ERROR: problem moving definition file to SAV folder. COPY_jdb_TO=%COPY_jdb_TO%  newsize=%newsize% (lastfile=%lastfile% lastsize=%lastsize%). >> jdbdown.log
goto END

:BADOS
echo.
echo ERROR: this script needs Windows 2000 or better.
echo %date% %time%  ERROR: this script needs Windows 2000 or better. >> jdbdown.log
goto END

:END
if exist "%jdbTEMP%\jdbtmp\check.txt" del "%jdbTEMP%\jdbtmp\check.txt"
if exist "%jdbTEMP%\jdbtmp\down.txt" del "%jdbTEMP%\jdbtmp\down.txt"
if exist "%jdbTEMP%\jdbtmp\chk.lst" del "%jdbTEMP%\jdbtmp\chk.lst"
rd "%jdbTEMP%\jdbtmp"
set COPY_jdb_TO=
set RAPIDRELEASE=
set lastsize=
set lastfile=
set newsize=
set jdbsize=
set jdbfile=
set jdbfolder=
set jdbtemp=