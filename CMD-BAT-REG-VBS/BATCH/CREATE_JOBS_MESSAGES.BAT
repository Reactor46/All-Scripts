@echo off
SETLOCAL

REM -------------------------------------------------------------------------
REM     This SAMPLE script is used to create Jobs & Data Messages for 
REM     SQL Server to support the App Virt schema.
REM -------------------------------------------------------------------------

IF "%1"=="" GOTO USAGE
IF "%2"=="" GOTO USAGE
IF "%1"=="/?" GOTO USAGE
IF "%1"=="-?" GOTO USAGE
IF "%1"=="?" GOTO USAGE
IF /I "%2"=="master" GOTO USAGE

REM Cleanup old log files.
IF EXIST create_jobs_messages.log DEL create_jobs_messages.log 1>nul 2>nul
IF EXIST create_jobs_messages.log GOTO NO_ACCESS

echo.
echo Begin create_tables...                                >> create_jobs_messages.log
IF NOT EXIST create_jobs_messages.log GOTO NO_ACCESS


echo.                                                      >> create_jobs_messages.log
echo ----------------------------------------------------- >> create_jobs_messages.log
echo    C  R  E  A  T  E   J O B S   M E S S A G E S ++    >> create_jobs_messages.log
echo ----------------------------------------------------- >> create_jobs_messages.log
DATE /T                                                    >> create_jobs_messages.log
TIME /T                                                    >> create_jobs_messages.log
echo ----------------------------------------------------- >> create_jobs_messages.log
echo.                                                      >> create_jobs_messages.log



REM -------------------------------------------------------------------------
echo running data_messages.sql
echo running data_messages.sql                             >> create_jobs_messages.log
osql -n -E -S%1 -d%2 -idata_messages.sql                   >> create_jobs_messages.log

REM -------------------------------------------------------------------------
echo running alerts_jobs.sql
echo running alerts_jobs.sql                               >> create_jobs_messages.log
osql -n -E -S%1 -d%2 -ialerts_jobs.sql                     >> create_jobs_messages.log
FIND "Msg " create_jobs_messages.log|FIND ", Level"|FIND ", State"
IF %ERRORLEVEL% EQU 0 GOTO BAD_CREATE

REM -------------------------------------------------------------------------
echo running 2645392_UPDATE.SQL
echo running 2645392_UPDATE.SQL                            >> create_jobs_messages.log
osql -n -E -S%1 -d%2 -i2645392_UPDATE.SQL                  >> create_jobs_messages.log

echo.
echo SCRIPT SUCCEEDED      
echo.                                                      >> create_jobs_messages.log
echo SCRIPT SUCCEEDED                                      >> create_jobs_messages.log
GOTO DONE

:BAD_CREATE
echo.
echo SCRIPT FAILED - Check create_jobs_messages.log for results.
echo.                                                      >> create_jobs_messages.log
echo SCRIPT FAILED - Check create_jobs_messages.log for results.  >> create_jobs_messages.log
GOTO DONE

:NO_ACCESS
echo.
echo SCRIPT FAILED - Failed to create log file
echo Make sure access rights are correct for this directory.
GOTO DONE


:USAGE
echo.
echo USAGE:
echo.
echo     create_jobs_messages.bat SERVERNAME[\INSTANCE] APPVIRTDATABASE
echo.
echo     Note: Cannot specify the 'master' database as the APPVIRTDATABASE name.
echo.
GOTO DONE

:DONE
echo.                                                      >> create_jobs_messages.log
echo ----------------------------------------------------- >> create_jobs_messages.log
DATE /T                                                    >> create_jobs_messages.log
TIME /T                                                    >> create_jobs_messages.log
echo ----------------------------------------------------- >> create_jobs_messages.log
echo.                                                      >> create_jobs_messages.log
echo.
IF "%1"=="" pause
ENDLOCAL
