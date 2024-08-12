@ECHO OFF
:: =================== CONFIG ============================================
:: FROM Path (can be a share or fully qualified dir)
SET fromdir=\\172.20.0.231\NextGenRoot\Prod\BBP_Files
REM SET fromdir1=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\OMP
REM SET fromdir2=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\PAC_AAC
REM SET fromdir3=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\USON
REM SET fromdir4=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\OMP
REM SET fromdir5=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\PAC
REM SET fromdir6=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\ONC
REM SET fromdir7=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\ONC
REM SET fromdir8=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\ORTHO
REM SET fromdir9=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\ORTHO
REM SET fromdir10=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\EDI\OCMG PULM
REM SET fromdir11=\\172.20.0.231\NextGenRoot\Prod\BBP_Files\OCMG PULM

SET todir=E:\BBP_Archive


:: Path for log file !!!USE TRAILING SLASH!!! (default is a dir called "logs" in this script's dir)
SET logdir=E:\Scripts\LOGS\
:: =================== CONFIG ============================================

:: =================== SCRIPT ============================================
:: !!! NO NEED TO EDIT BEYOND THIS POINT !!!
:: Make a date_time stamp like 20130125_134200
SET hh=%time:~0,2%
:: Add a zero when this is run before 10 am.
IF "%time:~0,1%"==" " set hh=0%hh:~1,1%
SET yyyymmdd_hhmmss=%DATE:~-4%%DATE:~4,2%%DATE:~7,2%_%hh%%time:~3,2%%time:~6,2%
:: Make a name for the log file
SET logfile=%logdir%%yyyymmdd_hhmmss%_%~n0.log
:: (If not exist) create a logdir directory
IF NOT EXIST %logdir% MKDIR %logdir%


:: The Robocopy magic
	robocopy "%fromdir%" "%todir%" /COPY:DATSO /B /MIR /LOG+:%logfile% /ZB /R:10 /W:30 /TEE
:: =================== SCRIPT ============================================