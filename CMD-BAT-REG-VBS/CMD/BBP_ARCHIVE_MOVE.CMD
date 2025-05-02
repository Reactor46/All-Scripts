@ECHO OFF
:: =================== CONFIG ============================================
:: FROM Path (can be a share or fully qualified dir)
SET fromdir=\\usonpsvrfpf\E$\BBP_Archive
SET todir=E:\BBP_Archive\Unsorted


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