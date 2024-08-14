@echo off
rem ==========================================================================================================================
rem - Script: instances.bat
rem - 
rem - Author: Jeff Mason aka TNJMAN aka bitdoctor
rem - Date: October 9, 2013
rem - Purpose: Many times a database administrator will want to know what SQL Server instances are running on remote servers.
rem - Below is a simple Windows Shell Script-based method for determining that information, either against a single server 
rem - or multiple servers.
rem - Notes:
rem -  How to Run this script:
rem -  1) Save this script as "c:\scripts\instances.bat"
rem -     a) To run against a single computer
rem -        Execute this script, passing the remote computer name as the only parameter
rem -         c:\scripts\instances.bat remote-sql-server
rem -           where remote-sql-server is the remote computer for which you need to list SQL instances
rem -     b) To run against multiple computers:
rem -        i. Make a wrapper "bat" file to 'call' this (instances.bat) script - 
rem -           There is a "wrapper file" example ("call-instances.bat") in the "rem" statements at the end of this script
rem -       ii. In the "wrapper" file, place a series of "call" statements, each on a single line, for each remote computer
rem -           i.e. "call c:\scripts\instances Server1" (that would run this script against the remote computer "Server1")
rem -      iii. Once you've entered all the lines containing all the target remote computer names, 
rem -           save that "wrapper" script as "c:\scripts\call-instances.bat"
rem -       iv. Execute the "wrapper" file, which will automatically generate a "results log/report"
rem -           c:\scripts\call-instances
rem -     c) Results log/report will automatically open in Notepad for review ("notepad c:\scripts\instances-log.txt")
rem -     
rem - Assumptions: 
rem - 1) You have (or create) a c:\scripts folder
rem - 2) You have adequate permissions/privileges to run necessary commands against a remote server
rem - NOTE: Also reports on inactive (not started) SQL Server instances by using "state= all" in the 'sc query' command
rem         If you want ONLY 'active/started' SQL instances, remove "state= all" in the 'sc query' command line
rem ==========================================================================================================================
rem
set mode=single
set server=%1
if "%1"=="batch" (
  set mode=batch
  set server=%2
  goto skipclean
)  
del c:\scripts\instances-log.txt > nul 2>&1
:skipclean
echo.
echo "RUNNING IN %mode% mode..."
set loc="c:\scripts"
set pfx="instances"
rem - c:\scripts\Instances.bat
echo.
echo "Listing instances of SQL Server on %server%"
echo "Listing instances of SQL Server on %server%" >> %loc%\%pfx%-log.txt
sc \\%server% query state= all | find /I "SQL Server (" >> %loc%\%pfx%-log.txt
echo. >> %loc%\%pfx%-log.txt
goto %mode%
:single
  echo.
  echo "FINISHED running."
  echo "Opening report/log in Notepad"
  notepad "c:\scripts\instances-log.txt"
:batch
rem --------------------------
rem -- END of instances.bat --
rem --------------------------
rem
rem [SAVE below as c:\scripts\call-instances.bat - and change "remote-sqlserverN" to your needs]
rem AFTER you modify "remote-sqlserver1, 2, 3 below to meet your needs, and save
rem   those lines as c:\scripts\call-instances.bat
rem THEN Execute that newly-saved script: c:\scripts\call-instances
rem   and that script will call the above "instances.bat" script, generating a report/log
rem
rem REMOVE 1st "rem" from each line below, BEFORE SAVING those as c:\scripts\call-instances.bat
rem @echo off
rem rem ---------------------
rem rem - call-instances.bat
rem rem ---------------------
rem rem
rem set loc="c:\scripts"
rem del c:\scripts\instances-log.txt > nul 2>&1
rem call %loc%\instances.bat batch remote-sqlserver1
rem call %loc%\instances.bat batch remote-sqlserver2
rem call %loc%\instances.bat batch remote-sqlserver3
rem echo.
rem echo "FINISHED running."
rem echo "Opening report/log in Notepad"
rem notepad "c:\scripts\instances-log.txt"
