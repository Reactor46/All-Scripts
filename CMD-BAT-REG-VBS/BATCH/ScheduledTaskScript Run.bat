@echo off

rem Parse today's date into month, day, year, hour, min
set tempdate=%date%
set tempdate=%tempdate:~4,15%
set month=%tempdate:~0,2%
set day=%tempdate:~3,2%
set year=%tempdate:~6,4%
set temptime=%time%
set hour=%temptime:~0,2%
set hour=%hour: =%
set min=%temptime:~3,2%

rem Un-remark the following lines for debugging
rem echo %tempdate%
rem echo %month%
rem echo %day%
rem echo %year%
rem echo %hour%
rem echo %min%

C:
CD \
CD Scripts
CD Project
CD UserProvisioning


call powershell -File C:\Scripts\Project\ProvisionUsers.ps1 > log.txt
REM ***OLD****
REM if not exist C:\Scripts\Project\Logs\%month%_%day%_%year%_%hour%_%min% md C:\Scripts\Project\Logs\%month%_%day%_%year%_%hour%_%min%
REM move *.csv C:\Scripts\Project\Logs\%month%_%day%_%year%_%hour%_%min%
REM move *.txt C:\Scripts\Project\Logs\%month%_%day%_%year%_%hour%_%min%

if not exist C:\Scripts\Project\Logs\%year% md C:\Scripts\Project\Logs\%year%
if not exist C:\Scripts\Project\Logs\%year%\%month% md C:\Scripts\Project\Logs\%year%\%month%
if not exist C:\Scripts\Project\Logs\%year%\%month%\%day% md C:\Scripts\Project\Logs\%year%\%month%\%day%
if not exist C:\Scripts\Project\Logs\%year%\%month%\%day%\%hour%_%min% md C:\Scripts\Project\Logs\%year%\%month%\%day%\%hour%_%min%
move *.csv C:\Scripts\Project\Logs\%year%\%month%\%day%\%hour%_%min%
move *.txt C:\Scripts\Project\Logs\%year%\%month%\%day%\%hour%_%min%

echo Batch Finished

:end