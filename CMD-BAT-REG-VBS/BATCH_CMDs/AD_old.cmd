:::::::::::: How to use this script file ::::::::::::
:: rootOU is the DN for your site ie. OU=ComputerOU,DC=ad,DC=company,DC=com
:: inactiveTime is the number of weeks that a computer can be inactive before being show in this list
:: No other edits need to be done to this file

@echo off
cls

:::::::::::: Your Settings Here ::::::::::::

set rootOU=DC=domain,DC=local
set inactiveTime=1

:::::::::::: Do not edit below this line ::::::::::::



set file=%temp%\info.txt
echo Computers that have been inactive for greater then %inactiveTime% weeks: >%file%
dsquery computer -inactive %inactiveTime% -limit 0 %rootOU% >>%file%
echo.>>%file%
echo.>>%file%
echo.>>%file%

echo Computers that are disabled:>>%file%
dsquery computer -disabled %rootOU%>>%file%
echo.>>%file%
echo.>>%file%
echo.>>%file%

::echo Users that are disabled:>>%file%
::dsquery user %rootOU% -disabled -limit 1000 -o samid>>%file%

start "" "%file%"