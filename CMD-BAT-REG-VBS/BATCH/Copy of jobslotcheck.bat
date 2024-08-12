@echo off

::pushd \\lasargent02\D$\ARGENT\QueueEngine\JOBDEF
goto sloterror

:wait
if not defined mail (
echo %time:~0,5% No Jobslot errors found.
) else (
echo %time:~0,5% - The last Jobslot error occurred at %errtme%.
)
if %time:~0,2%%time:~3,2% GTR 2300 goto exit
sleep 300
 

:sloterror
pushd \\lasargent02\D$\ARGENT\QueueEngine\JOBDEF
for /F "tokens=*" %%A in ('dir /B /A:-D *.js2') do call :rename %%A
popd
if not defined mail goto wait
if %mail%==1 (
call mail.vbs
) & (
set mail=0
)
goto wait

:rename
set js2=%1
set file=%js2:~0,-4%
if not exist %file%.js1 (
copy %js2% %file%.JS1 
) & (
Echo %DATE:~4,10% %time:~0,5% - %file%.js1 was recreated.
) & (
set mail=1
) & (
set errtme=%time:~0,5%
) & (
Echo %DATE:~4,10% %time:~0,5% - Job number %file:~2% was missing a JS1 definition file. %file%.js1 was recreated.  >> "D:\Argent Jobslot Check\Jobslot.log"
)
goto :eof

:exit
::popd
exit
