@echo on

set count=0
set mail=0
goto slotst

:wait
set count=0
if %mail% == 0 (
echo %time:~0,5% No Jobslot errors found.
) else (
echo %time:~0,5% - The last Jobslot error occurred at %errtme%.
)
if %time:~0,2%%time:~3,2% GTR 2300 goto exit
sleep 30
 
:slotst
for /F "tokens=*" %%G in ('dir /B /A:-D \\lasargent02\D$\ARGENT\QueueEngine\JOBDEF\*.js2') do call :tst %%G
if defined count (if %count% gtr 2 (goto sloterror) else (set count=0) & (goto wait)) else (set count=0) & (goto wait)

:tst
set js2=%1
set file=%js2:~0,-4%
if not exist \\lasargent02\D$\ARGENT\QueueEngine\JOBDEF\%file%.js1 (if not defined count (set count=1) else (set /a count+=1))
echo %count%
goto :eof


 
:sloterror
call :ampm
echo. >> "%~dp0Jobslot.log"
echo. >> "%~dp0Jobslot.log"
echo --- %DATE:~4,10% %tampm% --->> "%~dp0Jobslot.log"
echo %count% Jobslot errors detected:>> "%~dp0Jobslot.log"
echo -------------------------->> "%~dp0Jobslot.log"
pushd \\lasargent02\D$\ARGENT\QueueEngine\
for /F "tokens=*" %%H in ('dir /B /A:-D JOBDEF\*.js2') do call :rename %%H
popd
echo. >> "%~dp0Jobslot.log"
echo Please verify through Argent logs that the listed jobs are running.>> "%~dp0Jobslot.log"
echo Please track job dependencies.>> "%~dp0Jobslot.log"
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
if not exist JOBDEF\%file%.js1 for /F "tokens=*" %%I in ('dir /B /A:-D JOBLOG\*%file:*J_=%*') do call :gjbnme %%I
if not exist JOBDEF\%file%.js1 (
if %mail%==0 set mail=1
) & (
set errtme=%time:~0,5%
) & (
Echo %jobname% [%file:~2%]>> "%~dp0Jobslot.log"
) & (
copy JOBDEF\%js2% JOBDEF\%file%.JS1 
)
goto :eof

:exit
ren log.txt log.%DATE:~4,10%%time:~0,2%%time:~3,2%
exit /1

:gjbnme
set joblogf=%1
set jobname=%joblogf:~0,-12%
echo %jobname%
goto :eof


:ampm
if %time:~0,2% gtr 12 (set /a hr=%time:~0,2% - 12)
if %time:~0,2% gtr 12 (set tampm=%hr%%time:~2,3% PM) else (set tampm=%time:~0,5% AM)
goto :eof

:hou
set /a hr=%time:~0,2% - 12
echo %hr%
goto :eof

