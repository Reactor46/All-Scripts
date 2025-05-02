@ECHO off
SETLOCAL 
::
:: By Gastone Canali
:: v.02 28.01.2012
title=UpToDate-sysinternals.cmd
set filename=%~n0
set logfile="%temp%\_%filename%.txt"
set null= 2^>^&1
set log=^>^>%logfile% 2^>^&1
::set log=

::SET path=c:\path\ToRobocopy;%path%
SET SysIntShare=\\live.sysinternals.com\tools
SET SysIntFolder=c:\sysinternals

PushD %SysIntFolder% || goto :_SYSFOLDERERR
if not exist "%SysIntShare%" goto :_SysIntShareERR
robocopy . . .?.?.? /w:0 /r:0 >nul 2>nul && (
   robocopy "%SysIntShare%"  "%SysIntFolder%" /w:1 /r:1 /xf Thumbs.db 
 ) || (
 goto :_NoRobocopy )

goto :EOF

:_NoRobocopy
 FOR /F "skip=7 tokens=4,*"  %%F in ('dir /A-D  "%SysIntShare%" ^|sort') DO (
     IF exist "%%F" (
       echo %%F Exist chk date
       FOR /F "skip=7 tokens=*" %%R in ('dir /A-D  "%SysIntShare%\%%F" ^|sort') DO (
          FOR /F "skip=7 tokens=*" %%L in ('dir /A-D  "%SysIntFolder%\%%F" ^|sort') DO (
             echo %%R|find /i "%%L" || echo Local %%F is older  & xcopy "%SysIntShare%\%%F" "%SysIntFolder%\%%F" /y /c /r /q 
          )
       ) 
     ) ELSE (  echo %%F not present & xcopy "%SysIntShare%\%%F" . /y /c /r /q )   
   )%log%
goto :EOF
:_SYSFOLDERERR
    echo ERROR:%SysIntFolder% Not Found
goto :EOF
:_SysIntShareERR
    echo ERROR:%SysIntShare%   Not Found 
goto :EOF
