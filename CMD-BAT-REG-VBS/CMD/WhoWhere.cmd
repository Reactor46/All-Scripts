@echo off
setlocal
for /f "Tokens=1" %%c in ('net view /domain:"%USERDOMAIN%"^|Findstr /L /C:"\\"') do (
 for /f "Tokens=*" %%u in ('PsLoggedOn -L %%c^|find /i "%USERDOMAIN%\"') do (
  call :report %%c "%%u"
 )
)
endlocal
goto :EOF

:report
set work=%1
set comp=%work:~2%
set user=%2
set user=%user:"=%
call set user=%%user:*%USERDOMAIN%\=%%
@echo %comp% %user%