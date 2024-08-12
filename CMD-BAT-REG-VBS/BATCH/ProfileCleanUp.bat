@ECHO OFF
::Tested on 2000/XP/VISTA/Windows 7 only

if not exist "%SYSTEMDRIVE%\IT\ProfileCleanUp.bat" goto COPY

:WinVersion
cls
echo ## Definig Windows Version
ver>"%temp%\ver.tmp"
find /i "4.0" "%temp%\ver.tmp">nul
if %ERRORLEVEL% EQU 0 set WinVersion=WinNT4
find /i "5.0" "%temp%\ver.tmp">nul
if %ERRORLEVEL% EQU 0 set WinVersion=Win2k
find /i "5.1" "%temp%\ver.tmp">nul
if %ERRORLEVEL% EQU 0 set WinVersion=WinXP
find /i "5.2" "%temp%\ver.tmp">nul
if %ERRORLEVEL% EQU 0 set WinVersion=Win2k3
find /i "6.0" "%temp%\ver.tmp">nul
if %ERRORLEVEL% EQU 0 set WinVersion=WinVista/Server2008
find /i "6.1" "%temp%\ver.tmp">nul
if %ERRORLEVEL% EQU 0 set WinVersion=Win7/Server2008R2
if "%WinVersion%" EQU "" set WinVersion=UNKNOWN
if %WinVersion% EQU UNKNOWN goto WARN0
if %WinVersion% EQU WinVista/Server2008 goto WARN1
if %WinVersion% EQU Win7/Server2008R2 goto WARN1
goto START

:START
cls
%homedrive% 
cd %USERPROFILE%
cd..
set profiles=%cd%

for /f "tokens=* delims= " %%u in ('dir /b/ad') do (

cls
title Deleting %%u Cookies. . .
if exist "%profiles%\%%u\cookies" echo Deleting....
if exist "%profiles%\%%u\cookies" cd "%profiles%\%%u\cookies"
if exist "%profiles%\%%u\cookies" del *.* /F /S /Q /A: R /A: H /A: A

cls
title Deleting %%u Temp Files. . .
if exist "%profiles%\%%u\Local Settings\Temp" echo Deleting....
if exist "%profiles%\%%u\Local Settings\Temp" cd "%profiles%\%%u\Local Settings\Temp"
if exist "%profiles%\%%u\Local Settings\Temp" del *.* /F /S /Q /A: R /A: H /A: A
if exist "%profiles%\%%u\Local Settings\Temp" rmdir /s /q "%profiles%\%%u\Local Settings\Temp"

cls
title Deleting %%u Temp Files. . .
if exist "%profiles%\%%u\AppData\Local\Temp" echo Deleting....
if exist "%profiles%\%%u\AppData\Local\Temp" cd "%profiles%\%%u\AppData\Local\Temp"
if exist "%profiles%\%%u\AppData\Local\Temp" del *.* /F /S /Q /A: R /A: H /A: A
if exist "%profiles%\%%u\AppData\Local\Temp" rmdir /s /q "%profiles%\%%u\AppData\Local\Temp"

cls
title Deleting %%u Temporary Internet Files. . .
if exist "%profiles%\%%u\Local Settings\Temporary Internet Files" echo Deleting....
if exist "%profiles%\%%u\Local Settings\Temporary Internet Files" cd "%profiles%\%%u\Local Settings\Temporary Internet Files"
if exist "%profiles%\%%u\Local Settings\Temporary Internet Files" del *.* /F /S /Q /A: R /A: H /A: A
if exist "%profiles%\%%u\Local Settings\Temporary Internet Files" rmdir /s /q "%profiles%\%%u\Local Settings\Temporary Internet Files"

cls
title Deleting %%u Temporary Internet Files. . .
if exist "%profiles%\%%u\AppData\Local\Microsoft\Windows\Temporary Internet Files" echo Deleting....
if exist "%profiles%\%%u\AppData\Local\Microsoft\Windows\Temporary Internet Files" cd "%profiles%\%%u\AppData\Local\Microsoft\Windows\Temporary Internet Files"
if exist "%profiles%\%%u\AppData\Local\Microsoft\Windows\Temporary Internet Files" del *.* /F /S /Q /A: R /A: H /A: A
if exist "%profiles%\%%u\AppData\Local\Microsoft\Windows\Temporary Internet Files" rmdir /s /q "%profiles%\%%u\AppData\Local\Microsoft\Windows\Temporary Internet Files"

)

cls
title Deleting %Systemroot%\Temp
if exist "%Systemroot%\Temp" echo Deleting....
if exist "%Systemroot%\Temp" cd "%Systemroot%\Temp" 
if exist "%Systemroot%\Temp" del *.* /F /S /Q /A: R /A: H /A: A
if exist "%Systemroot%\Temp" rmdir /s /q "%Systemroot%\Temp"

cls
title Deleting %SYSTEMDRIVE%\Temp
if exist "%SYSTEMDRIVE%\Temp" echo Deleting....
if exist "%SYSTEMDRIVE%\Temp" cd "%SYSTEMDRIVE%\Temp"
if exist "%SYSTEMDRIVE%\Temp" del *.* /F /S /Q /A: R /A: H /A: A
if exist "%SYSTEMDRIVE%\Temp" rmdir /s /q "%Systemroot%\Temp"

cls
goto END

:WARN0
cls
title Warning
echo This program has only been tested for use on:
echo Windows 2000
echo Windows XP
echo Windows Vista
echo Windows 7
echo.
echo Continue at your own risk!
echo.
echo Press 'Y' to continue or any other key to exit.
echo.
Set /P input=
if /I %input% EQU Y goto :START
goto END

:WARN1
cls
title Warning
echo For this program to work successfully be sure to
echo Right Click and select
echo Run as Administrator
echo.
echo If you have already done so ignore this warning.
echo.
echo Press 'Y' to continue or any other key to exit.
echo.
Set /P input=
if /I %input% EQU Y goto :START
goto END

:COPY
if not exist "%SYSTEMDRIVE%\IT" md "%SYSTEMDRIVE%\IT"
copy "ProfileCleanUp.bat" "%SYSTEMDRIVE%\IT\ProfileCleanUp.bat"
if not exist "%SYSTEMDRIVE%\IT\ProfileCleanUp.bat" goto FAIL
cls
goto START

:FAIL
cls
color 4F
echo.
echo "%SYSTEMDRIVE%\IT\ProfileCleanUp.bat" not found
echo.
echo Copy ProfileCleanUp.bat to "%SYSTEMDRIVE%\IT"
pause
goto END

:END
exit