@echo off
set currdir=%~dp0
IF "%currdir:~-1%"=="\" SET currdir=%currdir:~0,-1%
cd /d "%currdir%"

set dlver=17
set dlfile=jdk-%dlver%_windows-x64_bin.exe
set dlurl=https://download.oracle.com/java/17/latest/%dlfile%

curl --fail --output %dlfile% %dlurl%
if %errorlevel% neq 0 (
    echo Fail to download %dlfile%
    pause
    goto End
)

if not exist %dlfile% (
    echo %dlfile% not found
    pause
    goto End
)

%dlfile% /s

:END
echo bye
timeout /t 3
