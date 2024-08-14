@echo off
set currdir=%~dp0
IF "%currdir:~-1%"=="\" SET currdir=%currdir:~0,-1%
cd /d "%currdir%"

set dlver=8.10.1
set dlfile=solr-%dlver%.zip
set dlurl=https://archive.apache.org/dist/lucene/solr/8.10.1/%dlfile%
set foldername=solr-%dlver%

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

powershell -command "Expand-Archive -Force %dlfile% ."

if not exist %foldername%\bin\solr (
    echo Fail to extract %dlfile%
    pause
    goto End
)

rename %foldername% solr

:END
echo bye
timeout /t 3
