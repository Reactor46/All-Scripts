@echo off
set currdir=%~dp0
IF "%currdir:~-1%"=="\" SET currdir=%currdir:~0,-1%
cd /d "%currdir%"

C:\solr\bin\nssm stop solr
C:\solr\bin\nssm remove solr confirm

:END
echo bye
timeout /t 3
