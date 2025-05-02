cls

set startDate=%date%
set startTime=%time%

set sdy=%startDate:~10%
set /a sdm=1%startDate:~4,2% - 100
set /a sdd=1%startDate:~7,2% - 100
set /a sth=%startTime:~0,2%
set /a stm=1%startTime:~3,2% - 100
set /a sts=1%startTime:~6,2% - 100
set timestamp=%sdm%-%sdd%-%sdy%-%sth%-%stm%-%sts%

time /t

powershell.exe Set-ExecutionPolicy RemoteSigned
powershell.exe -noexit .\UploadMasterPages.ps1 "http://sitecollection"

time /t