REM         Copyright (C) 2013 Microsoft Corporation
REM
REM You have a royalty-free right to use, modify, reproduce and distribute
REM the Sample Application Files (and/or any modified version) in any way
REM you find useful, provided that you agree that Microsoft has no warranty,
REM obligations or liability for any Sample Application Files.

@echo off
for /f "delims= " %%a in ('"wmic path win32_useraccount where name='%UserName%' get sid"') do (
   if not "%%a"=="SID" (          
      set USID=%%a
      goto :loop_end
   )   
)

:loop_end
REM echo %%USID%%=%USID%
reg delete HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice /ve /f
reg delete HKEY_CURRENT_USER\Software\Classes\https\DefaultIcon" /ve /f
REG ADD HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice /v Progid /t REG_SZ /d IE.AssocFile.HTM /f
REG ADD HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.htm\UserChoice /v Progid /t REG_SZ /d IE.AssocFile.HTM /f
REG ADD HKEY_USERS\%USID%\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\ftp\UserChoice /v Progid /t REG_SZ /d IE.FTP /f
REG ADD HKEY_USERS\%USID%\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice /v Progid /t REG_SZ /d IE.HTTP /f
REG ADD HKEY_USERS\%USID%\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice /v Progid /t REG_SZ /d IE.HTTPS /f
