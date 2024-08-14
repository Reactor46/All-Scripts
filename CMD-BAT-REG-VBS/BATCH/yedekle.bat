
 @echo off
 echo ### Yedek alinacak surucu (HARDDISK) harfini girin (Where you will take backup)...
 SET /P N=Degerlerinden birisini girin D, E, F, G, H, I, Z sonra ENTER'a basin(Choose your Hard Drive Letter):
 IF "%N%"=="D" GOTO D
 IF "%N%"=="E" GOTO E
 IF "%N%"=="F" GOTO F
 IF "%N%"=="G" GOTO G
 IF "%N%"=="H" GOTO H
 IF "%N%"=="I" GOTO I
 IF "%N%"=="Z" GOTO Z

 set drive=D:\Backup\%date:~0,2%-%date:~3,2%-%date:~6,6%_%time:~0,2%-%time:~3,2%
 set backupcmd=xcopy /s /c /d /e /h /i /r /y
 GOTO SISTEMBUL
 :E
 set drive=E:\Backup\%date:~0,2%-%date:~3,2%-%date:~6,6%_%time:~0,2%-%time:~3,2%
 set backupcmd=xcopy /s /c /d /e /h /i /r /y
 GOTO SISTEMBUL
 :F
 set drive=F:\Backup\%date:~0,2%-%date:~3,2%-%date:~6,6%_%time:~0,2%-%time:~3,2%
 set backupcmd=xcopy /s /c /d /e /h /i /r /y
 GOTO SISTEMBUL
 :G
 set drive=G:\Backup\%date:~0,2%-%date:~3,2%-%date:~6,6%_%time:~0,2%-%time:~3,2%
 set backupcmd=xcopy /s /c /d /e /h /i /r /y
 GOTO SISTEMBUL
 :H
 set drive=H:\Backup\%date:~0,2%-%date:~3,2%-%date:~6,6%_%time:~0,2%-%time:~3,2%
 set backupcmd=xcopy /s /c /d /e /h /i /r /y
 GOTO SISTEMBUL
 :I
 set drive=I:\Backup\%date:~0,2%-%date:~3,2%-%date:~6,6%_%time:~0,2%-%time:~3,2%
 set backupcmd=xcopy /s /c /d /e /h /i /r /y
 GOTO SISTEMBUL
 :Z
 set drive=Z:\Backup\%date:~0,2%-%date:~3,2%-%date:~6,6%_%time:~0,2%-%time:~3,2%
 set backupcmd=xcopy /s /c /d /e /h /i /r /y
 GOTO SISTEMBUL

 :SISTEMBUL
 echo ### ISLETIM SISTEMI BULUNUYOR(Finding OS)...
 VER |find "6.2" >NUL
 IF NOT ERRORLEVEL 1 goto :WIN8

 echo ### ISLETIM SISTEMI BULUNUYOR(Finding OS)...
 VER |find "6.1" >NUL
 IF NOT ERRORLEVEL 1 goto :WIN7

 VER |find "5.1" >NUL
 IF NOT ERRORLEVEL 1 goto :WINXP


:WIN8
 echo ### GECERLI SISTEM (Current OS) WINDOWS 8 - AKTIF PROFIL YEDEKLENIYOR(Profile Backup)...
 %backupcmd% "%USERPROFILE%\Documents" "%drive%\Documents"
 %backupcmd% "%USERPROFILE%\Pictures" "%drive%\Pictures"
 %backupcmd% "%USERPROFILE%\Downloads" "%drive%\Downloads"
 %backupcmd% "%USERPROFILE%\Music" "%drive%\Music"
 %backupcmd% "%USERPROFILE%\Videos" "%drive%\Videos"
 %backupcmd% "%USERPROFILE%\Contacts" "%drive%\Contacts"
 %backupcmd% "%USERPROFILE%\Saved Games" "%drive%\Saved Games"
 %backupcmd% "%USERPROFILE%\Links" "%drive%\Links"
 %backupcmd% "%USERPROFILE%\Desktop" "%drive%\Desktop"
 %backupcmd% "%USERPROFILE%\Favorites" "%drive%\Favorites"
 echo ### MAIL PROGRAMLARI YEDEKLERI ALINIYOR...
 %backupcmd% "%USERPROFILE%\AppData\Local\Microsoft\Windows Live Mail" "%drive%\Windows Live Mail"
 %backupcmd% "%USERPROFILE%\AppData\Roaming\Identities" "%drive%\Outlook Express"
 %backupcmd% "%USERPROFILE%\AppData\Local\Microsoft\Outlook" "%drive%\Outlook"
 GOTO YAZILIMYEDEK

 :WIN7
 echo ### GECERLI SISTEM (Current OS) WINDOWS 7 - AKTIF PROFIL YEDEKLENIYOR(Profile Backup)...
 %backupcmd% "%USERPROFILE%\Documents" "%drive%\Documents"
 %backupcmd% "%USERPROFILE%\Pictures" "%drive%\Pictures"
 %backupcmd% "%USERPROFILE%\Downloads" "%drive%\Downloads"
 %backupcmd% "%USERPROFILE%\Music" "%drive%\Music"
 %backupcmd% "%USERPROFILE%\Videos" "%drive%\Videos"
 %backupcmd% "%USERPROFILE%\Contacts" "%drive%\Contacts"
 %backupcmd% "%USERPROFILE%\Saved Games" "%drive%\Saved Games"
 %backupcmd% "%USERPROFILE%\Links" "%drive%\Links"
 %backupcmd% "%USERPROFILE%\Desktop" "%drive%\Desktop"
 %backupcmd% "%USERPROFILE%\Favorites" "%drive%\Favorites"
 echo ### MAIL PROGRAMLARI YEDEKLERI ALINIYOR...
 %backupcmd% "%USERPROFILE%\AppData\Local\Microsoft\Windows Live Mail" "%drive%\Windows Live Mail"
 %backupcmd% "%USERPROFILE%\AppData\Roaming\Identities" "%drive%\Outlook Express"
 %backupcmd% "%USERPROFILE%\AppData\Local\Microsoft\Outlook" "%drive%\Outlook"
 GOTO YAZILIMYEDEK

 :WINXP
 echo ### GECERLI SISTEM (Current OS) WINDOWS XP - PROFIL YEDEKLENIYOR(Profile Backup)...
 %backupcmd% "%USERPROFILE%\My Documents" "%drive%\My Documents"
 %backupcmd% "%USERPROFILE%\Desktop" "%drive%\Desktop"
 %backupcmd% "%USERPROFILE%\Favorites" "%drive%\Favorites"
 echo ### MAIL PROGRAMLARI YEDEKLERI ALINIYOR...
 %backupcmd% "%USERPROFILE%\Application Data\Microsoft\Address Book" "%drive%\Address Book"
 %backupcmd% "%USERPROFILE%\Application Data\Microsoft\Windows Live Mail" "%drive%\Windows Live Mail"
 %backupcmd% "%USERPROFILE%\Local Settings\Application Data\Identities" "%drive%\Outlook Express"
 %backupcmd% "%USERPROFILE%\Local Settings\Application Data\Microsoft\Outlook" "%drive%\Outlook"
 GOTO YAZILIMYEDEK

 :YAZILIMYEDEK
 echo ### KAYIT DEFTERI YEDEKLENIYOR(regedit backup)...
 if not exist "%drive%\Registry" mkdir "%drive%\Registry"
 if exist "%drive%\Registry\regbackup.reg" del "%drive%\Registry\regbackup.reg"
 regedit /e "%drive%\Registry\regbackup.reg"

 echo YEDEKLEME TAMAMLANDI(Backup is completed)!
 @pause 