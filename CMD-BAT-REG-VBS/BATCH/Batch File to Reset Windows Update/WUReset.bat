@ECHO OFF
echo.
echo =============================================
echo Simple Script to Reset / Clear Windows Update
echo.
echo Created by "Vishal Gupta" for AskVG.com
echo.
echo Thanks to DTM for sharing the commands
echo =============================================
echo.
PAUSE
echo.
attrib -h -r -s %windir%\system32\catroot2
attrib -h -r -s %windir%\system32\catroot2\*.*
echo Stopping Windows Update related services...
echo ===========================================
echo.
net stop wuauserv
net stop CryptSvc
net stop BITS
echo Done!
echo.
echo Renaming Windows Update related folders...
echo ==========================================
echo.
ren %windir%\system32\catroot2 catroot2.old 
ren %windir%\SoftwareDistribution sold.old
ren "%ALLUSERSPROFILE%\application data\Microsoft\Network\downloader" downloader.old
echo Done!
echo.
echo Starting Windows Update related services...
echo ===========================================
echo.
net Start BITS
net start CryptSvc
net start wuauserv
echo Done!
echo.
echo Task completed successfully...
echo.
PAUSE