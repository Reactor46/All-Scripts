@echo on
@echo Please read:
@echo -----------------------------------------
@echo:
@echo This totally resets all of your Windows Update Agent settings.
@echo:
@echo Many times, the computer will do a full reset and will not be able to
@echo install updates for the rest of the day. This is so that the server
@echo does not get overutilized because of the reset.
@echo:
@echo If you don't receive any updates after this script runs, please
@echo wait until tomorrow.
@echo:
@echo Re-running this script will reset the PC again and it will have
@echo to wait again.
@echo:

net stop bits
net stop wuauserv
net stop appidsvc
net stop cryptsvc
regsvr32 /u wuaueng.dll /s

sc.exe sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)
sc.exe sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)

@echo Deleting AU cache...
del /f /s /q %windir%\SoftwareDistribution\*.*
del /f /s /q %windir%\windowsupdate.log
if exist "%systemroot%\system32\catroot2.bak" ren %systemroot%\system32\catroot2.bak catroot2
if exist "%systemroot%\system32\catroot2.old" rd /q /s %systemroot%\system32\catroot2.old
md %systemroot%\system32\catroot2.old
xcopy %systemroot%\system32\catroot2 %systemroot%\system32\catroot2.old /s /Y

@echo Registering DLLs...
regsvr32.exe /s atl.dll
regsvr32.exe /s urlmon.dll
regsvr32.exe /s mshtml.dll
regsvr32.exe /s shdocvw.dll
regsvr32.exe /s browseui.dll
regsvr32.exe /s jscript.dll
regsvr32.exe /s vbscript.dll
regsvr32.exe /s scrrun.dll
regsvr32.exe /s msxml.dll
regsvr32.exe /s msxml3.dll
regsvr32.exe /s msxml6.dll
regsvr32.exe /s actxprxy.dll
regsvr32.exe /s softpub.dll
regsvr32.exe /s wintrust.dll
regsvr32.exe /s dssenh.dll
regsvr32.exe /s rsaenh.dll
regsvr32.exe /s gpkcsp.dll
regsvr32.exe /s sccbase.dll
regsvr32.exe /s slbcsp.dll
regsvr32.exe /s cryptdlg.dll
regsvr32.exe /s oleaut32.dll
regsvr32.exe /s ole32.dll
regsvr32.exe /s shell32.dll
regsvr32.exe /s initpki.dll
regsvr32.exe /s wuapi.dll
regsvr32.exe /s wuaueng.dll
regsvr32.exe /s wuaueng1.dll
regsvr32.exe /s wucltui.dll
regsvr32.exe /s wups.dll
regsvr32.exe /s wups2.dll
regsvr32.exe /s wuweb.dll
regsvr32.exe /s qmgr.dll
regsvr32.exe /s qmgrprxy.dll
regsvr32.exe /s wucltux.dll
regsvr32.exe /s muweb.dll
regsvr32.exe /s wuwebv.dll
regsvr32.exe /s wudriver.dll
regsvr32.exe /s wuaueng.dll 
regsvr32.exe /s wuaueng1.dll 


net start bits
net start wuauserv
net start appidsvc
net start cryptsvc

@echo Checking in...
@echo:
@echo It's possible the server will not release the updates in
@echo just one session, so it's ok if this script does not immediately
@echo install updates.
@echo:
@echo This is due to the full reset on this PC. Just let it be for a few
@echo hours and updates should resume as normal.
wuauclt.exe /resetauthorization /detectnow