cd /d %windir%\System32\Wbem
net stop winmgmt

sc sdset winmgmt D:(A;;CCDCLCSWRPWPDTLOCRRC;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;DA)(A;;CCDCLCSWRPWPDTLOCRRC;;;PU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;SY)

REM REG IMPORT %windir%\WBEM.reg

winmgmt /clearadap
winmgmt /kill
winmgmt /unregserver
winmgmt /regserver
winmgmt /resyncperf

del %windir%\System32\Wbem\Repository /Q
del %windir%\System32\Wbem\AutoRecover /Q

for %%i in (*.dll) do Regsvr32 -s %%i
for %%i in (*.mof,*.mfl) do Mofcomp %%i
wmiadap.exe /Regsvr32
wmiapsrv.exe /Regsvr32
wmiprvse.exe /Regsvr32

net start winmgmt