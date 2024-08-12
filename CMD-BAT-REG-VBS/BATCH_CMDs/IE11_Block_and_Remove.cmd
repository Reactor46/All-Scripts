::Block IE11 Install  
reg add "HKLM\SOFTWARE\Microsoft\Internet Explorer\Setup\11.0" /v DoNotAllowIE11 /t REG_DWORD /d 1 /f 

net stop "Windows Update"
net stop "Background Intelligent Transfer Service"
del /f /s /q %windir%\SoftwareDistribution\*.*

:Uninstall_IE11
FORFILES /P %WINDIR%\servicing\Packages /M Microsoft-Windows-InternetExplorer-*11.*.mum /c "cmd /c echo Uninstalling package @fname && start /w pkgmgr /up:@fname /quiet"