@ECHO ON

net stop "Windows Update"
net stop "Background Intelligent Transfer Service"
del /f /s /q %windir%\SoftwareDistribution\*.*
net start "Background Intelligent Transfer Service"
net start "Windows Update"
wuauclt.exe /resetauthorization /detectnow 
gpupdate /force