@ECHO On
CD C:\Windows\System32
lodctr /R
CD C:\Windows\SysWOW64
lodctr /R
CD C:\Windows\Inf\ASP.NET\
lodctr aspnet_perf2.ini
WINMGMT.EXE /RESYNCPERF

net stop RemoteRegistry
net start RemoteRegistry
net stop pla
net start pla
net stop Winmgmt /yes
net start iphlpsvc
net start UALSVC
net start Winmgmt /yes