@ECHO OFF

netsh advfirewall set allprofiles firewallpolicy allowinbound,allowoutbound

netsh advfirewall set allprofiles state off

taskkill -f -im ERAAgent.exe
net start "EraAgentSvc"