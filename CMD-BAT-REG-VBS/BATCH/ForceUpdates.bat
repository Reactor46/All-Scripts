@ECHO OFF

ECHO INSTALLING WINDOWS UPDATES
CD C:\Patches

start /wait cscript updatehf_v2.8a.vbs action:install mode:silent email:winsysadmin@creditone.com;noc@creditone.com restart:1


