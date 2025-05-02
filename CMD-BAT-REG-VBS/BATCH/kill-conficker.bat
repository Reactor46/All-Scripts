net stop server /y
KidoKiller.exe -y
AT /Delete /Yes
net stop “Task Scheduler”
sc stop “srservice”
sc config “srservice” start= disabled 
cacls “c:\System Volume Information” /E /G %username%:F 
rd “c:\System Volume Information” /s /q
sc config “wuauserv” start= auto
sc config “bits” start= demand
sc config “ersvc” start= auto
net user administrator ******
reg.exe add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\SHOWALL /v CheckedValue /t REG_DWORD /d 0×1 /f 
reg.exe add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer /v NoDriveTypeAutoRun /t REG_DWORD /d 0xff /f
reg.exe add HKLM\SYSTEM\CurrentControlSet\Services\lanmanserver\parameters /v AutoShareWks /t REG_DWORD /d 0×00 /f
WindowsXP-KB958644-x86-ENU.exe /passive /norestart
WindowsXP-KB957097-x86-ENU.exe /passive /norestart
WindowsXP-KB958687-x86-ENU.exe /passive /norestart
windows-kb890830-v2.7.exe
cls

echo off

echo You need to restart your computer as soon as possible!

pause