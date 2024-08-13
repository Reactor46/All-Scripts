@echo off

cd "C:\NG2ServerStatus"

del /f /q C:\NG2ServerStatus\assets\servers.json
del /f /q C:\NG2ServerStatus\assets\ping.json

C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe "C:\NG2ServerStatus\info.ps1" -ExecutionPolicy Unrestricted
