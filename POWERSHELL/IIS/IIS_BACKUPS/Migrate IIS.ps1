# PowerShell Migration/Backup Scripts for IIS
# Script 1
$source = Read-Host "Source IIS Server with Sites"
$target = Read-Host "Target IIS Server"

$hostsession = New-PSSession -ComputerName $Source 
start-process -filepath C:\Windows\System32\inetsrv\appcmd.exe -argumentlist "list apppool /config /xml > C:\Staging"
Remove-PSSession $hostsession

\\$source\c$\Windows\system32\inetsrv\appcmd.exe list site /config /xml  > C:\Staging

start-process -filepath \\$target\C$\Windows\system32\inetsrv\appcmd.exe -argumentlist "add apppool /in < C:\Staging\apppools.xml"
start-process -filepath \\$target\C$\Windows\system32\inetsrv\appcmd.exe -argumentlist "add site /in < C:\Staging\websites.xml"

Write-Host "Sites Migrated to Target Server Successfully."
# End Script 1