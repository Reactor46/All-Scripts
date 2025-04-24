# APPCMD BACKUP:
D:
CD D:\IIS_BACKUPS
%windir%\system32\inetsrv\appcmd add backup VM119IS01WEBP6_10-09-2023
%windir%\system32\inetsrv\appcmd list apppool /config /xml > D:\IIS_BACKUPS\VM119IS01WEBP6_10-09-2023_apppools.xml
%windir%\system32\inetsrv\appcmd list site /config /xml > D:\IIS_BACKUPS\VM119IS01WEBP6_10-09-2023_sites.xml


# APPCMD RESTORE:
%windir%\system32\inetsrv\appcmd list backup
%windir%\system32\inetsrv\appcmd restore backup /stop:true VM119IS01WEBP6_10-09-2023

# APPCMD Export & Import App Pools:
%windir%\system32\inetsrv\appcmd add apppool /in < D:\IIS_BACKUPS\VM119IS01WEBP6\VM119IS01WEBP6-APP-POOLS.XML
%windir%\system32\inetsrv\appcmd add site /in < D:\IIS_BACKUPS\VM119IS01WEBP6\VM119IS01WEBP6_10-09-2023_sites.xml



%SystemRoot%\System32\inetsrv\appcmd add apppool /in < "D:\IIS_BACKUPS\IIS_Configs\IIS_APP_POOLS.xml"
%SystemRoot%\System32\inetsrv\appcmd add site /in < "D:\IIS_BACKUPS\ALL_Websites.xml"
iisreset
