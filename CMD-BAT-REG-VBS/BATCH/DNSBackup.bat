dnscmd /enumzones > C:\Script\AllZones.txt
for /f %%a in (C:\Script\AllZones.txt) do dnscmd /ZoneExport %%a Export\%%a.dns