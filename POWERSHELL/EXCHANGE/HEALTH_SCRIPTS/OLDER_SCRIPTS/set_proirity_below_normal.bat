echo
wmic process where name="jusched.exe" CALL setpriority 16384
wmic process where name="ccSvcHst.exe" CALL setpriority 16384
wmic process where name="TrustedInstaller.exe" CALL setpriority 16384
wmic process where name="wuauserv.exe" CALL setpriority 16384
sc config trustedinstaller start= demand
sc config wuauserv start= demand
exit