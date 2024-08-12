del ADHealth.txt


echo ================================ >> ADHealth.txt

echo Domain Controllers In the Domain >> ADHealth.txt

echo ================================ >> ADHealth.txt

C:\Windows\System32\DSQUERY Server -o rdn >> ADHealth.txt



echo ====================== >> ADHealth.txt

echo Repadmin - Syncall - e >> ADHealth.txt

echo ====================== >> ADHealth.txt

C:\Windows\System32\repadmin.exe /syncall /e >> ADHealth.txt



echo ====================== >> ADHealth.txt

echo Repadmin - Syncall - a >> ADHealth.txt

echo ====================== >> ADHealth.txt

C:\Windows\System32\repadmin.exe /syncall /A >> ADHealth.txt



echo ====================== >> ADHealth.txt

echo Repadmin - Syncall - d >> ADHealth.txt

echo ====================== >> ADHealth.txt

C:\Windows\System32\repadmin.exe /syncall /d >> ADHealth.txt


echo ====================== >> ADHealth.txt

echo Repadmin - Replsummary >> ADHealth.txt

echo ====================== >> ADHealth.txt

C:\Windows\System32\repadmin.exe /replsummary * >> ADHealth.txt



echo ============== >> ADHealth.txt

echo Repadmin - KCC >> ADHealth.txt

echo ============== >> ADHealth.txt

C:\Windows\System32\repadmin.exe /kcc * >> ADHealth.txt


echo ===================== >> ADHealth.txt

echo Repadmin - showbackup >> ADHealth.txt

echo ===================== >> ADHealth.txt

C:\Windows\System32\repadmin.exe /showbackup * >> ADHealth.txt


echo =================== >> ADHealth.txt

echo Repadmin - Showrepl >> ADHealth.txt

echo =================== >> ADHealth.txt

C:\Windows\System32\repadmin.exe /showrepl *  >> ADHealth.txt



echo ================ >> ADHealth.txt

echo Repadmin - Queue >> ADHealth.txt

echo ================ >> ADHealth.txt

C:\Windows\System32\repadmin.exe /queue *  >> ADHealth.txt



echo ====================== >> ADHealth.txt

echo Repadmin - Bridgeheads >> ADHealth.txt

echo ====================== >> ADHealth.txt

C:\Windows\System32\repadmin.exe /bridgeheads * /verbose >> ADHealth.txt



echo =============== >> ADHealth.txt

echo Repadmin - ISTG >> ADHealth.txt

echo =============== >> ADHealth.txt

C:\Windows\System32\repadmin.exe /istg * /verbose >> ADHealth.txt



echo ======================= >> ADHealth.txt

echo Repadmin - Showoutcalls >> ADHealth.txt

echo ======================= >> ADHealth.txt

C:\Windows\System32\repadmin.exe /showoutcalls * >> ADHealth.txt



echo ==================== >> ADHealth.txt

echo Repadmin - Failcache >> ADHealth.txt

echo ==================== >> ADHealth.txt

C:\Windows\System32\repadmin.exe /failcache * >> ADHealth.txt



echo ==================== >> ADHealth.txt

echo Repadmin - Showtrust >> ADHealth.txt

echo ==================== >> ADHealth.txt

C:\Windows\System32\repadmin.exe /showtrust * >> ADHealth.txt



echo =============== >> ADHealth.txt

echo Repadmin - Bind >> ADHealth.txt

echo =============== >> ADHealth.txt

C:\Windows\System32\repadmin.exe /bind * >> ADHealth.txt




echo ====== >> ADHealth.txt

echo Dcdiag >> ADHealth.txt

echo ====== >> ADHealth.txt

C:\Windows\System32\dcdiag /c /e /v >> ADHealth.txt

start ADHealth.txt





















