@echo off

echo. 
echo Gathering Report for DCLIST = %1 
echo. 
Echo Report for DCLIST = %1 > replreport.txt

echo. >> replreport.txt 
echo. >> replreport.txt

echo Gathering Verbose Replication and Connections 
echo Verbose Replication and Connections >> replreport.txt echo. >> replreport.txt 
repadmin /showrepl %1 /all >> replreport.txt 
echo. >> replreport.txt

echo Gathering Bridgeheads 
echo Bridgeheads >> replreport.txt 
echo. >> replreport.txt 
repadmin /bridgeheads %1 /verbose >> replreport.txt 
echo. >> replreport.txt

echo Gathering ISTG 
echo ISTG >> replreport.txt 
echo. >> replreport.txt 
repadmin /istg %1 >> replreport.txt 
echo. >> replreport.txt

echo Gathering DRS Calls 
echo Outbound DRS Calls >> replreport.txt 
echo. >> replreport.txt 
repadmin /showoutcalls %1 >> replreport.txt 
echo. >> replreport.txt

echo Gathering Queue 
echo Queue >> replreport.txt 
echo. >> replreport.txt 
repadmin /queue %1 >> replreport.txt 
echo. >> replreport.txt

echo Gathering KCC Failures 
echo KCC Failures >> replreport.txt 
echo. >> replreport.txt 
repadmin /failcache %1 >> replreport.txt 
echo. >> replreport.txt

echo Gathering Trusts 
echo Trusts >> replreport.txt 
echo. >> replreport.txt 
repadmin /showtrust %1 >> replreport.txt 
echo. >> replreport.txt

echo Gathering Replication Flags 
echo Replication Flags >> replreport.txt 
echo. >> replreport.txt 
repadmin /bind %1 >> replreport.txt 
echo. >> replreport.txt

echo Done.