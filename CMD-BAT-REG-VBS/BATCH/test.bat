@echo off
cls
set /p n=Source_server:
set /p h=Destination_server:
set /p m=database:
sqlcmd -S %n% -E -i"C:\Documents and Settings\MayurS2\Desktop\sd\query1.sql" -v db=%m% s2=%h% -o "mydata.csv" -W -w 999 -s"," 
call mydata.csv
exit
