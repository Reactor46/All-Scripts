@echo off
SET SERVER=SERVERNAME\SQLEXPRESSOME
SET USR=test_user
SET PWD=test_password
SET LOG=_Load Tables Log.txt

Echo Loading Tables > "%LOG%"
Echo Time:%time% >> "%LOG%"
Echo Date: %date% >> "%LOG%"
Echo "" >> "%LOG%"

FOR /f %%i IN ('DIR *_Data.sql /B') do call :RunScript %%i
GOTO :END

:RunScript

Echo Executing Script: %1 >> "%LOG%"

SQLCMD -S %SERVER% -E -i %1 >> "%LOG%"
REM SQLCMD -S %SERVER% -U %USR% -P %PWD% -i %1 >> "%LOG%"

Echo Completed Script: %1 >> "%LOG%"
Echo "" >> "%LOG%"

:END