@ECHO OFF
:: STARTUP ORACLE on LASDAWH1--RUN BEFORE REBOOT OF SERVER AND SQL on LASETL01
::ORA
cscript.exe SvcManager.vbs Contosocorp LASDAWH1 Oracleagent10gAgent start 30
cscript.exe SvcManager.vbs Contosocorp LASDAWH1 OracleOraHome92050TNSListener start 30
cscript.exe SvcManager.vbs Contosocorp LASDAWH1 OracleServicedw start 30
::SQL
cscript.exe SvcManager.vbs Contosocorp LASETL01 MSSQLSERVER start 30
cscript.exe SvcManager.vbs Contosocorp LASETL01 SQLSERVERAGENT start 30