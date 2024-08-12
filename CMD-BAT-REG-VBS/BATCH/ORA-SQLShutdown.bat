@ECHO OFF
:: SHUTS DOWN ORACLE on LASDAWH1--RUN BEFORE REBOOT OF SERVER AND SQL on LASETL01
::ORA
cscript.exe SvcManager.vbs Contosocorp LASDAWH1 Oracleagent10gAgent stop 30
cscript.exe SvcManager.vbs Contosocorp LASDAWH1 OracleOraHome92050TNSListener stop 30
cscript.exe SvcManager.vbs Contosocorp LASDAWH1 OracleServicedw stop 30
::SQL
cscript.exe SvcManager.vbs Contosocorp LASETL01 MSSQLSERVER stop 30
cscript.exe SvcManager.vbs Contosocorp LASETL01 SQLSERVERAGENT stop 30