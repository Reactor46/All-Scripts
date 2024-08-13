@ECHO OFF
:: SHUTS DOWN Whos On Chat Application services.

cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOnServiceMonitor stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnServiceMonitor stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOnGateway stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnGateway stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOnQuery stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnQuery stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOnReports stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnReports stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOn stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOn stop 10