@ECHO OFF
:: STARTS UP Whos On Chat Application services on LASCHAT02.

cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOn start 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnGateway start 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnQuery start 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnReports start 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnServiceMonitor start 10