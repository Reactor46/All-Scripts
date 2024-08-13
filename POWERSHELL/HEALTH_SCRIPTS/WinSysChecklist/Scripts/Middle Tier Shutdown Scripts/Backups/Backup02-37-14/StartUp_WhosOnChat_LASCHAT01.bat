@ECHO OFF
:: STARTS UP Whos On Chat Application services on LASCHAT01.

cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOn start 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOnGateway start 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOnQuery start 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOnReports start 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOnServiceMonitor start 10