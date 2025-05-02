@ECHO OFF
:: STARTSUP CAPS SERVICES


cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationProcessingService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 BoardingService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoDebitCardHolderFileWatcher start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoIdentityCheckService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 FromPPSExchangeFileWatcherService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoIPFraudCheckService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 CreditPullService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationProcessingService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationImportService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationParsingService start 15





















