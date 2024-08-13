@ECHO OFF
:: STARTS DATAGUARD TESTING SERVICES

cscript.exe SvcManager.vbs Contosocorp lascaps03 W3svc start 15
cscript.exe SvcManager.vbs Contosocorp lascaps04 W3svc start 15
cscript.exe SvcManager.vbs Contosocorp lascas04 W3svc start 15
cscript.exe SvcManager.vbs Contosocorp lascas05 W3svc start 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 W3svc start 15
cscript.exe SvcManager.vbs Contosocorp lascasmt01 CacheDataManager start 15
cscript.exe SvcManager.vbs Contosocorp lascasmt02 CacheDataManager start 15
cscript.exe SvcManager.vbs Contosocorp lascasmt01 ContosoDataLayerService start 15
cscript.exe SvcManager.vbs Contosocorp lascasmt02 ContosoDataLayerService start 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 CentralizedCacheService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 NetConnectTransactionsSavingService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 BoardingService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoDebitCardHolderFileWatcher start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoIdentityCheckService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 FromPPSExchangeFileWatcherService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoIPFraudCheckService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 CreditPullService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationProcessingService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationImportService start 15