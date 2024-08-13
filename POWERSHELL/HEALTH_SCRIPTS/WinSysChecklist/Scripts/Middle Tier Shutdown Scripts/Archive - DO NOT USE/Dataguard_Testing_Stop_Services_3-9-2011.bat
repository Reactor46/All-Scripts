@ECHO OFF
:: SHUTS DOWN DATAGUARD TESTING SERVICES

cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationProcessingService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationImportService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 CreditPullService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoIPFraudCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 FromPPSExchangeFileWatcherService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoIdentityCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoDebitCardHolderFileWatcher stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 BoardingService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 NetConnectTransactionsSavingService stop 15
cscript.exe SvcManager.vbs Contosocorp lascasmt01 CacheDataManager stop 15
cscript.exe SvcManager.vbs Contosocorp lascasmt02 CacheDataManager stop 15
cscript.exe SvcManager.vbs Contosocorp lascasmt01 ContosoDataLayerService stop 15
cscript.exe SvcManager.vbs Contosocorp lascasmt02 ContosoDataLayerService stop 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 CentralizedCacheService stop 15
cscript.exe SvcManager.vbs Contosocorp lascaps03 W3svc stop 15
cscript.exe SvcManager.vbs Contosocorp lascaps04 W3svc stop 15
cscript.exe SvcManager.vbs Contosocorp lascas04 W3svc stop 15
cscript.exe SvcManager.vbs Contosocorp lascas05 W3svc stop 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 W3svc stop 15