@ECHO OFF
:: STARTS UP CreditOne Application IIS and Windows Services.

cscript.exe SvcManager.vbs creditoneapp LASAUTH01.CREDITONEAPP.BIZ W3SVC start 300
cscript.exe SvcManager.vbs creditoneapp LASAUTH02.CREDITONEAPP.BIZ W3SVC start 300

cscript.exe SvcManager.vbs Contosocorp LASSVC01 CentralizedCacheService start 10
cscript.exe SvcManager.vbs Contosocorp lassvc01 ContosoCheckRequestService start 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 ContosoLPSService start 15

cscript.exe SvcManager.vbs Contosocorp lascsmt01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp lascasmt01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp lascasmt02 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp lascasmt03 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL05 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL06 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT02 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT03 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT04 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT05 ContosoDataLayerService start 10

cscript.exe SvcManager.vbs Contosocorp LASCASMT01 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT02 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT03 W3SVC start 300

cscript.exe SvcManager.vbs Contosocorp lasmerit01 CreditEngine start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 CreditPullService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoIPFraudCheckService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 BoardingService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoIdentityCheckService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationProcessingService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 NetConnectTransactionsSavingService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoDebitCardHolderFileWatcher start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 FromPPSExchangeFileWatcherService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationImportService start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationParsingService start 15

cscript.exe SvcManager.vbs Contosocorp LASSVC01 CollectionsAgentTimeService start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CreditOne.LogParser.Service start 10
cscript.exe SvcManager.vbs Contosocorp lassvc01 CreditOneBatchLetterRequestService start 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 ContosoFinCenService start 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 ContosoQueueProcessorService start 10

cscript.exe SvcManager.vbs Contosocorp LASCAPS03 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCAPS04 W3SVC start 300

cscript.exe SvcManager.vbs Contosocorp LASCOLL01 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL05 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL06 W3SVC start 300

cscript.exe SvcManager.vbs Contosocorp LASCAS03 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCAS04 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCAS05 W3SVC start 300
