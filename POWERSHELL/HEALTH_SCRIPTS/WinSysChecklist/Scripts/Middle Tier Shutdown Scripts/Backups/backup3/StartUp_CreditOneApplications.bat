@ECHO OFF
:: STARTS UP CreditOne Application IIS and Windows Services.

cscript.exe SvcManager.vbs creditoneapp LASAUTH01.CREDITONEAPP.BIZ W3SVC start 300
cscript.exe SvcManager.vbs creditoneapp LASAUTH02.CREDITONEAPP.BIZ W3SVC start 300

cscript.exe SvcManager.vbs Contosocorp LASSVC01 CentralizedCacheService start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ContosoCheckRequestService start 15
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ContosoLPSService start 15
cscript.exe SvcManager.vbs Contosocorp LASSVC02 CentralizedCacheService start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoCheckRequestService start 15
cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoLPSService start 15

cscript.exe SvcManager.vbs Contosocorp LASCASMT01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT02 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT03 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT04 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT05 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT06 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT07 ContosoDataLayerService start 10

cscript.exe SvcManager.vbs Contosocorp LASCOLL01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL05 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL06 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL07 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL08 ContosoDataLayerService start 10

cscript.exe SvcManager.vbs Contosocorp LASMT01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT02 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT03 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT04 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT05 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT06 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT07 ContosoDataLayerService start 10

cscript.exe SvcManager.vbs Contosocorp LASCASMT01 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT02 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT03 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT04 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT05 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT06 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT07 W3SVC start 300

cscript.exe SvcManager.vbs Contosocorp LASMERIT01 CreditEngine start 15
cscript.exe SvcManager.vbs Contosocorp LASMERIT02 CreditEngine start 15

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 CreditPullService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoIPFraudCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 BoardingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoIdentityCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoApplicationProcessingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 NetConnectTransactionsSavingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoDebitCardHolderFileWatcher start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 FromPPSExchangeFileWatcherService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoApplicationImportService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoApplicationParsingService start 15

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 CreditPullService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 ContosoIPFraudCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 ContosoIdentityCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 ContosoApplicationProcessingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 NetConnectTransactionsSavingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 ContosoApplicationImportService start 15

cscript.exe SvcManager.vbs Contosocorp LASSVC01 CollectionsAgentTimeService start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CreditOne.LogParser.Service start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CreditOneBatchLetterRequestService start 15
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ContosoQueueProcessorService start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 CollectionsAgentTimeService start 10
::cscript.exe SvcManager.vbs Contosocorp LASSVC02 CreditOneBatchLetterRequestService start 15
::cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoQueueProcessorService start 10
::cscript.exe SvcManager.vbs Contosocorp LASSVC02 CreditOne.LogParser.Service start 10

cscript.exe SvcManager.vbs Contosocorp LASCAPS01 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCAPS02 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCAPS03 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCAPS04 W3SVC start 300

cscript.exe SvcManager.vbs Contosocorp LASCOLL01 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL05 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL06 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL07 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL08 W3SVC start 300

cscript.exe SvcManager.vbs Contosocorp LASCAS01 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCAS02 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCAS03 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCAS04 W3SVC start 300
cscript.exe SvcManager.vbs Contosocorp LASCAS05 W3SVC start 300