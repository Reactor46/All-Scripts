@ECHO OFF
:: SHUT DOWN CreditOne Application IIS and Windows Services.

cscript.exe SvcManager.vbs Contosocorp LASCAS03 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCAS04 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCAS05 W3SVC stop 300

cscript.exe SvcManager.vbs Contosocorp LASCOLL01 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL05 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL06 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL07 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCOLL08 W3SVC stop 300

cscript.exe SvcManager.vbs Contosocorp LASCAPS03 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCAPS04 W3SVC stop 300

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoApplicationParsingService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoApplicationImportService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 NetConnectTransactionsSavingService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoDebitCardHolderFileWatcher stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 FromPPSExchangeFileWatcherService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoIdentityCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 BoardingService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoApplicationProcessingService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoIPFraudCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 CreditPullService stop 15
cscript.exe SvcManager.vbs Contosocorp LASMERIT01 CreditEngine stop 15
cscript.exe SvcManager.vbs Contosocorp LASMERIT02 CreditEngine stop 15

cscript.exe SvcManager.vbs Contosocorp LASSVC01 CollectionsAgentTimeService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CreditOne.LogParser.Service stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CreditOneBatchLetterRequestService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ContosoFinCenService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ContosoQueueProcessorService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 CollectionsAgentTimeService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 CreditOne.LogParser.Service stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 CreditOneBatchLetterRequestService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoFinCenService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoQueueProcessorService stop 10

cscript.exe SvcManager.vbs Contosocorp LASCASMT01 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT02 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT03 W3SVC stop 300

cscript.exe SvcManager.vbs Contosocorp LASCSMT01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT02 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT03 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL05 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL06 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL07 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL08 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT02 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT03 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT04 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT05 ContosoDataLayerService stop 10

cscript.exe SvcManager.vbs Contosocorp LASSVC01 CentralizedCacheService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ContosoCheckRequestService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ContosoLPSService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC02 CentralizedCacheService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoCheckRequestService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoLPSService stop 15

cscript.exe SvcManager.vbs creditoneapp LASAUTH01.CREDITONEAPP.BIZ W3SVC stop 300
cscript.exe SvcManager.vbs creditoneapp LASAUTH02.CREDITONEAPP.BIZ W3SVC stop 300