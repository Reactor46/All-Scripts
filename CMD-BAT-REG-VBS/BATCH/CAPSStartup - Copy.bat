@ECHO OFF
:: STARTSUP CAPS SERVICES


cscript.exe SvcManager.vbs Contosocorp lassvc01 ContosoFinCenService start 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 ContosoCheckRequestService start 15
cscript.exe SvcManager.vbs Contosocorp lasmerit01 CreditEngine start 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 NetConnectTransactionsSavingService start 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 ContosoLPSService start 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 CreditOneBatchLetterRequestService start 15
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
cscript.exe SvcManager.vbs Contosocorp lassvc01 CreditOne.AccountTakeOverFileWatcher start 15





















