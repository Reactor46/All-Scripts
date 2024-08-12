@ECHO OFF
:: SHUTS DOWN CAPS SERVICES

cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationProcessingService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationParsingService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoApplicationImportService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 CreditPullService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoIPFraudCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 FromPPSExchangeFileWatcherService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoIdentityCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 ContosoDebitCardHolderFileWatcher stop 15
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 BoardingService stop 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 CreditOneBatchLetterRequestService stop 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 ContosoLPSService stop 30
cscript.exe SvcManager.vbs Contosocorp lascapsmt01 NetConnectTransactionsSavingService stop 15
cscript.exe SvcManager.vbs Contosocorp lasmerit01 CreditEngine stop 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 ContosoCheckRequestService stop 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 ContosoFinCenService stop 15
cscript.exe SvcManager.vbs Contosocorp lassvc01 CreditOne.AccountTakeOverFileWatcher stop 15

