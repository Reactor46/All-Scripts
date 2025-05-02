
:: START UP CreditOne Application IIS and Windows Services.


echo Do you really want to START services? (Y/N)
set INPUT=
set /P INPUT=Type input: %=%
If /I "%INPUT%"=="y" goto RUNSCRIPT 
If /I "%INPUT%"=="n" goto CANCELSCRIPT
echo Incorrect input & goto ENDScript

:RUNSCRIPT
@ECHO OFF
:: STARTS UP CreditOne Application IIS and Windows Services.

cscript.exe SvcManager.vbs Contosocorp LASMERIT01 CreditEngine start 15
cscript.exe SvcManager.vbs Contosocorp LASMERIT02 CreditEngine start 15

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 CreditPullService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoIPFraudCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoIdentityCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoApplicationProcessingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoApplicationImportService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 FromPPSExchangeFileWatcherService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoDebitCardHolderFileWatcher start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoApplicationParsingService start 15

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 CreditPullService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoIPFraudCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoIdentityCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoApplicationProcessingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoDebitCardHolderFileWatcher start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 FromPPSExchangeFileWatcherService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoApplicationImportService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoApplicationParsingService start 15

:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 CreditPullService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 ContosoIPFraudCheckService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 ContosoIdentityCheckService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 ContosoApplicationProcessingService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 ContosoApplicationImportService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 FromPPSExchangeFileWatcherService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 ContosoDebitCardHolderFileWatcher start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 ContosoApplicationParsingService start 15

:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 CreditPullService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 ContosoIPFraudCheckService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 ContosoIdentityCheckService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 ContosoApplicationProcessingService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 ContosoApplicationImportService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 FromPPSExchangeFileWatcherService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 ContosoDebitCardHolderFileWatcher start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 ContosoApplicationParsingService start 15

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 CreditPullService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 ContosoIPFraudCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 ContosoIdentityCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 ContosoApplicationProcessingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 ContosoApplicationImportService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 FromPPSExchangeFileWatcherService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 ContosoDebitCardHolderFileWatcher start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 ContosoApplicationParsingService start 15

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 CreditPullService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 ContosoIPFraudCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 ContosoIdentityCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 ContosoApplicationProcessingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 ContosoApplicationImportService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 FromPPSExchangeFileWatcherService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 ContosoDebitCardHolderFileWatcher start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 ContosoApplicationParsingService start 15

cscript.exe SvcManager.vbs Contosocorp LASSVC03 CollectionsAgentTimeService start 10

cscript.exe SvcManager.vbs Contosocorp LASSVC04 CollectionsAgentTimeService start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC04 CreditOneBatchLetterRequestService start 15
cscript.exe SvcManager.vbs Contosocorp LASSVC04 CreditOne.LogParser.Service start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC04 ContosoQueueProcessorService start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC04 FdrOutGoingFileWatcher start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC04 ContosoCentralizedCacheService start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC04 ValidationTriggerWatcher start 10

goto ENDSCRIPT

:CANCELSCRIPT

echo You decided to exit script by entering %input%

:ENDScript

echo end of script
