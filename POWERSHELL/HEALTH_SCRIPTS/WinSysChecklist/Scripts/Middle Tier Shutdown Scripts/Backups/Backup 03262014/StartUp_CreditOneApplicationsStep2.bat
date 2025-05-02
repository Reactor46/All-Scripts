
REM :: SHUT DOWN CreditOne Application IIS and Windows Services.


REM echo Do you really want to stop services? (Y/N)
REM set INPUT=
REM set /P INPUT=Type input: %=%
REM If /I "%INPUT%"=="y" goto RUNSCRIPT 
REM If /I "%INPUT%"=="n" goto DONOTSHUTDOWN
REM echo Incorrect input & goto ENDScript

:RUNSCRIPT
@ECHO OFF
:: STARTS UP CreditOne Application IIS and Windows Services.

cscript.exe SvcManager.vbs Contosocorp LASMERIT01 CreditEngine start 15
cscript.exe SvcManager.vbs Contosocorp LASMERIT02 CreditEngine start 15

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 CreditPullService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoIPFraudCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoIdentityCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoApplicationProcessingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 NetConnectTransactionsSavingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoApplicationImportService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 BoardingService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 FromPPSExchangeFileWatcherService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoDebitCardHolderFileWatcher start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoApplicationParsingService start 15


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
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 BoardingService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 FromPPSExchangeFileWatcherService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 ContosoDebitCardHolderFileWatcher start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT03 ContosoApplicationParsingService start 15

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 CreditPullService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 ContosoIPFraudCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 ContosoIdentityCheckService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 ContosoApplicationProcessingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 NetConnectTransactionsSavingService start 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 ContosoApplicationImportService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 BoardingService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 FromPPSExchangeFileWatcherService start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 ContosoDebitCardHolderFileWatcher start 15
:: cscript.exe SvcManager.vbs Contosocorp LASCAPSMT04 ContosoApplicationParsingService start 15

cscript.exe SvcManager.vbs Contosocorp LASSVC01 CollectionsAgentTimeService start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CreditOne.LogParser.Service start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CreditOneBatchLetterRequestService start 15
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ContosoQueueProcessorService start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 CollectionsAgentTimeService start 10
::cscript.exe SvcManager.vbs Contosocorp LASSVC02 CreditOneBatchLetterRequestService start 15
::cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoQueueProcessorService start 10
::cscript.exe SvcManager.vbs Contosocorp LASSVC02 CreditOne.LogParser.Service start 10

goto ENDSCRIPT

:DONOTSHUTDOWN

echo You decided to exit script by entering %input%

:ENDScript

echo end of script
