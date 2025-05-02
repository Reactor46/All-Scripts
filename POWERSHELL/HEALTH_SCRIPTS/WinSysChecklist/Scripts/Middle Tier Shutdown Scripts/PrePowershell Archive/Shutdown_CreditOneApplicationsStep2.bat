@ECHO OFF

:: SHUT DOWN CreditOne Application IIS and Windows Services.


echo Do you really want to STOP services? (Y/N)
set INPUT=
set /P INPUT=Type input: %=%
If /I "%INPUT%"=="y" goto RUNSCRIPT 
If /I "%INPUT%"=="n" goto CANCELSCRIPT
echo Incorrect input & goto ENDScript

:RUNSCRIPT

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoApplicationParsingService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoApplicationImportService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoDebitCardHolderFileWatcher stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 FromPPSExchangeFileWatcherService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoIdentityCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoApplicationProcessingService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 ContosoIPFraudCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT01 CreditPullService stop 15

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoApplicationParsingService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoApplicationImportService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoDebitCardHolderFileWatcher stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 FromPPSExchangeFileWatcherService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoIdentityCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoApplicationProcessingService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 ContosoIPFraudCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT02 CreditPullService stop 15

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 ContosoApplicationParsingService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 ContosoApplicationImportService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 ContosoDebitCardHolderFileWatcher stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 FromPPSExchangeFileWatcherService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 ContosoIdentityCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 ContosoApplicationProcessingService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 ContosoIPFraudCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT05 CreditPullService stop 15

cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 ContosoApplicationParsingService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 ContosoApplicationImportService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 ContosoDebitCardHolderFileWatcher stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 FromPPSExchangeFileWatcherService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 ContosoIdentityCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 ContosoApplicationProcessingService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 ContosoIPFraudCheckService stop 15
cscript.exe SvcManager.vbs Contosocorp LASCAPSMT06 CreditPullService stop 15

cscript.exe SvcManager.vbs Contosocorp LASMERIT01 CreditEngine stop 15
cscript.exe SvcManager.vbs Contosocorp LASMERIT02 CreditEngine stop 15

cscript.exe SvcManager.vbs Contosocorp LASSVC01 CollectionsAgentTimeService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CreditOneBatchLetterRequestService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CreditOne.LogParser.Service stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ContosoQueueProcessorService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 FdrOutGoingFileWatcher stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ValidationTriggerWatcher stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ContosoFinCenService stop 15

cscript.exe SvcManager.vbs Contosocorp LASSVC02 CollectionsAgentTimeService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 CreditOneBatchLetterRequestService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC02 CreditOne.LogParser.Service stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoQueueProcessorService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 FdrOutGoingFileWatcher stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 ValidationTriggerWatcher stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoFinCenService stop 15

cscript.exe SvcManager.vbs Contosocorp LASSVC03 CollectionsAgentTimeService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC03 CreditOneBatchLetterRequestService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC03 CreditOne.LogParser.Service stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC03 ContosoQueueProcessorService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC03 FdrOutGoingFileWatcher stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC03 ValidationTriggerWatcher stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC03 ContosoFinCenService stop 15

cscript.exe SvcManager.vbs Contosocorp LASSVC04 CollectionsAgentTimeService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC04 CreditOneBatchLetterRequestService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC04 CreditOne.LogParser.Service stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC04 ContosoQueueProcessorService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC04 FdrOutGoingFileWatcher stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC04 ValidationTriggerWatcher stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC04 ContosoFinCenService stop 15

goto ENDSCRIPT


:CANCELSCRIPT

echo You decided to exit script by entering %input%


:ENDScript


echo end of script
