@ECHO OFF
:: STARTUP CAS SERVICES


REM -Underhill_WO23235_2/23/2012- cscript.exe SvcManager.vbs Contosocorp lascsmt01 ContosoCollectionsConnectorService start 10
cscript.exe SvcManager.vbs Contosocorp lascsmt01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp lascasmt03 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp lascasmt02 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp lascasmt01 ContosoDataLayerService start 10

REM cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOn start 10
REM cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnGateway start 10
REM cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnQuery start 10
REM cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnReports start 10
REM cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnServiceMonitor start 10

REM **Removed due to CSTART decom 11/16/2011** cscript.exe SvcManager.vbs Contosocorp LASCSMT01 CreditOneFoxProSideCollectionsReplicationService start 10
REM **Removed due to CSTART decom 11/16/2011** cscript.exe SvcManager.vbs Contosocorp LASCSMT01 CreditOneOracleSideCollectionsReplicationService start 10

REM - Removed due to server decom 2/26/2012 cscript.exe SvcManager.vbs Contosocorp LASIBIZ03 "AIS iBizflow Server" start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CentralizedCacheService start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CreditOne.LogParser.Service start 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CollectionsAgentTimeService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT01 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT02 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT03 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT04 ContosoDataLayerService start 10
cscript.exe SvcManager.vbs Contosocorp LASMT05 ContosoDataLayerService start 10

REM Moved ContosoQueueProcessorService to the bottom of the list per RCABALLERO's request (WO 30517) 9/24/2012
cscript.exe SvcManager.vbs Contosocorp lassvc01 ContosoQueueProcessorService start 10
