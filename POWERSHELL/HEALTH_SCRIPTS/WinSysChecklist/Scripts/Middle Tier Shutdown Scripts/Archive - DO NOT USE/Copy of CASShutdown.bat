@ECHO OFF
:: SHUTS DOWN CAS SERVICES

cscript.exe SvcManager.vbs Contosocorp lassvc01 ContosoQueueProcessorService stop 10
cscript.exe SvcManager.vbs Contosocorp lascsmt01 ContosoCollectionsConnectorService stop 10
cscript.exe SvcManager.vbs Contosocorp lascsmt01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp lascasmt02 CacheDataManager stop 10
cscript.exe SvcManager.vbs Contosocorp lascasmt02 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp lascasmt01 CacheDataManager stop 10
cscript.exe SvcManager.vbs Contosocorp lascasmt01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOnServiceMonitor stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnServiceMonitor stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOnGateway stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnGateway stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOnQuery stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnQuery stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOnReports stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOnReports stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT01 WhosOn stop 10
cscript.exe SvcManager.vbs Contosocorp LASCHAT02 WhosOn stop 10

REM **Removed due to CSTART decom 11/16/2011** cscript.exe SvcManager.vbs Contosocorp LASCSMT01 CreditOneFoxProSideCollectionsReplicationService stop 10
REM **Removed due to CSTART decom 11/16/2011** cscript.exe SvcManager.vbs Contosocorp LASCSMT01 CreditOneOracleSideCollectionsReplicationService stop 10

cscript.exe SvcManager.vbs Contosocorp LASIBIZ03 "AIS iBizflow Server" stop 10
cscript.exe SvcManager.vbs Contosocorp LASIBIZ04 "AIS iBizflow Server" stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CentralizedCacheService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CreditOne.LogParser.Service stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 CollectionsAgentTimeService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT02 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT03 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT04 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT05 ContosoDataLayerService stop 10






