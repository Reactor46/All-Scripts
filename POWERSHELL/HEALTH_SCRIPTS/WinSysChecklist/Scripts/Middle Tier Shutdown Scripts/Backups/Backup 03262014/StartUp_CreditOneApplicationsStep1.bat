
REM :: SHUT DOWN CreditOne Application IIS and Windows Services.


REM echo Do you really want to startup services? (Y/N)
REM set INPUT=
REM set /P INPUT=Type input: %=%
REM If /I "%INPUT%"=="y" goto RUNSCRIPT 
REM If /I "%INPUT%"=="n" goto DONOTSHUTDOWN
REM echo Incorrect input & goto ENDScript

:RUNSCRIPT
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

goto ENDSCRIPT

:DONOTSHUTDOWN

echo You decided to exit script by entering %input%

:ENDScript

echo end of script
