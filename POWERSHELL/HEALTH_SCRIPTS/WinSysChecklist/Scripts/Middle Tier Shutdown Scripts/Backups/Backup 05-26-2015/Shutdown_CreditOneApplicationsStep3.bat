@ECHO OFF

:: SHUT DOWN CreditOne Application IIS and Windows Services.


echo Do you really want to STOP services? (Y/N)
set INPUT=
set /P INPUT=Type input: %=%
If /I "%INPUT%"=="y" goto RUNSCRIPT 
If /I "%INPUT%"=="n" goto CANCELSCRIPT
echo Incorrect input & goto ENDScript

:RUNSCRIPT

cscript.exe SvcManager.vbs Contosocorp LASCASMT01 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT02 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT03 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT04 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT05 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT06 W3SVC stop 300
cscript.exe SvcManager.vbs Contosocorp LASCASMT07 W3SVC stop 300


cscript.exe SvcManager.vbs Contosocorp LASCASMT01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT02 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT03 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT04 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT05 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT06 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCASMT07 ContosoDataLayerService stop 10

cscript.exe SvcManager.vbs Contosocorp LASCOLL01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL02 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL03 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL04 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL05 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL06 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL07 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL08 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASCOLL09 ContosoDataLayerService stop 10

cscript.exe SvcManager.vbs Contosocorp LASMT01 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT02 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT03 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT04 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT07 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT08 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT09 ContosoDataLayerService stop 10
cscript.exe SvcManager.vbs Contosocorp LASMT10 ContosoDataLayerService stop 10

cscript.exe SvcManager.vbs Contosocorp LASSVC01 CentralizedCacheService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ContosoCheckRequestService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC01 ContosoLPSService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC02 CentralizedCacheService stop 10
cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoCheckRequestService stop 15
cscript.exe SvcManager.vbs Contosocorp LASSVC02 ContosoLPSService stop 15

cscript.exe SvcManager.vbs creditoneapp LASAUTH01.CREDITONEAPP.BIZ W3SVC stop 300
cscript.exe SvcManager.vbs creditoneapp LASAUTH02.CREDITONEAPP.BIZ W3SVC stop 300

goto ENDSCRIPT


:CANCELSCRIPT

echo You decided to exit script by entering %input%


:ENDScript


echo end of script
