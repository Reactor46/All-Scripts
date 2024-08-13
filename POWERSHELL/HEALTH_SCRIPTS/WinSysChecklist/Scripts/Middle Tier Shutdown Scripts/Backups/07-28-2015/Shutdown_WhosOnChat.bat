@ECHO OFF
:: SHUTS DOWN Whos On Chat Application services.


echo Do you really want to STOP services? (Y/N)
set INPUT=
set /P INPUT=Type input: %=%
If /I "%INPUT%"=="y" goto RUNSCRIPT 
If /I "%INPUT%"=="n" goto CANCELSCRIPT
echo Incorrect input & goto ENDScript

:RUNSCRIPT

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

goto ENDSCRIPT


:CANCELSCRIPT

echo You decided to exit script by entering %input%


:ENDScript


echo end of script
